package main

import (
	"bytes"
	"crypto/tls"
	"encoding/json"
	"fmt"
	"io"
	"net"
	"net/http"
	"net/smtp"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	appTitle           = "outlook自动转发程序"
	configName         = "outlook-auto-forward-config.json"
	olFolderInbox      = 6
	olMailItem         = 43
	defaultPollSeconds = 60
	scanLimit          = 30

	idWebhookRadio = 101
	idWebhookEdit  = 102
	idSMTPRadio    = 201
	idSMTPServer   = 202
	idSMTPSSL      = 203
	idSMTPPort     = 204
	idSMTPUser     = 205
	idSMTPPass     = 206
	idStartButton  = 301
	idBackButton   = 302
	idPollSeconds  = 303
	idLogEdit      = 401

	timerPoll = 1
)

type appConfig struct {
	Mode        string     `json:"mode"`
	Webhook     string     `json:"webhook"`
	PollSeconds int        `json:"poll_seconds"`
	SMTP        smtpConfig `json:"smtp"`
}

type smtpConfig struct {
	Server   string `json:"server"`
	SSL      bool   `json:"ssl"`
	Port     string `json:"port"`
	Username string `json:"username"`
	Password string `json:"password"`
}

type appState struct {
	hwnd       hwnd
	font       uintptr
	configPath string
	config     appConfig
	controls   []hwnd
	monitor    []hwnd
	logEdit    hwnd
	client     *outlookClient
	seen       map[string]bool
	monitoring bool
}

type outlookClient struct {
	app       *ole.IDispatch
	namespace *ole.IDispatch
	inbox     *ole.IDispatch
	items     *ole.IDispatch
}

type mailInfo struct {
	EntryID      string
	Subject      string
	SenderName   string
	SenderEmail  string
	ReceivedTime string
	Body         string
}

type feishuTextMessage struct {
	MsgType string `json:"msg_type"`
	Content struct {
		Text string `json:"text"`
	} `json:"content"`
}

type hwnd uintptr

var (
	user32   = syscall.NewLazyDLL("user32.dll")
	kernel32 = syscall.NewLazyDLL("kernel32.dll")

	procCreateWindowExW  = user32.NewProc("CreateWindowExW")
	procDefWindowProcW   = user32.NewProc("DefWindowProcW")
	procDestroyWindow    = user32.NewProc("DestroyWindow")
	procDispatchMessageW = user32.NewProc("DispatchMessageW")
	procGetMessageW      = user32.NewProc("GetMessageW")
	procGetWindowTextW   = user32.NewProc("GetWindowTextW")
	procLoadCursorW      = user32.NewProc("LoadCursorW")
	procPostQuitMessage  = user32.NewProc("PostQuitMessage")
	procRegisterClassExW = user32.NewProc("RegisterClassExW")
	procSendMessageW     = user32.NewProc("SendMessageW")
	procSetTimer         = user32.NewProc("SetTimer")
	procKillTimer        = user32.NewProc("KillTimer")
	procSetWindowTextW   = user32.NewProc("SetWindowTextW")
	procShowWindow       = user32.NewProc("ShowWindow")
	procTranslateMessage = user32.NewProc("TranslateMessage")
	procUpdateWindow     = user32.NewProc("UpdateWindow")
	procEnableWindow     = user32.NewProc("EnableWindow")
	procGetModuleHandleW = kernel32.NewProc("GetModuleHandleW")
	gdi32                = syscall.NewLazyDLL("gdi32.dll")
	procCreateFontW      = gdi32.NewProc("CreateFontW")
	procDeleteObject     = gdi32.NewProc("DeleteObject")
	procGetStockObject   = gdi32.NewProc("GetStockObject")
	currentState         *appState
)

func main() {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	if err := ole.CoInitialize(0); err != nil {
		messageBox("初始化 Windows COM 失败：" + err.Error())
		return
	}
	defer ole.CoUninitialize()

	state := &appState{
		configPath: filepath.Join(appDir(), configName),
		config:     defaultConfig(),
		seen:       make(map[string]bool),
	}
	state.loadConfig()
	currentState = state

	if err := runGUI(state); err != nil {
		messageBox(err.Error())
	}
}

func defaultConfig() appConfig {
	return appConfig{
		Mode:        "webhook",
		PollSeconds: defaultPollSeconds,
		SMTP: smtpConfig{
			SSL:  true,
			Port: "465",
		},
	}
}

func runGUI(state *appState) error {
	instance, _, _ := procGetModuleHandleW.Call(0)
	className := utf16Ptr("OutlookAutoForwardWindow")
	cursor, _, _ := procLoadCursorW.Call(0, uintptr(32512))
	brush, _, _ := procGetStockObject.Call(0)

	wc := wndclassex{
		cbSize:        uint32(unsafe.Sizeof(wndclassex{})),
		lpfnWndProc:   syscall.NewCallback(wndProc),
		hInstance:     instance,
		hCursor:       cursor,
		hbrBackground: brush,
		lpszClassName: uintptr(unsafe.Pointer(className)),
	}
	if r, _, err := procRegisterClassExW.Call(uintptr(unsafe.Pointer(&wc))); r == 0 {
		return fmt.Errorf("注册窗口失败：%v", err)
	}

	main, _, err := procCreateWindowExW.Call(
		0,
		uintptr(unsafe.Pointer(className)),
		uintptr(unsafe.Pointer(utf16Ptr(appTitle))),
		wsOverlappedWindow,
		200, 120, 780, 580,
		0, 0, instance, 0,
	)
	if main == 0 {
		return fmt.Errorf("创建窗口失败：%v", err)
	}

	state.hwnd = hwnd(main)
	state.font = createUIFont()
	buildConfigView(state)
	procShowWindow.Call(main, 1)
	procUpdateWindow.Call(main)

	var msg msg
	for {
		ret, _, _ := procGetMessageW.Call(uintptr(unsafe.Pointer(&msg)), 0, 0, 0)
		if int32(ret) <= 0 {
			break
		}
		procTranslateMessage.Call(uintptr(unsafe.Pointer(&msg)))
		procDispatchMessageW.Call(uintptr(unsafe.Pointer(&msg)))
	}
	return nil
}

func buildConfigView(state *appState) {
	parent := state.hwnd
	state.controls = append(state.controls,
		createGroup(parent, "功能1：webhook消息提示", 24, 22, 708, 96),
		createRadio(parent, "使用 webhook 推送", idWebhookRadio, 46, 54, 180, 24),
		createStatic(parent, "Webhook API", 250, 56, 120, 24),
		createEdit(parent, idWebhookEdit, state.config.Webhook, 360, 52, 330, 30, false),
		createGroup(parent, "功能2：smtp消息提示", 24, 138, 708, 224),
		createRadio(parent, "使用 SMTP 推送", idSMTPRadio, 46, 172, 180, 24),
		createStatic(parent, "SMTP服务器", 250, 166, 120, 24),
		createEdit(parent, idSMTPServer, state.config.SMTP.Server, 360, 162, 220, 30, false),
		createCheck(parent, "SSL", idSMTPSSL, 604, 164, 80, 26),
		createStatic(parent, "端口", 250, 206, 120, 24),
		createEdit(parent, idSMTPPort, state.config.SMTP.Port, 360, 202, 90, 30, false),
		createStatic(parent, "用户名/收件邮箱", 250, 246, 120, 24),
		createEdit(parent, idSMTPUser, state.config.SMTP.Username, 360, 242, 220, 30, false),
		createStatic(parent, "密码", 250, 286, 120, 24),
		createEdit(parent, idSMTPPass, state.config.SMTP.Password, 360, 282, 220, 30, true),
		createStatic(parent, "SMTP 会发送到用户名填写的邮箱。", 360, 322, 260, 24),
		createGroup(parent, "监听设置", 24, 382, 708, 64),
		createStatic(parent, "轮询间隔（秒）", 46, 410, 130, 24),
		createEdit(parent, idPollSeconds, strconv.Itoa(state.config.PollSeconds), 174, 406, 90, 30, false),
		createButton(parent, "开始监听", idStartButton, 24, 470, 136, 36),
	)

	if state.config.Mode == "smtp" {
		check(idSMTPRadio, true)
		check(idWebhookRadio, false)
	} else {
		check(idWebhookRadio, true)
		check(idSMTPRadio, false)
	}
	check(idSMTPSSL, state.config.SMTP.SSL)
}

func buildMonitorView(state *appState) {
	for _, ctl := range state.controls {
		show(ctl, false)
	}
	state.monitor = append(state.monitor,
		createStatic(state.hwnd, fmt.Sprintf("正在监听 Outlook 收件箱，轮询间隔 %d 秒。", state.config.PollSeconds), 24, 24, 460, 26),
		createButton(state.hwnd, "返回配置", idBackButton, 610, 22, 120, 32),
		createEdit(state.hwnd, idLogEdit, "", 24, 68, 706, 436, false),
	)
	state.logEdit = find(idLogEdit)
	appendLog("程序已启动，正在等待 Outlook 可用。")
}

func wndProc(window uintptr, message uint32, wParam, lParam uintptr) uintptr {
	switch message {
	case wmCommand:
		id := int(wParam & 0xffff)
		switch id {
		case idStartButton:
			currentState.startMonitoring()
		case idBackButton:
			currentState.stopMonitoring()
		case idWebhookRadio:
			check(idWebhookRadio, true)
			check(idSMTPRadio, false)
		case idSMTPRadio:
			check(idSMTPRadio, true)
			check(idWebhookRadio, false)
		}
	case wmTimer:
		if wParam == timerPoll {
			currentState.poll()
		}
	case wmClose:
		procDestroyWindow.Call(window)
	case wmDestroy:
		currentState.cleanup()
		procPostQuitMessage.Call(0)
	default:
		ret, _, _ := procDefWindowProcW.Call(window, uintptr(message), wParam, lParam)
		return ret
	}
	return 0
}

func (s *appState) startMonitoring() {
	cfg := readConfigFromUI()
	if err := validateConfig(cfg); err != nil {
		messageBox(err.Error())
		return
	}
	s.config = cfg
	if err := s.saveConfig(); err != nil {
		messageBox("保存配置失败：" + err.Error())
		return
	}

	procEnableWindow.Call(uintptr(find(idStartButton)), 0)
	s.monitoring = true
	buildMonitorView(s)
	procSetTimer.Call(uintptr(s.hwnd), timerPoll, uintptr(s.config.PollSeconds*1000), 0)
	s.poll()
}

func (s *appState) stopMonitoring() {
	procKillTimer.Call(uintptr(s.hwnd), timerPoll)
	if s.client != nil {
		s.client.Release()
		s.client = nil
	}
	s.monitoring = false
	for _, ctl := range s.monitor {
		procDestroyWindow.Call(uintptr(ctl))
	}
	s.monitor = nil
	s.logEdit = 0
	for _, ctl := range s.controls {
		show(ctl, true)
	}
	procEnableWindow.Call(uintptr(find(idStartButton)), 1)
}

func (s *appState) cleanup() {
	procKillTimer.Call(uintptr(s.hwnd), timerPoll)
	if s.client != nil {
		s.client.Release()
		s.client = nil
	}
	if s.font != 0 {
		procDeleteObject.Call(s.font)
		s.font = 0
	}
}

func (s *appState) poll() {
	if !s.monitoring {
		return
	}
	if s.client == nil {
		client, seen, err := connectAndSnapshot()
		if err != nil {
			appendLog("Outlook 暂不可用，15 秒后重试。原因：" + err.Error())
			return
		}
		s.client = client
		s.seen = seen
		appendLog(fmt.Sprintf("已连接 Outlook，正在监视收件箱。启动时已有邮件已忽略：%d", len(seen)))
		return
	}

	appendLog("正在监视 Outlook 收件箱...")
	mails, err := s.client.latestMails(scanLimit)
	if err != nil {
		appendLog("读取 Outlook 收件箱失败，将重新连接。原因：" + err.Error())
		s.client.Release()
		s.client = nil
		return
	}

	sort.Slice(mails, func(i, j int) bool {
		return mails[i].ReceivedTime < mails[j].ReceivedTime
	})

	for _, mail := range mails {
		if mail.EntryID == "" || s.seen[mail.EntryID] {
			continue
		}
		s.seen[mail.EntryID] = true

		if err := sendAlert(s.config, mail); err != nil {
			appendLog(fmt.Sprintf("发送失败：%s。原因：%v", emptyAsUnknown(mail.Subject), err))
			continue
		}
		appendLog("已发送：" + emptyAsUnknown(mail.Subject))
	}
}

func sendAlert(cfg appConfig, mail mailInfo) error {
	if cfg.Mode == "smtp" {
		return sendSMTPAlert(cfg.SMTP, mail)
	}
	return sendFeishuAlert(cfg.Webhook, mail)
}

func validateConfig(cfg appConfig) error {
	if cfg.PollSeconds < 5 {
		return fmt.Errorf("轮询间隔不能小于 5 秒")
	}
	if cfg.Mode == "smtp" {
		if cfg.SMTP.Server == "" || cfg.SMTP.Port == "" || cfg.SMTP.Username == "" || cfg.SMTP.Password == "" {
			return fmt.Errorf("请填写完整 SMTP 服务器、端口、用户名/收件邮箱和密码")
		}
		if _, err := strconv.Atoi(cfg.SMTP.Port); err != nil {
			return fmt.Errorf("SMTP 端口必须是数字")
		}
		return nil
	}
	if cfg.Webhook == "" {
		return fmt.Errorf("请填写 webhook api")
	}
	return nil
}

func readConfigFromUI() appConfig {
	cfg := defaultConfig()
	if isChecked(idSMTPRadio) {
		cfg.Mode = "smtp"
	} else {
		cfg.Mode = "webhook"
	}
	cfg.Webhook = text(idWebhookEdit)
	cfg.PollSeconds, _ = strconv.Atoi(text(idPollSeconds))
	cfg.SMTP.Server = text(idSMTPServer)
	cfg.SMTP.SSL = isChecked(idSMTPSSL)
	cfg.SMTP.Port = text(idSMTPPort)
	cfg.SMTP.Username = text(idSMTPUser)
	cfg.SMTP.Password = text(idSMTPPass)
	return cfg
}

func (s *appState) loadConfig() {
	data, err := os.ReadFile(s.configPath)
	if err != nil {
		return
	}
	var cfg appConfig
	if err := json.Unmarshal(data, &cfg); err != nil {
		return
	}
	if cfg.Mode == "" {
		cfg.Mode = "webhook"
	}
	if cfg.PollSeconds == 0 {
		cfg.PollSeconds = defaultPollSeconds
	}
	if cfg.SMTP.Port == "" {
		cfg.SMTP.Port = "465"
	}
	s.config = cfg
}

func (s *appState) saveConfig() error {
	data, err := json.MarshalIndent(s.config, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(s.configPath, data, 0600)
}

func connectAndSnapshot() (*outlookClient, map[string]bool, error) {
	client, err := newOutlookClient()
	if err != nil {
		return nil, nil, fmt.Errorf("连接本机 Outlook 失败，请确认 Outlook 已安装、已登录且允许外部程序访问：%w", err)
	}

	seen, err := client.snapshotSeen()
	if err != nil {
		client.Release()
		return nil, nil, fmt.Errorf("读取 Outlook 收件箱失败：%w", err)
	}

	return client, seen, nil
}

func newOutlookClient() (*outlookClient, error) {
	unknown, err := oleutil.GetActiveObject("Outlook.Application")
	if err != nil {
		unknown, err = oleutil.CreateObject("Outlook.Application")
		if err != nil {
			return nil, fmt.Errorf("make sure Outlook is installed and signed in on this computer: %w", err)
		}
	}

	app, err := unknown.QueryInterface(ole.IID_IDispatch)
	unknown.Release()
	if err != nil {
		return nil, err
	}

	namespaceVariant, err := oleutil.CallMethod(app, "GetNamespace", "MAPI")
	if err != nil {
		app.Release()
		return nil, err
	}
	namespace := namespaceVariant.ToIDispatch()

	inboxVariant, err := oleutil.CallMethod(namespace, "GetDefaultFolder", olFolderInbox)
	if err != nil {
		namespace.Release()
		app.Release()
		return nil, err
	}
	inbox := inboxVariant.ToIDispatch()

	itemsVariant, err := oleutil.GetProperty(inbox, "Items")
	if err != nil {
		inbox.Release()
		namespace.Release()
		app.Release()
		return nil, err
	}
	items := itemsVariant.ToIDispatch()

	if _, err := oleutil.CallMethod(items, "Sort", "[ReceivedTime]", true); err != nil {
		items.Release()
		inbox.Release()
		namespace.Release()
		app.Release()
		return nil, err
	}

	return &outlookClient{app: app, namespace: namespace, inbox: inbox, items: items}, nil
}

func (c *outlookClient) Release() {
	if c.items != nil {
		c.items.Release()
	}
	if c.inbox != nil {
		c.inbox.Release()
	}
	if c.namespace != nil {
		c.namespace.Release()
	}
	if c.app != nil {
		c.app.Release()
	}
}

func (c *outlookClient) snapshotSeen() (map[string]bool, error) {
	mails, err := c.latestMails(scanLimit)
	if err != nil {
		return nil, err
	}

	seen := make(map[string]bool, len(mails))
	for _, mail := range mails {
		if mail.EntryID != "" {
			seen[mail.EntryID] = true
		}
	}
	return seen, nil
}

func (c *outlookClient) latestMails(limit int) ([]mailInfo, error) {
	countVariant, err := oleutil.GetProperty(c.items, "Count")
	if err != nil {
		return nil, err
	}
	count := int(countVariant.Val)
	if count < limit {
		limit = count
	}

	mails := make([]mailInfo, 0, limit)
	for i := 1; i <= limit; i++ {
		itemVariant, err := oleutil.CallMethod(c.items, "Item", i)
		if err != nil {
			continue
		}
		item := itemVariant.ToIDispatch()
		if item == nil {
			continue
		}

		isMail, err := isMailItem(item)
		if err != nil || !isMail {
			item.Release()
			continue
		}

		mails = append(mails, mailInfo{
			EntryID:      getStringProperty(item, "EntryID"),
			Subject:      getStringProperty(item, "Subject"),
			SenderName:   getStringProperty(item, "SenderName"),
			SenderEmail:  getStringProperty(item, "SenderEmailAddress"),
			ReceivedTime: getStringProperty(item, "ReceivedTime"),
			Body:         bodySummary(getStringProperty(item, "Body")),
		})
		item.Release()
	}
	return mails, nil
}

func isMailItem(item *ole.IDispatch) (bool, error) {
	classVariant, err := oleutil.GetProperty(item, "Class")
	if err != nil {
		return false, err
	}
	return int(classVariant.Val) == olMailItem, nil
}

func getStringProperty(dispatch *ole.IDispatch, name string) string {
	value, err := oleutil.GetProperty(dispatch, name)
	if err != nil {
		return ""
	}
	return strings.TrimSpace(value.ToString())
}

func sendFeishuAlert(webhook string, mail mailInfo) error {
	msg := feishuTextMessage{MsgType: "text"}
	msg.Content.Text = alertText(mail)

	body, err := json.Marshal(msg)
	if err != nil {
		return err
	}

	req, err := http.NewRequest(http.MethodPost, webhook, bytes.NewReader(body))
	if err != nil {
		return err
	}
	req.Header.Set("Content-Type", "application/json; charset=utf-8")

	client := http.Client{Timeout: 10 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	respBody, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return fmt.Errorf("HTTP %s: %s", resp.Status, strings.TrimSpace(string(respBody)))
	}

	var result struct {
		Code int    `json:"code"`
		Msg  string `json:"msg"`
	}
	if err := json.Unmarshal(respBody, &result); err == nil && result.Code != 0 {
		if result.Msg == "" {
			result.Msg = strings.TrimSpace(string(respBody))
		}
		return fmt.Errorf("飞书返回 code=%d, msg=%s", result.Code, result.Msg)
	}
	return nil
}

func sendSMTPAlert(cfg smtpConfig, mail mailInfo) error {
	addr := net.JoinHostPort(cfg.Server, cfg.Port)
	from := cfg.Username
	to := cfg.Username
	message := []byte(strings.Join([]string{
		"From: " + from,
		"To: " + to,
		"Subject: " + encodeSubject("Outlook新邮件："+emptyAsUnknown(mail.Subject)),
		"MIME-Version: 1.0",
		"Content-Type: text/plain; charset=UTF-8",
		"",
		alertText(mail),
	}, "\r\n"))

	if cfg.SSL {
		conn, err := tls.DialWithDialer(&net.Dialer{Timeout: 10 * time.Second}, "tcp", addr, &tls.Config{ServerName: cfg.Server})
		if err != nil {
			return err
		}
		defer conn.Close()
		client, err := smtp.NewClient(conn, cfg.Server)
		if err != nil {
			return err
		}
		defer client.Quit()
		return smtpSend(client, cfg, from, to, message)
	}

	auth := smtp.PlainAuth("", cfg.Username, cfg.Password, cfg.Server)
	return smtp.SendMail(addr, auth, from, []string{to}, message)
}

func smtpSend(client *smtp.Client, cfg smtpConfig, from, to string, message []byte) error {
	auth := smtp.PlainAuth("", cfg.Username, cfg.Password, cfg.Server)
	if err := client.Auth(auth); err != nil {
		return err
	}
	if err := client.Mail(from); err != nil {
		return err
	}
	if err := client.Rcpt(to); err != nil {
		return err
	}
	writer, err := client.Data()
	if err != nil {
		return err
	}
	if _, err := writer.Write(message); err != nil {
		writer.Close()
		return err
	}
	return writer.Close()
}

func alertText(mail mailInfo) string {
	return fmt.Sprintf(
		"New mail received\nSender: %s <%s>\nSubject: %s\nTime: %s\n\nBody:\n%s",
		emptyAsUnknown(mail.SenderName),
		emptyAsUnknown(mail.SenderEmail),
		emptyAsUnknown(mail.Subject),
		emptyAsUnknown(mail.ReceivedTime),
		emptyAsUnknown(mail.Body),
	)
}

func bodySummary(body string) string {
	body = strings.TrimSpace(strings.ReplaceAll(body, "\r\n", "\n"))
	body = strings.Join(strings.Fields(body), " ")
	const maxRunes = 1200
	runes := []rune(body)
	if len(runes) > maxRunes {
		return string(runes[:maxRunes]) + "..."
	}
	return body
}

func encodeSubject(subject string) string {
	return "=?UTF-8?B?" + base64Encode([]byte(subject)) + "?="
}

func base64Encode(data []byte) string {
	const table = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	var b strings.Builder
	for i := 0; i < len(data); i += 3 {
		var n uint32
		remain := len(data) - i
		n = uint32(data[i]) << 16
		if remain > 1 {
			n |= uint32(data[i+1]) << 8
		}
		if remain > 2 {
			n |= uint32(data[i+2])
		}
		b.WriteByte(table[(n>>18)&63])
		b.WriteByte(table[(n>>12)&63])
		if remain > 1 {
			b.WriteByte(table[(n>>6)&63])
		} else {
			b.WriteByte('=')
		}
		if remain > 2 {
			b.WriteByte(table[n&63])
		} else {
			b.WriteByte('=')
		}
	}
	return b.String()
}

func appendLog(line string) {
	if currentState == nil || currentState.logEdit == 0 {
		return
	}
	existing := text(idLogEdit)
	next := fmt.Sprintf("[%s] %s", time.Now().Format("2006-01-02 15:04:05"), line)
	if existing != "" {
		next = existing + "\r\n" + next
	}
	setText(idLogEdit, next)
}

func emptyAsUnknown(value string) string {
	if value == "" {
		return "未知"
	}
	return value
}

func appDir() string {
	exe, err := os.Executable()
	if err != nil {
		dir, _ := os.Getwd()
		return dir
	}
	return filepath.Dir(exe)
}

func messageBox(text string) {
	user32.NewProc("MessageBoxW").Call(0, uintptr(unsafe.Pointer(utf16Ptr(text))), uintptr(unsafe.Pointer(utf16Ptr(appTitle))), 0)
}

func createStatic(parent hwnd, label string, x, y, w, h int) hwnd {
	return createControl("STATIC", label, wsChild|wsVisible, 0, parent, 0, x, y, w, h)
}

func createGroup(parent hwnd, label string, x, y, w, h int) hwnd {
	return createControl("BUTTON", label, wsChild|wsVisible|bsGroupBox, 0, parent, 0, x, y, w, h)
}

func createButton(parent hwnd, label string, id, x, y, w, h int) hwnd {
	return createControl("BUTTON", label, wsChild|wsVisible|wsTabStop|bsPushButton, 0, parent, id, x, y, w, h)
}

func createRadio(parent hwnd, label string, id, x, y, w, h int) hwnd {
	return createControl("BUTTON", label, wsChild|wsVisible|wsTabStop|bsAutoRadioButton, 0, parent, id, x, y, w, h)
}

func createCheck(parent hwnd, label string, id, x, y, w, h int) hwnd {
	return createControl("BUTTON", label, wsChild|wsVisible|wsTabStop|bsAutoCheckBox, 0, parent, id, x, y, w, h)
}

func createEdit(parent hwnd, id int, value string, x, y, w, h int, password bool) hwnd {
	style := uint32(wsChild | wsVisible | wsTabStop | wsBorder | esAutoHScroll)
	if id == idLogEdit {
		style = wsChild | wsVisible | wsBorder | esMultiline | esAutoVScroll | wsVScroll | esReadOnly
	}
	if password {
		style |= esPassword
	}
	return createControl("EDIT", value, style, 0, parent, id, x, y, w, h)
}

func createControl(class, label string, style uint32, exStyle uint32, parent hwnd, id, x, y, w, h int) hwnd {
	handle, _, _ := procCreateWindowExW.Call(
		uintptr(exStyle),
		uintptr(unsafe.Pointer(utf16Ptr(class))),
		uintptr(unsafe.Pointer(utf16Ptr(label))),
		uintptr(style),
		uintptr(x), uintptr(y), uintptr(w), uintptr(h),
		uintptr(parent), uintptr(id), 0, 0,
	)
	if currentState != nil && currentState.font != 0 {
		procSendMessageW.Call(handle, wmSetFont, currentState.font, 1)
	}
	return hwnd(handle)
}

func createUIFont() uintptr {
	fontName := utf16Ptr("Microsoft YaHei UI")
	font, _, _ := procCreateFontW.Call(
		uintptr(^uint32(15)+1), 0, 0, 0,
		400, 0, 0, 0,
		1, 0, 0, 5, 0,
		uintptr(unsafe.Pointer(fontName)),
	)
	return font
}

func find(id int) hwnd {
	ret, _, _ := user32.NewProc("GetDlgItem").Call(uintptr(currentState.hwnd), uintptr(id))
	return hwnd(ret)
}

func text(id int) string {
	h := find(id)
	buf := make([]uint16, 4096)
	procGetWindowTextW.Call(uintptr(h), uintptr(unsafe.Pointer(&buf[0])), uintptr(len(buf)))
	return syscall.UTF16ToString(buf)
}

func setText(id int, value string) {
	procSetWindowTextW.Call(uintptr(find(id)), uintptr(unsafe.Pointer(utf16Ptr(value))))
}

func check(id int, checked bool) {
	value := uintptr(0)
	if checked {
		value = 1
	}
	procSendMessageW.Call(uintptr(find(id)), bmSetCheck, value, 0)
}

func isChecked(id int) bool {
	ret, _, _ := procSendMessageW.Call(uintptr(find(id)), bmGetCheck, 0, 0)
	return ret == 1
}

func show(h hwnd, visible bool) {
	cmd := uintptr(0)
	if visible {
		cmd = 1
	}
	procShowWindow.Call(uintptr(h), cmd)
}

func utf16Ptr(value string) *uint16 {
	return syscall.StringToUTF16Ptr(value)
}

type wndclassex struct {
	cbSize        uint32
	style         uint32
	lpfnWndProc   uintptr
	cbClsExtra    int32
	cbWndExtra    int32
	hInstance     uintptr
	hIcon         uintptr
	hCursor       uintptr
	hbrBackground uintptr
	lpszMenuName  uintptr
	lpszClassName uintptr
	hIconSm       uintptr
}

type msg struct {
	hwnd    uintptr
	message uint32
	wParam  uintptr
	lParam  uintptr
	time    uint32
	pt      point
}

type point struct {
	x int32
	y int32
}

const (
	wsOverlappedWindow = 0x00CF0000
	wsChild            = 0x40000000
	wsVisible          = 0x10000000
	wsBorder           = 0x00800000
	wsVScroll          = 0x00200000
	wsTabStop          = 0x00010000
	wsExClientEdge     = 0x00000200

	bsPushButton      = 0x00000000
	bsAutoCheckBox    = 0x00000003
	bsGroupBox        = 0x00000007
	bsAutoRadioButton = 0x00000009

	esAutoHScroll = 0x00000080
	esAutoVScroll = 0x00000040
	esMultiline   = 0x00000004
	esPassword    = 0x00000020
	esReadOnly    = 0x00000800

	wmCommand = 0x0111
	wmClose   = 0x0010
	wmDestroy = 0x0002
	wmTimer   = 0x0113
	wmSetFont = 0x0030

	bmGetCheck = 0x00F0
	bmSetCheck = 0x00F1
)
