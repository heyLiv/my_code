#Requires AutoHotkey v2.0
#SingleInstance Force

; 1. 自动请求管理员权限
if !A_IsAdmin
    Run '*RunAs "' A_ScriptFullPath '"'

; 2. 自动设置开机自启
ShortcutPath := A_Startup "\MyAutoHotkeyScript.lnk"
if !FileExist(ShortcutPath) {
    FileCreateShortcut A_ScriptFullPath, ShortcutPath
}

; ==============================================================================
; 通用窗口切换函数
; ==============================================================================
SmartActivate(WinTitle, TargetPath := "") {
    if WinExist(WinTitle) {
        if WinActive(WinTitle) {
            WinMinimize 
        } else {
            WinActivate 
        }
    } else if (TargetPath != "") {
        try {
            Run TargetPath
        } catch {
            MsgBox "无法启动程序，请检查路径：`n" TargetPath
        }
    }
}

; ==============================================================================
; 按键重映射
; ==============================================================================
Home::F11
End::^r

; ==============================================================================
; 软件快捷键设置 (Alt + Key)
; ==============================================================================

; Alt + A → Edge 浏览器
!a::SmartActivate("ahk_exe msedge.exe", "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")

; Alt + W → 微信
!w::SmartActivate("ahk_exe Weixin.exe", "C:\Program Files\Tencent\Weixin\Weixin.exe")

; Alt + S → VS Code
!s::SmartActivate("ahk_exe Code.exe", "C:\Users\liv\AppData\Local\Programs\Microsoft VS Code\Code.exe")

; Alt + Q → 资源管理器
!q::SmartActivate("ahk_class CabinetWClass", "explorer.exe")

; Alt + Z → WPS Office
!z::SmartActivate("ahk_exe wps.exe", "C:\Users\liv\AppData\Local\Kingsoft\WPS Office\ksolaunch.exe")

; Alt + X → SumatraPDF
!x::SmartActivate("ahk_exe SumatraPDF.exe", "C:\Users\liv\AppData\Local\SumatraPDF\SumatraPDF.exe")

; Alt + , (逗号) → Bilibili
!,::SmartActivate("ahk_exe 哔哩哔哩.exe", "C:\Program Files\bilibili\哔哩哔哩.exe")

; Alt + . (句号) → QQ (已改为 QQNT 路径)
!.::SmartActivate("ahk_exe QQ.exe", "C:\Program Files\Tencent\QQNT\QQ.exe")

; Alt + \ (反斜杠) → Clash Verge
!\::SmartActivate("ahk_exe clash-verge.exe", "C:\Program Files\Clash Verge\clash-verge.exe")

#q::return