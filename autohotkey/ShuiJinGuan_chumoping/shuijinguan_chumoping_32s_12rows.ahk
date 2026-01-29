#Requires AutoHotkey v2.0
; --- 全局配置 ---
CoordMode("Mouse", "Screen")
SetDefaultMouseSpeed(2)

; --- 参数定义 ---
PhaseRows := 12              ; 每列 12 行

; 1. 右侧 Excel 坐标参数
Excel_BlankX := 1356         ; 激活 Excel 的空白位置
Excel_BlankY := 215          
ExcelX := 1347               ; Excel 第一行单元格 X
ExcelY_Start := 297          ; Excel 第一行单元格 Y

; 2. 左侧 MCGS 表格行坐标 (12行新坐标)
TargetY_Start := 280         
TargetY_Step := 46.18        ; 步进值: (788-280)/11
TargetX_Phase1 := 92         
TargetX_Phase2 := 594        

; 3. 固定属性框坐标 (粘贴位)
PasteX := 538
PasteY := 1201

TargetWin := "ahk_exe McgsSetPro.exe"

F1:: {
    Loop PhaseRows * 2 {
        ; --- 计算当前所属阶段和对应的 MCGS 坐标 ---
        if (A_Index <= PhaseRows) {
            ; 第一阶段：1-12行
            CurrentTargetX := TargetX_Phase1
            CurrentTargetY := TargetY_Start + (A_Index - 1) * TargetY_Step
        } else {
            ; 第二阶段：13-24行 (Y轴重置)
            CurrentTargetX := TargetX_Phase2
            CurrentTargetY := TargetY_Start + (A_Index - PhaseRows - 1) * TargetY_Step
        }

        ; --- 步骤 1: Excel 激活与复制 (完全方向键逻辑) ---
        ; 点击空白处激活窗口，不破坏选中状态
        MouseMove(Excel_BlankX, Excel_BlankY)
        Click()
        Sleep(300) 
        
        if (A_Index == 1) {
            ; 只有第 1 次运行点击起始坐标
            MouseMove(ExcelX, ExcelY_Start)
            Click()
            Sleep(200)
        } else {
            ; 之后所有行全部通过向下方向键移动
            Send("{Down}")
            Sleep(200)
        }
        
        ; 强制退出可能意外进入的编辑模式并复制
        Send("{Esc}") 
        Sleep(100)
        A_Clipboard := "" 
        Send("^c")
        
        if !ClipWait(1.5) {
            MsgBox("第 " A_Index " 次循环 Excel 复制失败，请检查窗口状态。")
            return
        }

        ; --- 步骤 2: 激活 MCGS 并执行单次点击 ---
        if WinExist(TargetWin) {
            WinActivate(TargetWin)
            WinWaitActive(TargetWin, , 1)
        }

        MouseMove(CurrentTargetX, CurrentTargetY)
        Sleep(200)
        Click()            ; 选中 MCGS 表格行
        Sleep(600)         ; 等待底部属性框加载内容

        ; --- 步骤 3: 属性框粘贴 ---
        MouseMove(PasteX, PasteY)
        Sleep(150)
        Click() 
        Sleep(200)
        Send("^v")

        ; 行间缓冲
        Sleep(600)
    }
    MsgBox("24次任务 (12x2) 已全部完成！")
}

; --- 功能快捷键 ---
F2::Pause()      ; 暂停/恢复
F3::Reload()     ; 重启脚本
F4::ExitApp()    ; 退出脚本