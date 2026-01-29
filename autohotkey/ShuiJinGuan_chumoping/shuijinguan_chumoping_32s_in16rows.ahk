#Requires AutoHotkey v2.0
; --- 全局配置 ---
CoordMode("Mouse", "Screen")
SetDefaultMouseSpeed(2)

; --- 参数定义 ---
PhaseRows := 16              

; 1. 右侧 Excel 坐标
Excel_BlankX := 1356         
Excel_BlankY := 215          
ExcelX := 1347               
ExcelY_Start := 297          

; 2. 左侧 MCGS 坐标 (如果这边间距没问题，暂不修改)
TargetY_Start := 271         
TargetY_Step := 35.0         
TargetX_Phase1 := 93         
TargetX_Phase2 := 598        

; 3. 固定属性框坐标
PasteX := 538
PasteY := 1201

TargetWin := "ahk_exe McgsSetPro.exe"

F1:: {
    Loop PhaseRows * 2 {
        ; --- 计算 MCGS 阶段坐标 ---
        if (A_Index <= PhaseRows) {
            CurrentTargetX := TargetX_Phase1
            CurrentTargetY := TargetY_Start + (A_Index - 1) * TargetY_Step
        } else {
            CurrentTargetX := TargetX_Phase2
            CurrentTargetY := TargetY_Start + (A_Index - PhaseRows - 1) * TargetY_Step
        }

        ; --- 步骤 1: Excel 复制 (方向键精准切换版) ---
        ; 激活 Excel 窗口
        MouseMove(Excel_BlankX, Excel_BlankY)
        Click()
        Sleep(300) 
        
        ; 仅在第一行时点击具体坐标，后续行使用方向键
        if (A_Index == 1) {
            MouseMove(ExcelX, ExcelY_Start)
            Click()
            Sleep(200)
        } else {
            ; 模拟按下向下键，Excel 会自动选中下一行，永不偏离
            Send("{Down}")
            Sleep(200)
        }
        
        Send("{Esc}")      ; 确保退出可能的编辑模式
        Sleep(100)
        
        A_Clipboard := ""  ; 清空剪贴板
        Send("^c")
        
        if !ClipWait(1.5) {
            MsgBox("第 " A_Index " 行 Excel 复制失败。`n提示：请确保 Excel 窗口处于选中状态。")
            return
        }

        ; --- 步骤 2: 激活 MCGS 并执行单次点击 ---
        if WinExist(TargetWin) {
            WinActivate(TargetWin)
            WinWaitActive(TargetWin, , 1)
        }

        MouseMove(CurrentTargetX, CurrentTargetY)
        Sleep(200)
        Click() 
        Sleep(600) 

        ; --- 步骤 3: 属性框粘贴 ---
        MouseMove(PasteX, PasteY)
        Sleep(150)
        Click() 
        Sleep(200)
        Send("^v")

        Sleep(600)
    }
    MsgBox("32次任务流程执行完毕！")
}

; --- 快捷键 ---
F2::Pause()
F3::Reload()
F4::ExitApp()