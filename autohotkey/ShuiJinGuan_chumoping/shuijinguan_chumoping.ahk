; 设置
RowCount := 15
StartX := 1770
StartY := 293
PasteX := 159
PasteY := 269
YOffset := 39

Loop RowCount {
    ; --- 复制阶段 ---
    A_Clipboard := "" ; 清空剪贴板，方便后续检测是否复制成功
    MouseMove(StartX, StartY + (YOffset * (A_Index - 1)))
    Click()
    Sleep(100)
    Send("^c")
    
    ; 等待剪贴板包含数据，最多等1秒
    if !ClipWait(1) {
        MsgBox "复制失败，行数: " . A_Index
        return
    }

    ; --- 粘贴阶段 ---
    MouseMove(PasteX, PasteY + (YOffset * (A_Index - 1)))
    Click()
    Sleep(200) ; 增加一点延迟，确保目标框已激活
    Send("^v")
    
    Sleep(500) ; 等待处理
}

MsgBox "任务完成！"