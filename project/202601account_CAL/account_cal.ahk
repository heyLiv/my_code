#Requires AutoHotkey v2.0
#SingleInstance Force

; ==============================================================================
; 【1. 账户配置】随时修改，程序自动适配
; ==============================================================================
Global Accounts := ["微信钱包", "支付宝", "中国银行", "建设银行2240", "招商银行", "建行社保5126", "花呗欠款"]
Global DataDir := A_MyDocuments "\MyAssetsManager"
Global DataFile := DataDir "\Assets_History_v4.csv"
Global BackupDir := DataDir "\Backups"

; 初始化文件夹和备份
if !DirExist(BackupDir)
    DirCreate(BackupDir)
if FileExist(DataFile) ; 每次启动自动创建一个今日备份
    FileCopy(DataFile, BackupDir "\AutoBackup_" FormatTime(,"yyyyMMdd") ".csv", 1)

; ------------------------------------------------------------------------------
; 【2. 界面构建】
; ------------------------------------------------------------------------------
MainGui := Gui("+Resize", "个人资产管家 v4.1 - 长期稳健版")
MainGui.SetFont("s10", "Microsoft YaHei")

; --- 左侧：录入面板 ---
MainGui.Add("GroupBox", "x15 y15 w250 h" (190 + Accounts.Length * 38), "📝 资产录入/编辑")
Edits := Map()
curY := 50
for index, name in Accounts {
    MainGui.Add("Text", "x30 y" curY " w85", name ":")
    defaultVal := (name = "花呗欠款") ? "-0.0" : "0.0"
    Edits[name] := MainGui.Add("Edit", "x120 yp-3 w110 h26 Center", defaultVal)
    curY += 38
}
MainGui.Add("Text", "x30 y" curY, "备注 (双击右侧行可修改):")
NoteEdit := MainGui.Add("Edit", "x30 y+5 w220 r2")
SaveBtn := MainGui.Add("Button", "x30 y+20 w220 h45 Default", "确认保存当前快照")
SaveBtn.OnEvent("Click", ProcessSave)

; --- 右侧：表格面板 ---
HeaderArray := ["年份/月份", "日期", "总余额"]
for aName in Accounts
    HeaderArray.Push(aName)
HeaderArray.Push("备注记录")

LV := MainGui.Add("ListView", "x285 y20 w950 h660 Grid -Multi", HeaderArray)

; 绑定交互事件
LV.OnEvent("DoubleClick", HandleDoubleClick) ; 双击回填
LV.OnEvent("ContextMenu", HandleRightClick)  ; 右键菜单（删除）

; 检查结构并加载
CheckFileStructure()
LoadData()
MainGui.Show()

; ------------------------------------------------------------------------------
; 【3. 核心功能函数】
; ------------------------------------------------------------------------------

; 检查账户数量是否变动，若变动则自动备份
CheckFileStructure() {
    if !FileExist(DataFile) {
        WriteNewHeader()
        return
    }
    fileContent := FileRead(DataFile, "UTF-8")
    firstLine := StrSplit(fileContent, "`n")[1]
    expectedCount := Accounts.Length + 4
    if (StrSplit(firstLine, ",").Length != expectedCount) {
        MsgBox("检测到账户配置已更改。系统已自动将旧账本备份至 Backups 文件夹，并为您创建新表头。", "配置更新", "Iconi")
        FileMove(DataFile, BackupDir "\OldStructure_Backup_" FormatTime(,"yyyyMMddHHmmss") ".csv")
        WriteNewHeader()
    }
}

WriteNewHeader() {
    hdr := "年份/月份,日期,总余额"
    for n in Accounts
        hdr .= "," n
    hdr .= ",备注`n"
    FileAppend(hdr, DataFile, "UTF-8")
}

; 加载数据并强制保留1位小数
LoadData(*) {
    LV.Delete()
    if !FileExist(DataFile)
        return
    
    content := FileRead(DataFile, "UTF-8")
    lastMonth := ""
    Loop parse, content, "`n", "`r" {
        if (A_Index = 1 || Trim(A_LoopField) = "") continue
        row := StrSplit(A_LoopField, ",")
        displayRow := []
        for i, val in row {
            if (i >= 3 && i < row.Length && IsNumber(val))
                displayRow.Push(Format("{:.1f}", Float(val)))
            else
                displayRow.Push(val)
        }
        if (displayRow.Length > 0) {
            if (displayRow[1] == lastMonth) 
                displayRow[1] := "" 
            else 
                lastMonth := row[1]
        }
        LV.Add(, displayRow*)
    }
    ; 设置列宽
    LV.ModifyCol(1, 100), LV.ModifyCol(2, 110), LV.ModifyCol(3, 110)
    Loop Accounts.Length
        LV.ModifyCol(A_Index + 3, 100)
    LV.ModifyCol(HeaderArray.Length, 300)
}

; 功能：双击某行，将数据填回输入框
HandleDoubleClick(GuiCtrl, RowNumber) {
    if (RowNumber = 0) return
    ; 提示用户
    SoundBeep()
    ; 重新读取该行原始数据（跳过年份/月份显示优化带来的空白）
    content := FileRead(DataFile, "UTF-8")
    lines := StrSplit(content, "`n", "`r")
    targetLine := lines[RowNumber + 1] ; +1 是因为跳过表头
    data := StrSplit(targetLine, ",")
    
    ; 填入各个账户输入框 (CSV中账户从第4列开始)
    for index, name in Accounts {
        if (data.Length >= index + 3)
            Edits[name].Value := Format("{:.1f}", Float(data[index+3]))
    }
    ; 填入备注
    NoteEdit.Value := data[data.Length]
    MsgBox("数据已回填，您可以修改后重新保存。`n(注：原记录不会自动消失，如需替换请右键删除旧记录)", "编辑模式", "T2")
}

; 功能：右键删除
HandleRightClick(GuiCtrl, ItemRow, *) {
    if (ItemRow = 0) return
    
    result := MsgBox("确定要删除这一行记录吗？`n删除后不可恢复（除非从备份文件夹找回）。", "警告", "YesNo Icon!")
    if (result = "Yes") {
        content := FileRead(DataFile, "UTF-8")
        lines := StrSplit(content, "`n", "`r")
        newContent := ""
        for index, line in lines {
            if (index == ItemRow + 1 || Trim(line) == "") continue
            newContent .= line "`n"
        }
        FileOpen(DataFile, "w", "UTF-8").Write(newContent)
        LoadData()
    }
}

ProcessSave(*) {
    total := 0.0
    detailStr := ""
    for name in Accounts {
        val := IsNumber(Edits[name].Value) ? Float(Edits[name].Value) : 0.0
        total += val
        detailStr .= "," Format("{:.1f}", Round(val, 1))
    }
    totalStr := Format("{:.1f}", Round(total, 1))
    newRow := FormatTime(, "yyyy年M月") "," FormatTime(, "yyyy-MM-dd") "," totalStr detailStr "," StrReplace(NoteEdit.Value, ",", " ") "`n"
    
    FileAppend(newRow, DataFile, "UTF-8")
    MsgBox("记录已保存！", , "T1")
    NoteEdit.Value := ""
    LoadData()
}