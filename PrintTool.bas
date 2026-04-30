Attribute VB_Name = "PrintTool"
Option Explicit

' =============================
' Excel VBA 印刷ツール
' =============================

'--- シート名
Private Const SHEET_MAIN As String = "メイン"
Private Const SHEET_LIST As String = "ファイル一覧"
Private Const SHEET_LOG As String = "ログ"

'--- メインシート入力セル
Private Const CELL_TARGET_FOLDER As String = "C4"
Private Const CELL_COPIES As String = "C5"

'--- ファイル一覧列
Private Const COL_PRINT_TARGET As Long = 1 'A
Private Const COL_FILE_TYPE As Long = 2    'B
Private Const COL_FILE_NAME As Long = 3    'C
Private Const COL_FULL_PATH As Long = 4    'D
Private Const COL_RESULT As Long = 5       'E
Private Const COL_MESSAGE As Long = 6      'F

'--- ログ列
Private Const LOG_COL_DATETIME As Long = 1
Private Const LOG_COL_PROCESS As Long = 2
Private Const LOG_COL_FILENAME As Long = 3
Private Const LOG_COL_RESULT As Long = 4
Private Const LOG_COL_MESSAGE As Long = 5
Private Const LOG_COL_PATH As Long = 6

Private Enum TargetJudge
    JudgePrint = 1
    JudgeSkipFalse = 2
    JudgeSkipInvalid = 3
End Enum

Public Sub 一覧抽出ボタン_Click()
    On Error GoTo EH

    Dim targetFolder As String
    targetFolder = SelectTargetFolder()
    If Len(targetFolder) = 0 Then Exit Sub

    If Not FolderExists(targetFolder) Then
        MsgBox "対象フォルダが存在しません。", vbExclamation
        Exit Sub
    End If

    Dim includeSubFolders As VbMsgBoxResult
    includeSubFolders = MsgBox("サブフォルダ内のファイルも対象にしますか？", vbYesNoCancel + vbQuestion)
    If includeSubFolders = vbCancel Then Exit Sub

    If HasExistingListData() Then
        Dim clearAnswer As VbMsgBoxResult
        clearAnswer = MsgBox("既存のファイル一覧をクリアして再抽出します。" & vbCrLf & "よろしいですか？", vbYesNoCancel + vbQuestion)
        If clearAnswer <> vbYes Then Exit Sub
    End If

    WriteLog "一覧抽出", "", "情報", "一覧抽出を開始しました。", ""

    ClearListData

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim files As Collection
    Set files = New Collection

    CollectTargetFiles fso.GetFolder(targetFolder), (includeSubFolders = vbYes), files

    OutputFileList files

    WriteLog "一覧抽出", "", "成功", "一覧抽出を終了しました。抽出件数=" & CStr(files.Count), ""
    MsgBox "一覧抽出が完了しました。抽出件数: " & CStr(files.Count) & "件", vbInformation
    Exit Sub
EH:
    WriteLog "一覧抽出", "", "失敗", "一覧抽出でエラーが発生しました: " & Err.Description, ""
    MsgBox "一覧抽出でエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Function SelectTargetFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "対象フォルダを選択してください"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        SelectTargetFolder = CStr(.SelectedItems(1))
    End With

    GetMainSheet().Range(CELL_TARGET_FOLDER).Value = SelectTargetFolder
End Function

Public Sub 印刷ボタン_Click()
    On Error GoTo FatalEH

    Dim copies As Long
    If Not TryGetCopies(copies) Then Exit Sub

    Dim wsList As Worksheet
    Set wsList = GetListSheet()

    Dim lastRow As Long
    lastRow = GetLastDataRow(wsList)
    If lastRow < 2 Then
        MsgBox "ファイル一覧が存在しません。", vbExclamation
        Exit Sub
    End If

    Dim printTargetCount As Long
    printTargetCount = CountPrintableRows(wsList, lastRow)
    If printTargetCount = 0 Then
        MsgBox "印刷対象ファイルが0件のため処理を中止します。", vbExclamation
        Exit Sub
    End If

    Dim totalJobs As Long
    totalJobs = printTargetCount * copies

    Dim msg As String
    msg = "以下の条件で印刷を開始します。" & vbCrLf & _
          "印刷対象ファイル数：" & CStr(printTargetCount) & "件" & vbCrLf & _
          "全体部数：" & CStr(copies) & "部" & vbCrLf & _
          "合計印刷処理回数：" & CStr(totalJobs) & "回" & vbCrLf & _
          "印刷を開始しますか？"

    If MsgBox(msg, vbYesNoCancel + vbQuestion) <> vbYes Then Exit Sub

    WriteLog "印刷", "", "情報", "印刷処理を開始しました。", ""

    Dim wordApp As Object
    Dim ppApp As Object

    If HasTypeInList(wsList, lastRow, "Word") Then
        Set wordApp = CreateObject("Word.Application")
        If wordApp Is Nothing Then
            WriteLog "印刷", "", "失敗", "Wordアプリケーションを起動できません。", ""
            MsgBox "Wordアプリケーションを起動できません。", vbCritical
            GoTo Cleanup
        End If
        wordApp.Visible = False
    End If

    If HasTypeInList(wsList, lastRow, "PowerPoint") Then
        Set ppApp = CreateObject("PowerPoint.Application")
        If ppApp Is Nothing Then
            MsgBox "PowerPointアプリケーションを起動できません。", vbCritical
            GoTo Cleanup
        End If
        ppApp.Visible = False
    End If

    Dim copyIndex As Long
    Dim r As Long
    For copyIndex = 1 To copies
        WriteLog "印刷", "", "情報", "全体印刷（" & CStr(copyIndex) & "部目）を開始しました。", ""
        For r = 2 To lastRow
            ProcessOneRow wsList, r, wordApp, ppApp
        Next r
    Next copyIndex

    WriteLog "印刷", "", "成功", "印刷処理を終了しました。", ""
    MsgBox "印刷処理が完了しました。", vbInformation

Cleanup:
    On Error Resume Next
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    If Not ppApp Is Nothing Then ppApp.Quit
    On Error GoTo 0
    Exit Sub

FatalEH:
    WriteLog "印刷", "", "失敗", "印刷処理を中止しました: " & Err.Description, ""
    MsgBox "印刷処理で致命的なエラーが発生しました: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub ログクリアボタン_Click()
    Dim ws As Worksheet
    Set ws = GetLogSheet()
    ws.Rows("2:" & ws.Rows.Count).ClearContents
    MsgBox "ログをクリアしました。", vbInformation
End Sub

Private Sub ProcessOneRow(ByVal wsList As Worksheet, ByVal rowIndex As Long, ByVal wordApp As Object, ByVal ppApp As Object)
    On Error GoTo EH

    Dim fileType As String, fileName As String, fullPath As String
    fileType = Trim$(CStr(wsList.Cells(rowIndex, COL_FILE_TYPE).Value))
    fileName = Trim$(CStr(wsList.Cells(rowIndex, COL_FILE_NAME).Value))
    fullPath = Trim$(CStr(wsList.Cells(rowIndex, COL_FULL_PATH).Value))

    Dim judge As TargetJudge
    judge = JudgePrintTarget(wsList.Cells(rowIndex, COL_PRINT_TARGET).Value)

    If judge = JudgeSkipFalse Then
        SetResult wsList, rowIndex, "スキップ", "印刷対象=FALSEのためスキップ"
        WriteLog "印刷", fileName, "スキップ", "印刷対象=FALSEのためスキップ", fullPath
        Exit Sub
    ElseIf judge = JudgeSkipInvalid Then
        SetResult wsList, rowIndex, "スキップ", "印刷対象列の値が不正のためスキップ"
        WriteLog "印刷", fileName, "スキップ", "印刷対象列の値が不正のためスキップ", fullPath
        Exit Sub
    End If

    If Len(fullPath) = 0 Or Not FileExists(fullPath) Then
        SetResult wsList, rowIndex, "失敗", "ファイルが存在しません"
        WriteLog "印刷", fileName, "失敗", "ファイルが存在しません", fullPath
        Exit Sub
    End If

    Select Case fileType
        Case "Excel"
            PrintExcelFile fullPath
        Case "Word"
            PrintWordFile wordApp, fullPath
        Case "PowerPoint"
            PrintPowerPointFile ppApp, fullPath
        Case Else
            Err.Raise vbObjectError + 5000, , "未対応のファイル種別です: " & fileType
    End Select

    SetResult wsList, rowIndex, "成功", "印刷完了"
    WriteLog "印刷", fileName, "成功", "印刷完了", fullPath
    Exit Sub

EH:
    SetResult wsList, rowIndex, "失敗", "印刷処理に失敗しました: " & Err.Description
    WriteLog "印刷", fileName, "失敗", "印刷処理に失敗しました: " & Err.Description, fullPath
End Sub

Private Sub PrintExcelFile(ByVal fullPath As String)
    Dim targetBook As Workbook
    Dim oldAutomationSecurity As MsoAutomationSecurity
    Dim hadError As Boolean
    Dim errMsg As String

    oldAutomationSecurity = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    On Error GoTo EH
    Set targetBook = Application.Workbooks.Open(Filename:=fullPath, UpdateLinks:=False, ReadOnly:=True)
    targetBook.PrintOut Copies:=1
Cleanup:
    On Error Resume Next
    If Not targetBook Is Nothing Then targetBook.Close SaveChanges:=False
    Application.AutomationSecurity = oldAutomationSecurity
    On Error GoTo 0
    If hadError Then
        Err.Raise vbObjectError + 5101, "PrintExcelFile", errMsg
    End If
    Exit Sub
EH:
    hadError = True
    errMsg = Err.Description
    Resume Cleanup
End Sub

Private Sub PrintWordFile(ByVal wordApp As Object, ByVal fullPath As String)
    Dim doc As Object
    Dim hadError As Boolean
    Dim errMsg As String
    On Error GoTo EH
    Set doc = wordApp.Documents.Open(FileName:=fullPath, ReadOnly:=True, AddToRecentFiles:=False)
    doc.PrintOut Background:=False, Copies:=1
Cleanup:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    On Error GoTo 0
    If hadError Then
        Err.Raise vbObjectError + 5102, "PrintWordFile", errMsg
    End If
    Exit Sub
EH:
    hadError = True
    errMsg = Err.Description
    Resume Cleanup
End Sub

Private Sub PrintPowerPointFile(ByVal ppApp As Object, ByVal fullPath As String)
    Dim pres As Object
    Dim hadError As Boolean
    Dim errMsg As String
    On Error GoTo EH
    Set pres = ppApp.Presentations.Open(FileName:=fullPath, ReadOnly:=msoTrue, Untitled:=msoFalse, WithWindow:=msoFalse)
    pres.PrintOut Copies:=1
Cleanup:
    On Error Resume Next
    If Not pres Is Nothing Then pres.Close
    On Error GoTo 0
    If hadError Then
        Err.Raise vbObjectError + 5103, "PrintPowerPointFile", errMsg
    End If
    Exit Sub
EH:
    hadError = True
    errMsg = Err.Description
    Resume Cleanup
End Sub

Private Function TryGetCopies(ByRef copies As Long) As Boolean
    Dim v As Variant
    v = GetMainSheet().Range(CELL_COPIES).Value

    If Len(Trim$(CStr(v))) = 0 Then
        MsgBox "全体部数を入力してください。", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(v) Then
        MsgBox "全体部数は1以上の整数で入力してください。", vbExclamation
        Exit Function
    End If

    Dim dv As Double
    dv = CDbl(v)

    If dv < 1# Or dv > 2147483647# Or dv <> Fix(dv) Then
        MsgBox "全体部数は1以上の整数で入力してください。", vbExclamation
        Exit Function
    End If

    copies = CLng(dv)
    TryGetCopies = True
End Function

Private Function JudgePrintTarget(ByVal v As Variant) As TargetJudge
    Dim s As String
    s = Trim$(CStr(v))

    If Len(s) = 0 Then
        JudgePrintTarget = JudgePrint
        Exit Function
    End If

    Select Case LCase$(s)
        Case "true": JudgePrintTarget = JudgePrint
        Case "false": JudgePrintTarget = JudgeSkipFalse
        Case Else: JudgePrintTarget = JudgeSkipInvalid
    End Select
End Function

Private Sub CollectTargetFiles(ByVal folder As Object, ByVal includeSub As Boolean, ByRef files As Collection)
    On Error GoTo EH

    Dim file As Object
    For Each file In folder.Files
        If IsTargetFile(CStr(file.Name)) Then
            files.Add CStr(file.Path)
        End If
    Next file

    If includeSub Then
        Dim subFolder As Object
        For Each subFolder In folder.SubFolders
            CollectTargetFiles subFolder, True, files
        Next subFolder
    End If
    Exit Sub

EH:
    WriteLog "一覧抽出", "", "警告", "フォルダを読み取れなかったため一部をスキップしました: " & CStr(folder.Path) & " / " & Err.Description, ""
    Err.Clear
End Sub

Private Function IsTargetFile(ByVal fileName As String) As Boolean
    If Left$(fileName, 2) = "~$" Then Exit Function

    Dim ext As String
    ext = LCase$(GetExtension(fileName))

    Select Case ext
        Case "xlsx", "xlsm", "xls", "docx", "docm", "doc", "pptx", "pptm", "ppt"
            IsTargetFile = True
    End Select
End Function

Private Function FileTypeByName(ByVal fileName As String) As String
    Dim ext As String
    ext = LCase$(GetExtension(fileName))
    Select Case ext
        Case "xlsx", "xlsm", "xls": FileTypeByName = "Excel"
        Case "docx", "docm", "doc": FileTypeByName = "Word"
        Case "pptx", "pptm", "ppt": FileTypeByName = "PowerPoint"
    End Select
End Function

Private Function GetExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then GetExtension = Mid$(fileName, p + 1)
End Function

Private Sub OutputFileList(ByVal files As Collection)
    Dim ws As Worksheet
    Set ws = GetListSheet()

    WriteListHeader ws

    If files.Count = 0 Then Exit Sub

    Dim arr() As String
    Dim i As Long
    ReDim arr(1 To files.Count)
    For i = 1 To files.Count
        arr(i) = CStr(files(i))
    Next i

    If files.Count > 1 Then
        QuickSortStrings arr, 1, UBound(arr)
    End If

    Dim r As Long
    r = 2
    For i = LBound(arr) To UBound(arr)
        ws.Cells(r, COL_PRINT_TARGET).Value = ""
        ws.Cells(r, COL_FILE_NAME).Value = Mid$(arr(i), InStrRev(arr(i), "\") + 1)
        ws.Cells(r, COL_FULL_PATH).Value = arr(i)
        ws.Cells(r, COL_FILE_TYPE).Value = FileTypeByName(CStr(ws.Cells(r, COL_FILE_NAME).Value))
        r = r + 1
    Next i
End Sub

Private Sub WriteListHeader(ByVal ws As Worksheet)
    ws.Cells(1, COL_PRINT_TARGET).Value = "印刷対象"
    ws.Cells(1, COL_FILE_TYPE).Value = "ファイル種別"
    ws.Cells(1, COL_FILE_NAME).Value = "ファイル名"
    ws.Cells(1, COL_FULL_PATH).Value = "フルパス"
    ws.Cells(1, COL_RESULT).Value = "印刷結果"
    ws.Cells(1, COL_MESSAGE).Value = "メッセージ"
End Sub

Private Function CountPrintableRows(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
    Dim r As Long
    For r = 2 To lastRow
        If JudgePrintTarget(ws.Cells(r, COL_PRINT_TARGET).Value) = JudgePrint Then
            CountPrintableRows = CountPrintableRows + 1
        End If
    Next r
End Function

Private Function HasTypeInList(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal targetType As String) As Boolean
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, COL_FILE_TYPE).Value) = targetType Then
            HasTypeInList = True
            Exit Function
        End If
    Next r
End Function

Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    GetLastDataRow = ws.Cells(ws.Rows.Count, COL_FULL_PATH).End(xlUp).Row
End Function

Private Function HasExistingListData() As Boolean
    HasExistingListData = (GetLastDataRow(GetListSheet()) >= 2)
End Function

Private Sub ClearListData()
    Dim ws As Worksheet
    Set ws = GetListSheet()
    ws.Cells.ClearContents
    WriteListHeader ws
End Sub

Private Sub SetResult(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal resultText As String, ByVal messageText As String)
    ws.Cells(rowIndex, COL_RESULT).Value = resultText
    ws.Cells(rowIndex, COL_MESSAGE).Value = messageText
End Sub

Private Sub WriteLog(ByVal processType As String, ByVal fileName As String, ByVal resultText As String, ByVal messageText As String, ByVal fullPath As String)
    Dim ws As Worksheet
    Set ws = GetLogSheet()

    If ws.Cells(1, LOG_COL_DATETIME).Value = "" Then
        ws.Cells(1, LOG_COL_DATETIME).Value = "実行日時"
        ws.Cells(1, LOG_COL_PROCESS).Value = "処理種別"
        ws.Cells(1, LOG_COL_FILENAME).Value = "ファイル名"
        ws.Cells(1, LOG_COL_RESULT).Value = "結果"
        ws.Cells(1, LOG_COL_MESSAGE).Value = "メッセージ"
        ws.Cells(1, LOG_COL_PATH).Value = "フルパス"
    End If

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, LOG_COL_DATETIME).End(xlUp).Row + 1

    ws.Cells(nextRow, LOG_COL_DATETIME).Value = Now
    ws.Cells(nextRow, LOG_COL_PROCESS).Value = processType
    ws.Cells(nextRow, LOG_COL_FILENAME).Value = fileName
    ws.Cells(nextRow, LOG_COL_RESULT).Value = resultText
    ws.Cells(nextRow, LOG_COL_MESSAGE).Value = messageText
    ws.Cells(nextRow, LOG_COL_PATH).Value = fullPath
End Sub

Private Function GetMainSheet() As Worksheet
    Set GetMainSheet = ThisWorkbook.Worksheets(SHEET_MAIN)
End Function

Private Function GetListSheet() As Worksheet
    Set GetListSheet = ThisWorkbook.Worksheets(SHEET_LIST)
End Function

Private Function GetLogSheet() As Worksheet
    Set GetLogSheet = ThisWorkbook.Worksheets(SHEET_LOG)
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long
    Dim pivot As String, temp As String

    low = first
    high = last
    pivot = arr((first + last) \ 2)

    Do While low <= high
        Do While arr(low) < pivot
            low = low + 1
        Loop
        Do While arr(high) > pivot
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortStrings arr, first, high
    If low < last Then QuickSortStrings arr, low, last
End Sub
