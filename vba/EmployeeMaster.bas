Attribute VB_Name = "EmployeeMaster"
Option Explicit
Public Busy As Boolean
Private Const PWD As String = "master"

Private Function Ws(ByVal name As String) As Object
    Set Ws = ThisWorkbook.Worksheets(name)
End Function
Private Function Tbl(ByVal name As String) As Object
    Dim s As Object, t As Object
    For Each s In ThisWorkbook.Worksheets
        For Each t In s.ListObjects
            If t.name = name Then Set Tbl = t: Exit Function
        Next t
    Next s
    Fail "必要なテーブルが見つかりません: " & name
End Function
Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 513, "従業員マスター", message
End Sub
Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Fail "セルに数式エラーがあります。"
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    Txt = Trim(CStr(v))
End Function
Private Function DayValue(ByVal v As Variant, ByVal label As String) As Date
    If Txt(v) = "" Then Fail label & "を入力してください。"
    If IsNumeric(v) Then
        If CDbl(v) < 1 Or CDbl(v) > 2958465 Then Fail label & "の日付が不正です。"
        DayValue = DateValue(CDate(CDbl(v)))
    ElseIf IsDate(v) Then
        DayValue = DateValue(CDate(v))
    Else
        Fail label & "は 2026/4/1 のように入力してください。"
    End If
End Function
Private Function MonthValue(ByVal v As Variant, ByVal label As String) As Date
    Dim d As Date: d = DayValue(v, label)
    If Day(d) <> 1 Then Fail label & "は月の1日を入力してください。例: 2026/4/1"
    MonthValue = d
End Function
Private Function DisplayPerson(ByVal r As Object) As String
    DisplayPerson = Txt(r.Cells(1, 2).Value2)
    If Txt(r.Cells(1, 4).Value2) <> "" Then DisplayPerson = DisplayPerson & " (" & Txt(r.Cells(1, 4).Value2) & ")"
End Function
Private Function FindEmployee(ByVal id As String) As Object
    Dim row As Object
    For Each row In Tbl("tEmployees").ListRows
        If Txt(row.Range.Cells(1, 1).Value2) = id Then Set FindEmployee = row.Range: Exit Function
    Next row
    Fail "従業員IDが見つかりません: " & id
End Function
Private Function Resolve(ByVal name As String, ByVal tableName As String) As String
    Dim row As Object, found As Long, label As String
    If name = "" Then Fail "氏名・店舗・役職の選択欄を確認してください。"
    For Each row In Tbl(tableName).ListRows
        If Txt(row.Range.Cells(1, 1).Value2) <> "" Then
            If tableName = "tEmployees" Then label = DisplayPerson(row.Range) Else label = Txt(row.Range.Cells(1, 2).Value2)
            If StrComp(label, name, vbTextCompare) = 0 Then found = found + 1: Resolve = Txt(row.Range.Cells(1, 1).Value2)
        End If
    Next row
    If found = 0 Then Fail "候補にない名称です: " & name & "。ホームの候補更新後に選び直してください。"
    If found > 1 Then Fail "同じ名称が複数あります: " & name & "。保守担当に識別補助の設定を依頼してください。"
End Function
Private Function HasID(ByVal id As String) As Boolean
    Dim row As Object
    For Each row In Tbl("tEmployees").ListRows
        If StrComp(Txt(row.Range.Cells(1, 1).Value2), id, vbTextCompare) = 0 Then HasID = True: Exit Function
    Next row
End Function
Private Function Dimension(ByVal kind As String) As String
    Select Case kind
        Case "異動": Dimension = "店舗"
        Case "役職変更": Dimension = "役職"
        Case "退職", "休職", "出向", "復職", "出向から復帰": Dimension = "状態"
        Case Else: Fail "変更内容を候補から選んでください。"
    End Select
End Function
Private Function NewRecord(ByVal id As Variant, ByVal name As Variant, ByVal kana As Variant, ByVal extra As Variant, ByVal joined As Variant, ByVal effective As Variant, ByVal store As Variant, ByVal role As Variant, ByVal state As Variant, ByVal pending As Object) As Variant
    Dim sid As String, rid As String, d As Date, m As Date, key As String, label As String, row As Object
    key = Txt(id): If key = "" Then Fail "新規登録の従業員IDを入力してください。"
    If HasID(key) Or pending.Exists("ID:" & LCase(key)) Then Fail "従業員IDが重複しています: " & key
    If Txt(name) = "" Then Fail "氏名を入力してください。"
    label = Txt(name): If Txt(extra) <> "" Then label = label & " (" & Txt(extra) & ")"
    For Each row In Tbl("tEmployees").ListRows
        If StrComp(DisplayPerson(row.Range), label, vbTextCompare) = 0 Then Fail "同じ氏名があります。識別補助を入力してください: " & label
    Next row
    If pending.Exists("NAME:" & LCase(label)) Then Fail "一括登録内に同じ氏名があります。識別補助を入力してください。"
    d = DayValue(joined, "入社日"): m = MonthValue(effective, "反映開始月")
    If m < DateSerial(Year(d), Month(d), 1) Then Fail "反映開始月が入社月より前です。"
    sid = Resolve(Txt(store), "tStores"): rid = Resolve(Txt(role), "tRoles")
    If Txt(state) = "" Then state = "在籍"
    If Txt(state) <> "在籍" And Txt(state) <> "休職" And Txt(state) <> "出向" Then Fail "初期状態は在籍・休職・出向から選んでください。"
    pending.Add "ID:" & LCase(key), True: pending.Add "NAME:" & LCase(label), True
    NewRecord = Array(key, Txt(name), Txt(kana), Txt(extra), d, m, sid, rid, Txt(state))
End Function
Private Function ChangeRecord(ByVal person As Variant, ByVal kindValue As Variant, ByVal occurred As Variant, ByVal store As Variant, ByVal role As Variant, ByVal effective As Variant, ByVal note As Variant, ByVal pending As Object, Optional ByVal stagedPeople As Variant) As Variant
    Dim id As String, kind As String, dimn As String, d As Date, m As Date, sid As String, rid As String
    Dim key As String, row As Object, e As Object, hdate As Date
    Dim staged As Boolean, employee As Variant, joined As Date, baseline As Date
    If Not IsMissing(stagedPeople) Then
        If stagedPeople.Exists(LCase(Txt(person))) Then
            employee = stagedPeople(LCase(Txt(person)))
            id = employee(0): joined = employee(4): baseline = employee(5): staged = True
        End If
    End If
    If Not staged Then
        id = Resolve(Txt(person), "tEmployees"): Set e = FindEmployee(id)
        joined = DayValue(e.Cells(1, 5).Value2, "入社日")
        baseline = MonthValue(e.Cells(1, 6).Value2, "従業員の反映開始月")
    End If
    kind = Txt(kindValue): dimn = Dimension(kind)
    d = DayValue(occurred, "発生日")
    If d < joined Then Fail "発生日が入社日より前です。"
    Select Case kind
        Case "異動"
            If Day(d) <> 1 Then Fail "異動日は月の1日を指定してください。"
            m = d: sid = Resolve(Txt(store), "tStores")
        Case "役職変更"
            m = MonthValue(effective, "反映開始月"): rid = Resolve(Txt(role), "tRoles")
            If m < DateSerial(Year(d), Month(d), 1) Then Fail "反映開始月が発生月より前です。"
        Case Else
            m = DateSerial(Year(d), Month(d) + 1, 1)
    End Select
    If m < baseline Then Fail "変更の反映月が従業員の反映開始月より前です。"
    key = id & "|" & CStr(CLng(m)) & "|" & dimn
    If pending.Exists(key) Then Fail "同じ人・反映月・変更区分が重複しています。"
    For Each row In Tbl("tHistory").ListRows
        If Txt(row.Range.Cells(1, 2).Value2) = id Then
            If MonthValue(row.Range.Cells(1, 5).Value2, "履歴の反映月") = m And Dimension(Txt(row.Range.Cells(1, 3).Value2)) = dimn Then Fail "同じ反映月の変更が登録済みです。訂正は保守担当へ連絡してください。"
            hdate = DayValue(row.Range.Cells(1, 4).Value2, "履歴の発生日")
            If Txt(row.Range.Cells(1, 3).Value2) = "退職" And d >= hdate Then Fail "退職日以降の変更は登録できません。保守担当へ連絡してください。"
            If kind = "退職" And hdate > d Then Fail "退職日より後の変更が登録済みです。保守担当へ連絡してください。"
        End If
    Next row
    pending.Add key, True
    ChangeRecord = Array("", id, kind, d, m, sid, rid, Txt(note), Now)
End Function
Private Sub AppendRecords(ByVal tableName As String, ByVal records As Collection)
    Dim t As Object, row As Object, rec As Variant, j As Long, oldRows As Long, nextID As Long, n As Long, message As String
    If tableName = "tHistory" Then ValidateBatchHistory records
    Set t = Tbl(tableName): oldRows = t.ListRows.Count
    If tableName = "tHistory" Then
        For Each row In t.ListRows
            If Left(Txt(row.Range.Cells(1, 1).Value2), 1) = "H" Then n = Val(Mid(Txt(row.Range.Cells(1, 1).Value2), 2)): If n > nextID Then nextID = n
        Next row
    End If
    On Error GoTo Rollback
    t.Parent.Unprotect PWD
    For Each rec In records
        If tableName = "tHistory" Then nextID = nextID + 1: rec(0) = "H" & Format(nextID, "000000000")
        Set row = t.ListRows.Add
        For j = LBound(rec) To UBound(rec)
            ' ID・名称・備考を文字列として保持し、先頭ゼロの消失や数式化を防ぎます。
            If VarType(rec(j)) = vbString Then row.Range.Cells(1, j + 1).NumberFormat = "@"
            row.Range.Cells(1, j + 1).Value = rec(j)
        Next j
    Next rec
    t.Parent.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    MarkDirty
    Exit Sub
Rollback:
    message = Err.Description
    If TrimAddedRows(t, oldRows) Then
        Fail "登録を取り消しました。" & message
    Else
        Fail "登録中にエラーが発生し、完全に元へ戻せませんでした。再登録せず、保守担当へ連絡してください。" & message
    End If
End Sub
Private Sub ValidateBatchHistory(ByVal records As Collection)
    Dim i As Long, j As Long, a As Variant, b As Variant
    For i = 1 To records.Count
        a = records(i)
        If a(2) = "退職" Then
            For j = 1 To records.Count
                If i <> j Then
                    b = records(j)
                    If a(1) = b(1) And b(3) >= a(3) Then Fail "一括登録に退職日以降の変更が含まれています。登録内容を確認してください。"
                End If
            Next j
        End If
    Next i
End Sub
Public Sub SaveEmployee()
    Dim c As New Collection, d As Object, s As Object, rec As Variant
    Dim committed As Boolean, message As String
    On Error GoTo ErrorHandler
    Set s = Ws("新規登録"): Set d = CreateObject("Scripting.Dictionary")
    rec = NewRecord(s.Range("D6").Value2, s.Range("D8").Value2, s.Range("D10").Value2, s.Range("D12").Value2, s.Range("D14").Value2, s.Range("D16").Value2, s.Range("D18").Value2, s.Range("D20").Value2, s.Range("D22").Value2, d)
    c.Add rec: Busy = True: AppendRecords "tEmployees", c
    committed = True
    ClearNewFields: RefreshChoices: Busy = False
    MsgBox "登録しました。ホームの「月別表を更新」で集計へ反映してください。", vbInformation
    Exit Sub
ErrorHandler:
    message = Err.Description: Busy = False
    If committed Then message = "台帳への登録は完了しています。重複登録しないでください。" & vbCrLf & message
    MsgBox message, vbExclamation, "登録内容を確認してください"
End Sub
Public Sub SaveChange()
    Dim c As New Collection, d As Object, s As Object, rec As Variant
    Dim committed As Boolean, message As String
    On Error GoTo ErrorHandler
    Set s = Ws("変更登録"): Set d = CreateObject("Scripting.Dictionary")
    rec = ChangeRecord(s.Range("D6").Value2, s.Range("D8").Value2, s.Range("D10").Value2, s.Range("D12").Value2, s.Range("D14").Value2, s.Range("D16").Value2, s.Range("D20").Value2, d)
    c.Add rec: Busy = True: AppendRecords "tHistory", c
    committed = True
    s.Range("D6,D10,D12,D14,D16,D20").ClearContents: Busy = False
    MsgBox "変更を登録しました。月別表を更新してください。", vbInformation
    Exit Sub
ErrorHandler:
    message = Err.Description: Busy = False
    If committed Then message = "台帳への登録は完了しています。重複登録しないでください。" & vbCrLf & message
    MsgBox message, vbExclamation, "登録内容を確認してください"
End Sub
Public Sub ClearNewFields()
    Ws("新規登録").Range("D6,D8,D10,D12,D14,D16,D18,D20").ClearContents
    Ws("新規登録").Range("D22").Value = "在籍"
End Sub
Public Sub ClearChangeFields()
    Ws("変更登録").Range("D6,D10,D12,D14,D16,D20").ClearContents
End Sub
Public Sub SaveBulkChanges()
    SaveBulkUnified
End Sub
Public Sub SaveBulkNew()
    SaveBulkUnified
End Sub
Private Function HasBulkInput(ByVal r As Object) As Boolean
    HasBulkInput = (Application.CountA(r.Resize(1, 11)) > 0)
End Function
Private Sub RequireEmpty(ByVal r As Object, ByVal columns As Variant)
    Dim col As Variant
    For Each col In columns
        If Txt(r.Cells(1, CLng(col)).Value2) <> "" Then Fail "この区分では「" & Tbl("tBulkUnified").HeaderRowRange.Cells(1, CLng(col)).Value2 & "」を空欄にしてください。"
    Next col
End Sub
Private Function ReadBulkNew(ByVal r As Object, ByVal pending As Object) As Variant
    RequireEmpty r, Array(11)
    ReadBulkNew = NewRecord(r.Cells(1, 7).Value2, r.Cells(1, 2).Value2, r.Cells(1, 9).Value2, r.Cells(1, 8).Value2, r.Cells(1, 3).Value2, r.Cells(1, 6).Value2, r.Cells(1, 4).Value2, r.Cells(1, 5).Value2, r.Cells(1, 10).Value2, pending)
End Function
Private Function ReadBulkChange(ByVal r As Object, ByVal pending As Object, ByVal stagedPeople As Object) As Variant
    Dim kind As String: kind = Txt(r.Cells(1, 1).Value2)
    RequireEmpty r, Array(7, 8, 9, 10)
    Select Case kind
        Case "異動": RequireEmpty r, Array(5, 6)
        Case "役職変更": RequireEmpty r, Array(4)
        Case Else: RequireEmpty r, Array(4, 5, 6)
    End Select
    ReadBulkChange = ChangeRecord(r.Cells(1, 2).Value2, kind, r.Cells(1, 3).Value2, r.Cells(1, 4).Value2, r.Cells(1, 5).Value2, r.Cells(1, 6).Value2, r.Cells(1, 11).Value2, pending, stagedPeople)
End Function
Private Function TryBulkNew(ByVal r As Object, ByVal pending As Object, ByRef rec As Variant) As String
    On Error GoTo Invalid
    rec = ReadBulkNew(r, pending)
    Exit Function
Invalid:
    TryBulkNew = Err.Description
End Function
Private Function TryBulkChange(ByVal r As Object, ByVal pending As Object, ByVal stagedPeople As Object, ByRef rec As Variant) As String
    On Error GoTo Invalid
    rec = ReadBulkChange(r, pending, stagedPeople)
    Exit Function
Invalid:
    TryBulkChange = Err.Description
End Function
Private Function CheckBulkHistory(ByVal records As Collection, ByVal sourceRows As Collection) As Boolean
    Dim i As Long, j As Long, a As Variant, b As Variant
    CheckBulkHistory = True
    For i = 1 To records.Count
        a = records(i)
        If a(2) = "退職" Then
            For j = 1 To records.Count
                If i <> j Then
                    b = records(j)
                    If a(1) = b(1) And b(3) >= a(3) Then
                        sourceRows(i).Cells(1, 12).Value = "退職日以降の変更が同じ一覧にあります。"
                        sourceRows(j).Cells(1, 12).Value = "同じ一覧の退職日以降の変更です。"
                        CheckBulkHistory = False
                    End If
                End If
            Next j
        End If
    Next i
End Function
Private Function TrimAddedRows(ByVal t As Object, ByVal oldCount As Long) As Boolean
    Dim i As Long
    On Error GoTo Failed
    t.Parent.Unprotect PWD
    For i = t.ListRows.Count To oldCount + 1 Step -1
        t.ListRows(i).Delete
    Next i
    t.Parent.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    TrimAddedRows = (t.ListRows.Count = oldCount)
    Exit Function
Failed:
    On Error Resume Next
    t.Parent.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    TrimAddedRows = False
End Function
Private Sub AppendMixedRecords(ByVal people As Collection, ByVal changes As Collection)
    Dim employees As Object, history As Object, oldPeople As Long, oldHistory As Long
    Dim message As String, peopleRestored As Boolean, historyRestored As Boolean
    Set employees = Tbl("tEmployees"): Set history = Tbl("tHistory")
    oldPeople = employees.ListRows.Count: oldHistory = history.ListRows.Count
    ValidateBatchHistory changes
    On Error GoTo Rollback
    If people.Count > 0 Then AppendRecords "tEmployees", people
    If changes.Count > 0 Then AppendRecords "tHistory", changes
    Exit Sub
Rollback:
    message = Err.Description
    historyRestored = TrimAddedRows(history, oldHistory)
    peopleRestored = TrimAddedRows(employees, oldPeople)
    If historyRestored And peopleRestored Then
        Fail "新規・変更の登録をすべて取り消しました。" & message
    Else
        Fail "登録中にエラーが発生し、完全に元へ戻せませんでした。再登録せず、保守担当へ連絡してください。" & message
    End If
End Sub
Public Sub SaveBulkUnified()
    Dim inputTable As Object, row As Object, r As Object, pending As Object, stagedPeople As Object
    Dim people As New Collection, changes As New Collection, changeRows As New Collection
    Dim rec As Variant, message As String, label As String, kind As String, errorCount As Long
    Dim committed As Boolean, historyOK As Boolean
    On Error GoTo Fatal
    Set inputTable = Tbl("tBulkUnified")
    Set pending = CreateObject("Scripting.Dictionary")
    Set stagedPeople = CreateObject("Scripting.Dictionary")
    Busy = True: inputTable.Parent.Unprotect PWD
    inputTable.ListColumns(12).DataBodyRange.ClearContents
    ' 新規行を先に検証し、保存せずに氏名とIDを用意します。行の並び順には依存しません。
    For Each row In inputTable.ListRows
        Set r = row.Range
        If HasBulkInput(r) Then
            kind = Txt(r.Cells(1, 1).Value2)
            If kind = "新規" Then
                message = TryBulkNew(r, pending, rec)
                If message <> "" Then
                    errorCount = errorCount + 1: r.Cells(1, 12).Value = message
                Else
                    people.Add rec
                    label = rec(1): If rec(3) <> "" Then label = label & " (" & rec(3) & ")"
                    stagedPeople.Add LCase(label), rec
                    r.Cells(1, 12).Value = "登録待ち"
                End If
            ElseIf kind = "" Then
                errorCount = errorCount + 1: r.Cells(1, 12).Value = "登録区分を選んでください。"
            End If
        End If
    Next row
    For Each row In inputTable.ListRows
        Set r = row.Range
        If HasBulkInput(r) Then
            kind = Txt(r.Cells(1, 1).Value2)
            If kind <> "新規" And kind <> "" Then
                message = TryBulkChange(r, pending, stagedPeople, rec)
                If message <> "" Then
                    errorCount = errorCount + 1: r.Cells(1, 12).Value = message
                Else
                    changes.Add rec: changeRows.Add r: r.Cells(1, 12).Value = "登録待ち"
                End If
            End If
        End If
    Next row
    historyOK = CheckBulkHistory(changes, changeRows)
    If errorCount > 0 Or Not historyOK Then
        inputTable.ListColumns(12).DataBodyRange.WrapText = True
        inputTable.DataBodyRange.Rows.AutoFit
        MsgBox "不備のある行を確認結果に表示しました。今回は1件も登録していません。修正してから再実行してください。", vbExclamation
        GoTo Finish
    End If
    If people.Count + changes.Count = 0 Then MsgBox "入力した行がありません。", vbInformation: GoTo Finish
    If MsgBox("新規 " & people.Count & "名、異動などの変更 " & changes.Count & "件をまとめて登録します。", vbOKCancel + vbQuestion) <> vbOK Then GoTo Finish
    AppendMixedRecords people, changes
    committed = True
    inputTable.DataBodyRange.ClearContents
    inputTable.DataBodyRange.RowHeight = 24
    RefreshChoices
    MsgBox "新規と変更をまとめて登録しました。月別表を更新してください。", vbInformation
Finish:
    On Error Resume Next
    inputTable.Parent.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    Busy = False
    Exit Sub
Fatal:
    message = Err.Description
    If committed Then message = "台帳への登録は完了しています。重複登録しないでください。" & vbCrLf & message
    MsgBox message, vbExclamation, "一括登録"
    GoTo Finish
End Sub
Public Sub ConfigureBulkRows(Optional ByVal changed As Variant)
    Dim t As Object, scope As Object, row As Object, cell As Object, s As Object, wasBusy As Boolean
    Dim message As String, fullRefresh As Boolean
    wasBusy = Busy
    On Error GoTo Failed
    Set t = Tbl("tBulkUnified"): Set s = t.Parent
    fullRefresh = IsMissing(changed)
    If fullRefresh Then Set scope = t.DataBodyRange Else Set scope = Application.Intersect(changed, t.DataBodyRange)
    If scope Is Nothing Then Exit Sub
    Busy = True: s.Unprotect PWD
    For Each row In scope.Rows
        Set cell = s.Cells(row.Row, 3)
        cell.Validation.Delete
        If Txt(s.Cells(row.Row, 2).Value2) <> "新規" Then
            cell.Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=EmployeeChoices"
            cell.Validation.InCellDropdown = True
            cell.Validation.ShowError = False
            cell.Validation.InputTitle = "氏名の選択"
            cell.Validation.InputMessage = "一覧から選択します。同じ一括登録の新規従業員は氏名を直接入力できます。"
            cell.Validation.ShowInput = True
        End If
        If Not fullRefresh Then s.Cells(row.Row, 13).ClearContents
    Next row
    s.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    Busy = wasBusy
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    s.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    Busy = wasBusy
    MsgBox "氏名の選択欄を更新できませんでした。" & message, vbExclamation
End Sub
Public Sub ToggleBulkDetails()
    Dim s As Object
    On Error GoTo Failed
    Set s = Ws("一括登録"): s.Unprotect PWD
    s.Columns("I:L").Hidden = Not s.Columns("I:L").Hidden
    s.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

Private Function TryTargetRecord(ByVal r As Object, ByVal seen As Object, ByRef rec As Variant) As String
    Dim d As Date, sid As String, value As Double, key As String
    On Error GoTo Invalid
    d = MonthValue(r.Cells(1, 1).Value2, "対象月")
    sid = Resolve(Txt(r.Cells(1, 2).Value2), "tStores")
    If Txt(r.Cells(1, 3).Value2) = "" Then Fail "目標値を数値で入力してください。"
    If Not IsNumeric(r.Cells(1, 3).Value2) Then Fail "目標値を数値で入力してください。"
    value = CDbl(r.Cells(1, 3).Value2)
    If value < 0 Then Fail "目標値は0以上にしてください。"
    key = CStr(CLng(d)) & "|" & sid
    If seen.Exists(key) Then Fail "同じ月・店舗が重複しています。"
    seen.Add key, True
    rec = Array(d, sid, value)
    Exit Function
Invalid:
    TryTargetRecord = Err.Description
End Function
Private Function RestoreTargets(ByVal dest As Object, ByVal oldCount As Long, ByVal oldValues As Variant) As Boolean
    On Error GoTo Failed
    If Not TrimAddedRows(dest, oldCount) Then Exit Function
    dest.Parent.Unprotect PWD
    If oldCount > 0 Then dest.DataBodyRange.Value2 = oldValues
    dest.Parent.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    RestoreTargets = True
    Exit Function
Failed:
    On Error Resume Next
    dest.Parent.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
End Function
Public Sub SaveTargets()
    Dim inputTable As Object, dest As Object, row As Object, existing As Object, rec As Variant
    Dim records As New Collection, seen As Object, oldCount As Long, oldValues As Variant
    Dim found As Boolean, committed As Boolean, message As String, bad As Long
    On Error GoTo Fatal
    If Txt(Ws("設定").Range("B4").Value2) <> "店舗に直接入力" Then Fail "このブックは個人目標を合算する設定です。店舗目標の入力は不要です。"
    Set inputTable = Tbl("tTargetInput"): Set dest = Tbl("tStoreTargets")
    Set seen = CreateObject("Scripting.Dictionary")
    Busy = True: inputTable.Parent.Unprotect PWD
    For Each row In inputTable.ListRows
        row.Range.Cells(1, 4).ClearContents
        If Application.CountA(row.Range.Resize(1, 3)) > 0 Then
            message = TryTargetRecord(row.Range, seen, rec)
            If message <> "" Then
                bad = bad + 1: row.Range.Cells(1, 4).Value = message
            Else
                records.Add rec: row.Range.Cells(1, 4).Value = "登録待ち"
            End If
        End If
    Next row
    If bad > 0 Then MsgBox "不備のある行を直してください。今回は登録していません。", vbExclamation: GoTo Finish
    If records.Count = 0 Then MsgBox "入力した行がありません。", vbInformation: GoTo Finish
    If MsgBox(records.Count & "件の店舗目標を保存します。同じ月・店舗の登録済み目標は置き換わります。", vbOKCancel + vbQuestion) <> vbOK Then GoTo Finish
    oldCount = dest.ListRows.Count
    If oldCount > 0 Then oldValues = dest.DataBodyRange.Value2
    On Error GoTo Rollback
    dest.Parent.Unprotect PWD
    For Each rec In records
        found = False
        For Each existing In dest.ListRows
            If Txt(existing.Range.Cells(1, 1).Value2) <> "" Then
                If CLng(existing.Range.Cells(1, 1).Value2) = CLng(rec(0)) And Txt(existing.Range.Cells(1, 2).Value2) = rec(1) Then
                    existing.Range.Cells(1, 3).Value2 = rec(2): found = True: Exit For
                End If
            End If
        Next existing
        If Not found Then
            Set existing = dest.ListRows.Add
            existing.Range.Cells(1, 1).Value = rec(0)
            existing.Range.Cells(1, 2).NumberFormat = "@"
            existing.Range.Cells(1, 2).Value = rec(1)
            existing.Range.Cells(1, 3).Value = rec(2)
        End If
    Next rec
    MarkDirty
    committed = True
    On Error GoTo Fatal
    inputTable.DataBodyRange.ClearContents
    MsgBox "店舗目標を保存しました。月別表を更新してください。", vbInformation
Finish:
    On Error Resume Next
    inputTable.Parent.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    dest.Parent.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    Busy = False: Exit Sub
Rollback:
    message = Err.Description
    If RestoreTargets(dest, oldCount, oldValues) Then
        MsgBox "保存を取り消しました。" & message, vbExclamation
    Else
        MsgBox "保存中にエラーが発生し、完全に元へ戻せませんでした。再登録せず、保守担当へ連絡してください。" & message, vbExclamation
    End If
    GoTo Finish
Fatal:
    message = Err.Description
    If committed Then message = "店舗目標の保存は完了しています。保存内容を確認してから操作してください。" & vbCrLf & message
    MsgBox message, vbExclamation: GoTo Finish
End Sub
Public Sub MarkDirty()
    Dim previous As Boolean: previous = Busy: Busy = True
    Ws("ホーム").Range("D14").Value = "未更新 - 月別表を更新してください"
    Busy = previous
End Sub
Public Sub RefreshChoices()
    Dim s As Object, row As Object, n As Long, col As Long, tableName As String, name As String
    Set s = Ws("_候補"): s.Unprotect PWD: s.Range("A2:C20000").ClearContents
    For col = 1 To 3
        Select Case col
            Case 1: tableName = "tEmployees": name = "EmployeeChoices"
            Case 2: tableName = "tStores": name = "StoreChoices"
            Case 3: tableName = "tRoles": name = "RoleChoices"
        End Select
        n = 1
        For Each row In Tbl(tableName).ListRows
            If Txt(row.Range.Cells(1, 1).Value2) <> "" Then
                n = n + 1
                If col = 1 Then s.Cells(n, col).Value = DisplayPerson(row.Range) Else s.Cells(n, col).Value = Txt(row.Range.Cells(1, 2).Value2)
            End If
        Next row
        On Error Resume Next: ThisWorkbook.Names(name).Delete: On Error GoTo 0
        ThisWorkbook.Names.Add Name:=name, RefersTo:="='_候補'!" & s.Range(s.Cells(2, col), s.Cells(Application.Max(2, n), col)).Address
    Next col
    s.Protect Password:=PWD, UserInterfaceOnly:=True
End Sub
Public Sub UpdateChoices()
    On Error GoTo Failed
    Busy = True: RefreshChoices: Busy = False
    MsgBox "氏名・店舗・役職の候補を更新しました。", vbInformation: Exit Sub
Failed:
    Busy = False: MsgBox Err.Description, vbExclamation
End Sub
Public Sub ConfigureChange()
    Dim s As Object, kind As String
    Set s = Ws("変更登録"): kind = Txt(s.Range("D8").Value2)
    s.Rows("12:12").Hidden = (kind <> "異動")
    s.Rows("14:16").Hidden = (kind <> "役職変更")
End Sub
Private Function QueryExists(ByVal name As String) As Boolean
    Dim q As Object
    For Each q In ThisWorkbook.Queries
        If q.name = name Then QueryExists = True: Exit Function
    Next q
End Function
Private Sub EnsureQueries()
    Dim code As String, i As Long, s As Object
    If Not QueryExists("qEmployeeCore") Then
        Set s = Ws("_PowerQuery")
        For i = 2 To s.Cells(s.Rows.Count, 1).End(-4162).Row: code = code & s.Cells(i, 1).Value2 & vbCrLf: Next i
        ThisWorkbook.Queries.Add Name:="qEmployeeCore", Formula:=code
    End If
    If Not QueryExists("qEmployees") Then ThisWorkbook.Queries.Add Name:="qEmployees", Formula:="let Source = qEmployeeCore[Employees] in Source"
    If Not QueryExists("qStoreTargets") Then ThisWorkbook.Queries.Add Name:="qStoreTargets", Formula:="let Source = qEmployeeCore[Stores] in Source"
End Sub
Private Function BindQuery(ByVal sheetName As String, ByVal tableName As String, ByVal queryName As String) As Object
    Dim s As Object, t As Object, qt As Object, connection As String
    Set s = Ws(sheetName): s.Unprotect PWD
    On Error Resume Next
    Set t = s.ListObjects(tableName): Set qt = t.QueryTable
    On Error GoTo 0
    If qt Is Nothing Then
        If Not t Is Nothing Then t.Unlist
        s.Range("A6:Q" & Application.Max(7, s.UsedRange.Rows.Count + s.UsedRange.Row)).ClearContents
        connection = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties="""""
        Set t = s.ListObjects.Add(SourceType:=0, Source:=connection, Destination:=s.Range("A6"))
        t.name = tableName: t.TableStyle = "TableStyleMedium2"
        Set qt = t.QueryTable
        qt.CommandType = 2: qt.CommandText = "SELECT * FROM [" & queryName & "]"
        qt.BackgroundQuery = False: qt.AdjustColumnWidth = False: qt.PreserveFormatting = True
    End If
    Set BindQuery = qt
End Function
Public Sub RefreshData()
    Dim endMonth As Date, qt As Object, qt2 As Object, message As String
    On Error GoTo Failed
    endMonth = MonthValue(Ws("ホーム").Range("D5").Value2, "集計対象月")
    If endMonth < MonthValue(Ws("設定").Range("B5").Value2, "集計開始月") Then Fail "集計対象月が集計開始月より前です。"
    Busy = True: Application.ScreenUpdating = False
    Ws("設定").Unprotect PWD: Ws("設定").Range("B6").Value = endMonth
    Ws("ホーム").Range("D14").Value = "更新中..."
    RefreshChoices: EnsureQueries
    Set qt = BindQuery("月別従業員", "tMonthlyEmployees", "qEmployees")
    If Not qt.Refresh(BackgroundQuery:=False) Then Fail "従業員表の更新が完了しませんでした。"
    Set qt2 = BindQuery("月別店舗目標", "tMonthlyStores", "qStoreTargets")
    If Not qt2.Refresh(BackgroundQuery:=False) Then Fail "店舗目標の更新が完了しませんでした。"
    Tbl("tMonthlyEmployees").ListColumns(1).Range.NumberFormat = "yyyy/mm"
    Tbl("tMonthlyEmployees").ListColumns(12).Range.NumberFormat = "yyyy/mm/dd"
    Tbl("tMonthlyEmployees").ListColumns(7).Range.NumberFormat = "0.##"
    Tbl("tMonthlyEmployees").ListColumns(8).Range.NumberFormat = "0.##"
    Tbl("tMonthlyStores").ListColumns(1).Range.NumberFormat = "yyyy/mm"
    Tbl("tMonthlyStores").ListColumns(7).Range.NumberFormat = "0.##"
    Ws("ホーム").Range("D18").Formula = "=COUNTIFS(tMonthlyEmployees[実績表用日程],D5)"
    Ws("ホーム").Range("H18").Formula = "=SUMIFS(tMonthlyStores[店舗目標],tMonthlyStores[実績表用日程],D5)"
    Ws("ホーム").Range("D14").Value = "更新完了  " & Format(Now, "yyyy/mm/dd hh:nn")
    Ws("ホーム").Range("D15").Value = endMonth
    ProtectAll: Application.Calculate: Busy = False: Application.ScreenUpdating = True
    MsgBox "月別表を更新しました。", vbInformation: Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    Ws("ホーム").Range("D14").Value = "更新失敗 - 出力を使用しないでください"
    ProtectAll: Busy = False: Application.ScreenUpdating = True
    MsgBox "月別表を更新できませんでした。入力元を確認してください。" & vbCrLf & message, vbExclamation
End Sub
Public Sub InitializeBook()
    On Error GoTo Failed
    Busy = True: ProtectAll: RefreshChoices: ConfigureChange: ConfigureBulkRows: Busy = False
    Exit Sub
Failed:
    Busy = False: MsgBox "初期設定を確認してください。" & vbCrLf & Err.Description, vbExclamation
End Sub
Private Function IsAdminSheet(ByVal name As String) As Boolean
    IsAdminSheet = (Left(name, 2) = "台帳" Or Left(name, 2) = "旧月" Or name = "設定" Or Left(name, 1) = "_")
End Function
Public Sub ProtectAll()
    Dim s As Object
    For Each s In ThisWorkbook.Worksheets
        s.Unprotect PWD
        s.Protect Password:=PWD, UserInterfaceOnly:=True, AllowFiltering:=True
    Next s
End Sub
Public Sub ShowAdmin()
    Dim s As Object
    For Each s In ThisWorkbook.Worksheets
        If IsAdminSheet(s.name) Then
            If Left(s.name, 1) <> "_" Then s.Visible = -1: s.Unprotect PWD
        End If
    Next s
    Ws("設定").Activate
End Sub
Public Sub CloseAdmin()
    Dim s As Object
    Ws("ホーム").Activate: ProtectAll: RefreshChoices: MarkDirty
    For Each s In ThisWorkbook.Worksheets
        If IsAdminSheet(s.name) Then s.Visible = 2
    Next s
End Sub
Public Sub ClearSampleEmployees()
    Dim t As Object, name As Variant, backup As String
    On Error GoTo Failed
    If ThisWorkbook.Path = "" Then Fail "先にこのブックを保存してください。"
    If InputBox("従業員・変更履歴・店舗目標・旧月データを空にします。店舗・ブロック・役職の設定は残します。実行する場合は「初期化」と入力してください。", "本番利用の準備") <> "初期化" Then Exit Sub
    backup = ThisWorkbook.Path & Application.PathSeparator & "従業員マスター_初期化前_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs backup
    Busy = True
    For Each name In Array("tEmployees", "tHistory", "tStoreTargets", "tLegacyEmployees", "tLegacyStores", "tMonthlyEmployees", "tMonthlyStores")
        Set t = Tbl(CStr(name)): t.Parent.Unprotect PWD
        If Not t.DataBodyRange Is Nothing Then t.DataBodyRange.ClearContents
    Next name
    Ws("設定").Unprotect PWD: Ws("設定").Range("B7").Value = "本番"
    RefreshChoices: ProtectAll: MarkDirty: Busy = False
    MsgBox "初期化しました。初期化前のコピーを同じフォルダーに保存しました。店舗・役職・ポイントを自社用に設定してください。", vbInformation: Exit Sub
Failed:
    Busy = False: On Error Resume Next: ProtectAll: MsgBox Err.Description, vbExclamation
End Sub
Private Sub Navigate(ByVal name As String)
    Ws(name).Visible = -1: Ws(name).Activate: Ws(name).Range("B2").Select
End Sub
Public Sub GoHome(): Navigate "ホーム": End Sub
Public Sub GoChange(): Navigate "変更登録": End Sub
Public Sub GoNew(): Navigate "新規登録": End Sub
Public Sub GoBulk(): Navigate "一括登録": ConfigureBulkRows: End Sub
Public Sub GoBulkNew(): GoBulk: End Sub
Public Sub GoTargets(): Navigate "店舗目標入力": End Sub
Public Sub GoEmployees(): Navigate "月別従業員": End Sub
Public Sub GoStores(): Navigate "月別店舗目標": End Sub
