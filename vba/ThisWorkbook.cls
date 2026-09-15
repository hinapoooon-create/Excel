Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Workbook_Open()
    EmployeeMaster.InitializeBook
End Sub
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If EmployeeMaster.Busy Then Exit Sub
    On Error GoTo Done
    If Sh.Name = "変更登録" Then
        If Not Intersect(Target, Sh.Range("D8")) Is Nothing Then EmployeeMaster.ConfigureChange
    End If
    If Sh.Name = "一括登録" Then EmployeeMaster.ConfigureBulkRows Target
    If Left(Sh.Name, 2) = "台帳" Or Left(Sh.Name, 2) = "旧月" Or Sh.Name = "設定" Then EmployeeMaster.MarkDirty
    If Sh.Name = "ホーム" Then
        If Not Intersect(Target, Sh.Range("D5")) Is Nothing Then EmployeeMaster.MarkDirty
    End If
Done:
End Sub
