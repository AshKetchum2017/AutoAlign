Option Explicit

Private mReport As AAInspectionReport

Public Sub LoadReport(ByVal report As AAInspectionReport)
    Dim i As Long, item As AAInspectionIssue, areaText As String
    Set mReport = report
    lblSummary.Caption = report.Summary
    txbDetail.MultiLine = True
    txbDetail.Locked = True
    txbDetail.ScrollBars = fmScrollBarsVertical
    lbxIssues.Clear
    lbxIssues.MultiSelect = fmMultiSelectSingle
    lbxIssues.ColumnCount = 5
    lbxIssues.ColumnWidths = "85 pt;45 pt;70 pt;70 pt;250 pt"
    For i = 1 To report.Count
        Set item = report.Issue(i)
        areaText = "-"
        If item.AreaNumber > 0 Then areaText = CStr(item.AreaNumber)
        lbxIssues.AddItem item.Category
        lbxIssues.List(i - 1, 1) = areaText
        If item.FirstID > 0 Then lbxIssues.List(i - 1, 2) = CStr(item.FirstID)
        If item.SecondID > 0 Then lbxIssues.List(i - 1, 3) = CStr(item.SecondID)
        lbxIssues.List(i - 1, 4) = item.Detail
    Next i
    If report.Count > 0 Then lbxIssues.ListIndex = 0
    RefreshDetail
End Sub

Private Sub lbxIssues_Click()
    RefreshDetail
End Sub

Private Sub RefreshDetail()
    Dim item As AAInspectionIssue
    cmdSelectIssue.Enabled = False
    If mReport Is Nothing Then Exit Sub
    ' Ringkasan harus tetap terbaca walau label pada layout form terpotong.
    txbDetail.Text = mReport.Summary & vbCrLf & vbCrLf & mReport.Context
    If lbxIssues.ListIndex < 0 Then Exit Sub
    Set item = mReport.Issue(lbxIssues.ListIndex + 1)
    txbDetail.Text = mReport.Summary & vbCrLf & vbCrLf & item.Category & vbCrLf & _
        item.Detail & vbCrLf & vbCrLf & mReport.Context
    cmdSelectIssue.Enabled = (item.FirstID > 0 Or item.SecondID > 0)
End Sub

Private Sub cmdSelectIssue_Click()
    If mReport Is Nothing Or lbxIssues.ListIndex < 0 Then Exit Sub
    On Error GoTo Failed
    mReport.SelectIssue lbxIssues.ListIndex + 1
    Exit Sub
Failed:
    MsgBox "Objek tidak dapat dipilih: " & Err.Description, vbExclamation, "Inspection Results"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Set mReport = Nothing
End Sub
