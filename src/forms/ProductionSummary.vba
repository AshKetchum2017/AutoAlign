Option Explicit

Private mReport As AAProductionReport

Public Sub LoadReport(ByVal report As AAProductionReport)
    Dim i As Long, column As Long
    Set mReport = report
    lblSummary.Caption = report.Summary
    txbDetails.MultiLine = True
    txbDetails.Locked = True
    txbDetails.ScrollBars = fmScrollBarsVertical
    txbDetails.text = report.Summary & vbCrLf & vbCrLf & report.Details
    lbxProduction.clear
    lbxProduction.MultiSelect = fmMultiSelectSingle
    lbxProduction.ColumnCount = 4
    lbxProduction.ColumnWidths = "45 pt;85 pt;85 pt;85 pt"
    For i = 1 To report.RowCount
        lbxProduction.AddItem report.CellText(i, 0)
        For column = 1 To 3
            lbxProduction.List(i - 1, column) = report.CellText(i, column)
        Next column
    Next i
End Sub

Private Sub cmdCopy_Click()
    Dim clipboard As MSForms.DataObject
    If mReport Is Nothing Then Exit Sub
    On Error GoTo CopyFailed
    Set clipboard = New MSForms.DataObject
    clipboard.SetText mReport.CopyText
    clipboard.PutInClipboard
    Exit Sub
CopyFailed:
    MsgBox "Ringkasan tidak dapat disalin: " & Err.description, vbExclamation, "Production Summary"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Set mReport = Nothing
End Sub
