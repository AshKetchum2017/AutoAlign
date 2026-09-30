Option Explicit

Private mPresenter As AAPresenter

Private Sub lbxObjects_Click()

End Sub

Private Sub txbAreaHeight_Change()

End Sub

Private Sub txbAreaWidth_Change()

End Sub

Private Sub txbGapHorizontal_Change()

End Sub

Private Sub txbGapVertical_Change()

End Sub

Private Sub UserForm_Initialize()
    Set mPresenter = New AAPresenter
    chkAuto.Value = False
    chkAuto.Enabled = True
    optStraight.Value = False
    optZigZag.Value = False
    optDense.Value = True
    AAUpdateModeControls
    txbAreaHeight.Enabled = False
    txbAreaWidth.Enabled = False
    txbGapHorizontal.Enabled = False
    txbGapVertical.Enabled = False
    AARefreshObjects
End Sub

Private Sub chkAuto_Click()
    AAUpdateModeControls
End Sub

Private Sub cmdClear_Click()
    mPresenter.ClearObjects
    AARefreshObjects
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Click()
    Dim layoutMode As Long

    If chkAuto.Value Then
        MsgBox "Mode Auto belum tersedia.", vbExclamation, "Auto Align"
        Exit Sub
    End If

    If optDense.Value Then
        layoutMode = AA_MODE_DENSE
    ElseIf optStraight.Value Then
        layoutMode = AA_MODE_STRAIGHT
    ElseIf optZigZag.Value Then
        layoutMode = AA_MODE_ZIGZAG
    Else
        MsgBox "Pilih mode penyusunan terlebih dahulu.", vbExclamation, "Auto Align"
        Exit Sub
    End If

    AARunMode layoutMode
End Sub

Private Sub cmdRemove_Click()
    Dim i As Long

    For i = lbxObjects.ListCount - 1 To 0 Step -1
        If lbxObjects.Selected(i) Then mPresenter.RemoveObjectAt i
    Next i
    AARefreshObjects
End Sub

Private Sub cmdSetCutLine_Click()
    AARegisterObjects AA_ROLE_CUT_LINE
End Sub

Private Sub cmdSetDesign_Click()
    AARegisterObjects AA_ROLE_DESIGN
End Sub

Private Sub optDense_Click()
    If Not optDense.Value Then Exit Sub
    optStraight.Value = False
    optZigZag.Value = False
End Sub

Private Sub optStraight_Click()
    If Not optStraight.Value Then Exit Sub
    optDense.Value = False
    optZigZag.Value = False
End Sub

Private Sub optZigZag_Click()
    If Not optZigZag.Value Then Exit Sub
    optDense.Value = False
    optStraight.Value = False
End Sub

Private Sub AAUpdateModeControls()
    optDense.Enabled = Not chkAuto.Value
    optStraight.Enabled = Not chkAuto.Value
    optZigZag.Enabled = Not chkAuto.Value
    cmdProcess.Enabled = Not chkAuto.Value
End Sub

Private Sub AARunMode(ByVal layoutMode As Long)
    Dim feedback As String
    Dim feedbackStyle As VbMsgBoxStyle

    If mPresenter Is Nothing Then Set mPresenter = New AAPresenter
    mPresenter.RunLayout layoutMode, feedback, feedbackStyle
    AARefreshObjects
    If Len(feedback) > 0 Then MsgBox feedback, feedbackStyle, "Auto Align"
End Sub

Private Sub AARegisterObjects(ByVal roleLabel As String)
    Dim feedback As String
    Dim feedbackStyle As VbMsgBoxStyle

    mPresenter.RegisterSelected roleLabel, feedback, feedbackStyle
    AARefreshObjects
    If Len(feedback) > 0 Then MsgBox feedback, feedbackStyle, "Auto Align"
End Sub

Private Sub AARefreshObjects()
    Dim i As Long

    lbxObjects.Clear
    For i = 0 To mPresenter.ObjectCount - 1
        lbxObjects.AddItem mPresenter.ObjectListText(i)
    Next i
End Sub
