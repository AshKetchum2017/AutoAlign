Option Explicit

Private mPresenter As AAPresenter

Private Sub UserForm_Initialize()
    Set mPresenter = New AAPresenter
    cmdProcess.Enabled = False
    cmdAuto.Enabled = False
    txbGap.Enabled = False
    AARefreshObjects
End Sub

Private Sub chkSequentially_Click()

End Sub

Private Sub txbGap_Change()

End Sub

Private Sub cmdAuto_Click()
    ' Scope berikutnya: Auto memilih AA_MODE_STRAIGHT atau AA_MODE_ZIGZAG.
End Sub

Private Sub cmdClear_Click()
    mPresenter.ClearObjects
    AARefreshObjects
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDense_Click()
    AARunMode AA_MODE_DENSE
End Sub

Private Sub cmdProcess_Click()
    ' Tiap mode saat ini dijalankan langsung dari Command Button.
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

Private Sub cmdStraight_Click()
    AARunMode AA_MODE_STRAIGHT
End Sub

Private Sub cmdZigZag_Click()
    AARunMode AA_MODE_ZIGZAG
End Sub

Private Sub lbxObjects_Click()

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
