Option Explicit

Private mPresenter As AAPresenter

Private Sub UserForm_Initialize()
    Set mPresenter = New AAPresenter
    cmdProcess.Enabled = False
    cmdAuto.Enabled = False
End Sub

Private Sub chkSequentially_Click()

End Sub

Private Sub cmdAuto_Click()
    ' Scope berikutnya: Auto memilih AA_MODE_STRAIGHT atau AA_MODE_ZIGZAG.
End Sub

Private Sub cmdClear_Click()

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

End Sub

Private Sub cmdSetCutLine_Click()

End Sub

Private Sub cmdSetDesign_Click()

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
    If Len(feedback) > 0 Then MsgBox feedback, feedbackStyle, "Auto Align"
End Sub
