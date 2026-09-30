Option Explicit

Private mPresenter As AAPresenter

Private Sub lbxObjects_Click()

End Sub

Private Sub txbAreaHeight_Change()

End Sub

Private Sub txbAreaWidth_Change()

End Sub

Private Sub txbQuantity_Change()

End Sub

Private Sub txbGapHorizontal_Change()
    If Not chkHPercent.Value Then Exit Sub
    If AAGapIsNegative(txbGapHorizontal.Text) Then
        MsgBox "Persen gap horizontal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkHPercent.Value = False
        AASetGapPercentText txbGapHorizontal, False
        Exit Sub
    End If
    AASetGapPercentText txbGapHorizontal, True
End Sub

Private Sub txbGapVertical_Change()
    If Not chkVPercent.Value Then Exit Sub
    If AAGapIsNegative(txbGapVertical.Text) Then
        MsgBox "Persen gap vertikal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkVPercent.Value = False
        AASetGapPercentText txbGapVertical, False
        Exit Sub
    End If
    AASetGapPercentText txbGapVertical, True
End Sub

Private Sub chkHPercent_Click()
    If chkHPercent.Value And AAGapIsNegative(txbGapHorizontal.Text) Then
        MsgBox "Persen gap horizontal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkHPercent.Value = False
        Exit Sub
    End If
    AASetGapPercentText txbGapHorizontal, chkHPercent.Value
End Sub

Private Sub chkVPercent_Click()
    If chkVPercent.Value And AAGapIsNegative(txbGapVertical.Text) Then
        MsgBox "Persen gap vertikal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkVPercent.Value = False
        Exit Sub
    End If
    AASetGapPercentText txbGapVertical, chkVPercent.Value
End Sub

Private Sub UserForm_Initialize()
    Set mPresenter = New AAPresenter
    chkAuto.Value = False
    chkAuto.Enabled = True
    optStraight.Value = False
    optZigZag.Value = False
    optDense.Value = True
    AAUpdateModeControls
    txbAreaHeight.Enabled = True
    txbAreaWidth.Enabled = True
    txbGapHorizontal.Enabled = True
    txbGapVertical.Enabled = True
    chkHPercent.Enabled = True
    chkVPercent.Enabled = True
    AASetGapPercentText txbGapHorizontal, chkHPercent.Value
    AASetGapPercentText txbGapVertical, chkVPercent.Value
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
        layoutMode = AA_MODE_AUTO
    ElseIf optDense.Value Then
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
    cmdProcess.Enabled = True
End Sub

Private Function AAGapIsNegative(ByVal gapText As String) As Boolean
    AAGapIsNegative = (Left$(Trim$(gapText), 1) = "-")
End Function

Private Sub AASetGapPercentText(ByVal gapBox As MSForms.TextBox, _
                                ByVal percentEnabled As Boolean)
    Dim numberText As String
    Dim displayText As String
    Dim cursorPosition As Long

    cursorPosition = gapBox.SelStart
    numberText = Trim$(Replace$(gapBox.Text, "%", vbNullString))
    displayText = numberText
    If percentEnabled And Len(numberText) > 0 Then displayText = numberText & "%"
    If gapBox.Text = displayText Then Exit Sub
    gapBox.Text = displayText
    If cursorPosition > Len(numberText) Then cursorPosition = Len(numberText)
    gapBox.SelStart = cursorPosition
End Sub

Private Sub AARunMode(ByVal layoutMode As Long)
    Dim feedback As String
    Dim feedbackStyle As VbMsgBoxStyle

    If mPresenter Is Nothing Then Set mPresenter = New AAPresenter
    mPresenter.RunLayout layoutMode, txbAreaWidth.Text, txbAreaHeight.Text, _
        txbGapHorizontal.Text, txbGapVertical.Text, chkHPercent.Value, _
        chkVPercent.Value, feedback, feedbackStyle
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
