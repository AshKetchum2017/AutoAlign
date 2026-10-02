Option Explicit

Private mPresenter As AAPresenter
Private mUpdatingQuantity As Boolean
Private mQuantityIndex As Long
Private mHintReady As Boolean
Private mUpdatingHint As Boolean
Private mHintActive(1 To 5) As Boolean
Private mInputColor(1 To 5) As Long
Private mFocusedInput As String

Private Sub lbxObjects_Click()
    AAUpdateQuantityEditor
End Sub

Private Sub optCustom_Click()
    If optCustom.Value Then
        optKissA.Value = False
        optDieA.Value = False
    End If
    AAUpdateModelControls
End Sub

Private Sub optDieA_Click()
    If optDieA.Value Then
        optCustom.Value = False
        optKissA.Value = False
    End If
    AAUpdateModelControls
End Sub

Private Sub optKissA_Click()
    If optKissA.Value Then
        optCustom.Value = False
        optDieA.Value = False
    End If
    AAUpdateModelControls
End Sub

Private Sub cmdOptimize_Click()
' One of the last scopes of AutoAlign is to optimize the existing group. This button triggers the optimization process.
End Sub

Private Sub optMaximum_Click()

End Sub

Private Sub optMedium_Click()

End Sub

Private Sub txbGapHorizontal_Change()
    If mUpdatingHint Then Exit Sub
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
    If mUpdatingHint Then Exit Sub
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
    If chkHPercent.Value And AAGapIsNegative(AAInputText(txbGapHorizontal)) Then
        MsgBox "Persen gap horizontal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkHPercent.Value = False
        Exit Sub
    End If
    AASetGapPercentText txbGapHorizontal, chkHPercent.Value
End Sub

Private Sub chkVPercent_Click()
    If chkVPercent.Value And AAGapIsNegative(AAInputText(txbGapVertical)) Then
        MsgBox "Persen gap vertikal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkVPercent.Value = False
        Exit Sub
    End If
    AASetGapPercentText txbGapVertical, chkVPercent.Value
End Sub

Private Sub txbQuantity_Change()
    Dim previousText As String

    If mUpdatingHint Or mUpdatingQuantity Or mQuantityIndex < 0 Then Exit Sub
    previousText = mPresenter.ObjectQuantityTextAt(mQuantityIndex)
    On Error GoTo InvalidQuantity
    mPresenter.SetObjectQuantityText mQuantityIndex, AAInputText(txbQuantity)
    lbxObjects.List(mQuantityIndex) = mPresenter.ObjectListText(mQuantityIndex)
    Exit Sub
InvalidQuantity:
    MsgBox Err.Description, vbExclamation, "Auto Align"
    mUpdatingQuantity = True
    txbQuantity.Text = previousText
    mUpdatingQuantity = False
End Sub

Private Sub UserForm_Initialize()
    Set mPresenter = New AAPresenter
    mQuantityIndex = -1
    mInputColor(1) = txbGapHorizontal.ForeColor
    mInputColor(2) = txbGapVertical.ForeColor
    mInputColor(3) = txbAreaWidth.ForeColor
    mInputColor(4) = txbAreaHeight.ForeColor
    mInputColor(5) = txbQuantity.ForeColor
    mHintReady = True
    optCustom.GroupName = "AAModels"
    optKissA.GroupName = "AAModels"
    optDieA.GroupName = "AAModels"
    optMinimum.GroupName = "AAModes"
    optMedium.GroupName = "AAModes"
    optMaximum.GroupName = "AAModes"
    optKissA.Value = False
    optDieA.Value = False
    optCustom.Value = True
    optMinimum.Value = True
    optMedium.Enabled = False
    optMaximum.Enabled = False
    AAUpdateModelControls
    txbGapHorizontal.Enabled = True
    txbGapVertical.Enabled = True
    chkHPercent.Enabled = True
    chkVPercent.Enabled = True
    AASetGapPercentText txbGapHorizontal, chkHPercent.Value
    AASetGapPercentText txbGapVertical, chkVPercent.Value
    AARefreshObjects
End Sub

Private Sub cmdClear_Click()
    mPresenter.ClearObjects
    AARefreshObjects
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Click()
    Dim modelMode As Long

    If Not optMinimum.Value Then
        MsgBox "Pilih mode Minimum terlebih dahulu.", vbExclamation, "Auto Align"
        Exit Sub
    End If
    If optCustom.Value Then
        modelMode = AA_MODEL_CUSTOM
    ElseIf optKissA.Value Then
        modelMode = AA_MODEL_KISS_A
    ElseIf optDieA.Value Then
        modelMode = AA_MODEL_DIE_A
    Else
        MsgBox "Pilih model area terlebih dahulu.", vbExclamation, "Auto Align"
        Exit Sub
    End If

    AARunMode modelMode
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

Private Sub AAUpdateModelControls()
    txbAreaWidth.Enabled = optCustom.Value
    txbAreaHeight.Enabled = optCustom.Value
    cmdProcess.Enabled = True
    AARefreshHints
End Sub

Private Sub AAUpdateQuantityEditor()
    Dim i As Long
    Dim selectedCount As Long
    Dim selectedIndex As Long

    mQuantityIndex = -1
    For i = 0 To lbxObjects.ListCount - 1
        If lbxObjects.Selected(i) Then
            selectedCount = selectedCount + 1
            selectedIndex = i
        End If
    Next i
    If selectedCount = 1 Then
        If mPresenter.ObjectRoleAt(selectedIndex) = AA_ROLE_DESIGN Then _
            mQuantityIndex = selectedIndex
    End If
    mUpdatingQuantity = True
    If mHintReady Then
        mHintActive(5) = False
        txbQuantity.ForeColor = mInputColor(5)
    End If
    txbQuantity.Enabled = (mQuantityIndex >= 0)
    If mQuantityIndex >= 0 Then
        txbQuantity.Text = mPresenter.ObjectQuantityTextAt(mQuantityIndex)
    Else
        txbQuantity.Text = vbNullString
    End If
    mUpdatingQuantity = False
    AAShowHint txbQuantity
End Sub

Private Function AAGapIsNegative(ByVal gapText As String) As Boolean
    AAGapIsNegative = (Left$(Trim$(gapText), 1) = "-")
End Function

Private Sub AASetGapPercentText(ByVal gapBox As MSForms.TextBox, _
                                ByVal percentEnabled As Boolean)
    Dim numberText As String
    Dim displayText As String
    Dim cursorPosition As Long

    If mUpdatingHint Then Exit Sub
    If mHintReady Then
        If mHintActive(AAInputIndex(gapBox)) Then Exit Sub
    End If
    cursorPosition = gapBox.SelStart
    numberText = Trim$(Replace$(gapBox.Text, "%", vbNullString))
    displayText = numberText
    If percentEnabled And Len(numberText) > 0 Then displayText = numberText & "%"
    If gapBox.Text = displayText Then Exit Sub
    gapBox.Text = displayText
    If cursorPosition > Len(numberText) Then cursorPosition = Len(numberText)
    gapBox.SelStart = cursorPosition
End Sub

Private Sub AARunMode(ByVal modelMode As Long)
    Dim feedback As String
    Dim feedbackStyle As VbMsgBoxStyle

    If mPresenter Is Nothing Then Set mPresenter = New AAPresenter
    mPresenter.RunLayout modelMode, AAInputText(txbAreaWidth), AAInputText(txbAreaHeight), _
        AAInputText(txbGapHorizontal), AAInputText(txbGapVertical), chkHPercent.Value, _
        chkVPercent.Value, feedback, feedbackStyle
    AARefreshObjects
    If Len(feedback) > 0 Then MsgBox feedback, feedbackStyle, "Auto Align"
End Sub

' Placeholder hanya presentasi; AAInputText selalu mengembalikannya sebagai kosong.
Private Function AAInputIndex(ByVal box As MSForms.TextBox) As Long
    Select Case box.Name
        Case "txbGapHorizontal": AAInputIndex = 1
        Case "txbGapVertical": AAInputIndex = 2
        Case "txbAreaWidth": AAInputIndex = 3
        Case "txbAreaHeight": AAInputIndex = 4
        Case "txbQuantity": AAInputIndex = 5
    End Select
End Function

Private Function AAInputText(ByVal box As MSForms.TextBox) As String
    If mHintReady Then
        If mHintActive(AAInputIndex(box)) Then Exit Function
    End If
    AAInputText = box.Text
End Function

Private Function AAHintText(ByVal box As MSForms.TextBox) As String
    Select Case AAInputIndex(box)
        Case 1
            If optCustom.Value Then AAHintText = "0" Else AAHintText = "1"
        Case 2
            If optCustom.Value Then AAHintText = "0 / 88%" Else AAHintText = "1 / 88%"
        Case 3: AAHintText = "320"
        Case 4: AAHintText = "470"
        Case 5: AAHintText = "-"
    End Select
End Function

Private Sub AAShowHint(ByVal box As MSForms.TextBox)
    Dim index As Long
    If Not mHintReady Then Exit Sub
    If mFocusedInput = box.Name Then Exit Sub
    index = AAInputIndex(box)
    If Not mHintActive(index) And Len(box.Text) > 0 Then Exit Sub
    mUpdatingHint = True
    mHintActive(index) = True
    box.ForeColor = RGB(128, 128, 128)
    box.Text = AAHintText(box)
    mUpdatingHint = False
End Sub

Private Sub AARefreshHints()
    If Not mHintReady Then Exit Sub
    AAShowHint txbGapHorizontal
    AAShowHint txbGapVertical
    AAShowHint txbAreaWidth
    AAShowHint txbAreaHeight
    AAShowHint txbQuantity
    txbGapHorizontal.ControlTipText = "Default dalam mm. KissA/DieA: persegi atau persegi panjang bersudut runcing memakai 0. Input manual diutamakan."
    txbGapVertical.ControlTipText = "Default: Straight (mm) / ZigZag (persen tinggi). KissA/DieA: rectangle bersudut runcing memakai gap 0 pada kedua pola."
    txbQuantity.ControlTipText = "Kosong (-): isi sisa area. 0: lewati Design."
End Sub

Private Sub AAEnterInput(ByVal box As MSForms.TextBox)
    Dim index As Long
    If Not mHintReady Then Exit Sub
    mFocusedInput = box.Name
    index = AAInputIndex(box)
    mUpdatingHint = True
    If mHintActive(index) Then box.Text = vbNullString
    mHintActive(index) = False
    box.ForeColor = mInputColor(index)
    mUpdatingHint = False
End Sub

Private Sub AAExitInput(ByVal box As MSForms.TextBox)
    mFocusedInput = vbNullString
    AAShowHint box
End Sub

Private Sub txbGapHorizontal_Enter()
    AAEnterInput txbGapHorizontal
End Sub

Private Sub txbGapHorizontal_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    AAExitInput txbGapHorizontal
End Sub

Private Sub txbGapVertical_Enter()
    AAEnterInput txbGapVertical
End Sub

Private Sub txbGapVertical_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    AAExitInput txbGapVertical
End Sub

Private Sub txbAreaWidth_Enter()
    AAEnterInput txbAreaWidth
End Sub

Private Sub txbAreaWidth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    AAExitInput txbAreaWidth
End Sub

Private Sub txbAreaHeight_Enter()
    AAEnterInput txbAreaHeight
End Sub

Private Sub txbAreaHeight_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    AAExitInput txbAreaHeight
End Sub

Private Sub txbQuantity_Enter()
    AAEnterInput txbQuantity
End Sub

Private Sub txbQuantity_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    AAExitInput txbQuantity
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
    AAUpdateQuantityEditor
End Sub
