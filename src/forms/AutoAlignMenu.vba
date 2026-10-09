Option Explicit
' Merge declarations into the top of the target UserForm code module.
Private pMRObserver As Object
Private pMRToken As String

Private mPresenter As AAPresenter
Private mUpdatingQuantity As Boolean
Private mQuantityIndex As Long
Private mHintReady As Boolean
Private mUpdatingHint As Boolean
Private mHintActive(1 To 5) As Boolean
Private mInputColor(1 To 5) As Long
Private mFocusedInput As String
Private mModeReady As Boolean
Private mChangingMode As Boolean
Private mMediumActive As Boolean
Private mMediumVisited As Boolean
Private mMinimumGapH As String, mMinimumGapV As String
Private mMediumGapH As String, mMediumGapV As String
Private mMinimumHPercent As Boolean, mMinimumVPercent As Boolean

Private Sub lbxObjects_Click()
    AAUpdateQuantityEditor
End Sub

Private Sub lbxObjects_Change()
    AAUpdateQuantityEditor
End Sub

Private Sub optCustom_Click()
    If optCustom.value Then
        optKissA.value = False
        optDieA.value = False
    End If
    AAUpdateModelControls
End Sub

Private Sub optDieA_Click()
    If optDieA.value Then
        optCustom.value = False
        optKissA.value = False
    End If
    AAUpdateModelControls
End Sub

Private Sub optKissA_Click()
    If optKissA.value Then
        optCustom.value = False
        optDieA.value = False
    End If
    AAUpdateModelControls
End Sub

Private Sub cmdOptimize_Click()
' One of the last scopes of AutoAlign is to optimize the existing group. This button triggers the optimization process.
End Sub

Private Sub optMaximum_Click()
'
End Sub

Private Sub optMedium_Click()
    AAUpdateModeControls
End Sub

Private Sub optMinimum_Click()
    AAUpdateModeControls
End Sub

Private Sub txbGapHorizontal_Change()
    If mUpdatingHint Then Exit Sub
    If Not chkHPercent.value Then Exit Sub
    If AAGapIsNegative(txbGapHorizontal.Text) Then
        MsgBox "Persen gap horizontal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkHPercent.value = False
        AASetGapPercentText txbGapHorizontal, False
        Exit Sub
    End If
    AASetGapPercentText txbGapHorizontal, True
End Sub

Private Sub txbGapVertical_Change()
    If mUpdatingHint Then Exit Sub
    If Not chkVPercent.value Then Exit Sub
    If AAGapIsNegative(txbGapVertical.Text) Then
        MsgBox "Persen gap vertikal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkVPercent.value = False
        AASetGapPercentText txbGapVertical, False
        Exit Sub
    End If
    AASetGapPercentText txbGapVertical, True
End Sub

Private Sub chkHPercent_Click()
    If mChangingMode Then Exit Sub
    If chkHPercent.value And AAGapIsNegative(AAInputText(txbGapHorizontal)) Then
        MsgBox "Persen gap horizontal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkHPercent.value = False
        Exit Sub
    End If
    AASetGapPercentText txbGapHorizontal, chkHPercent.value
End Sub

Private Sub chkVPercent_Click()
    If mChangingMode Then Exit Sub
    If chkVPercent.value And AAGapIsNegative(AAInputText(txbGapVertical)) Then
        MsgBox "Persen gap vertikal tidak boleh negatif.", vbExclamation, "Auto Align"
        chkVPercent.value = False
        Exit Sub
    End If
    AASetGapPercentText txbGapVertical, chkVPercent.value
End Sub

Private Sub txbQuantity_Change()
    Dim i As Long

    If mUpdatingHint Or mUpdatingQuantity Or mQuantityIndex < 0 Then Exit Sub
    On Error GoTo InvalidQuantity
    mUpdatingQuantity = True
    For i = 0 To lbxObjects.ListCount - 1
        If lbxObjects.Selected(i) Then
            If mPresenter.ObjectRoleAt(i) = AA_ROLE_DESIGN Then
                mPresenter.SetObjectQuantityText i, AAInputText(txbQuantity)
                lbxObjects.List(i) = mPresenter.ObjectListText(i)
            End If
        End If
    Next i
    mUpdatingQuantity = False
    Exit Sub
InvalidQuantity:
    MsgBox Err.Description, vbExclamation, "Auto Align"
    mUpdatingQuantity = False
    AAUpdateQuantityEditor
End Sub

Private Sub UserForm_Initialize()
    Set mPresenter = New AAPresenter
    mQuantityIndex = -1
    lbxObjects.MultiSelect = fmMultiSelectExtended
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
    optKissA.value = True
    optDieA.value = False
    optCustom.value = False
    optMinimum.value = True
    optMedium.Enabled = True
    optMaximum.Enabled = False
    AAUpdateModelControls
    txbGapHorizontal.Enabled = True
    txbGapVertical.Enabled = True
    chkHPercent.Enabled = True
    chkVPercent.Enabled = True
    AASetGapPercentText txbGapHorizontal, chkHPercent.value
    AASetGapPercentText txbGapVertical, chkVPercent.value
    mModeReady = True
    AAUpdateModeControls
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

    If Not optMinimum.value And Not optMedium.value Then
        MsgBox "Pilih mode Minimum atau Medium terlebih dahulu.", vbExclamation, "Auto Align"
        Exit Sub
    End If
    If optCustom.value Then
        modelMode = AA_MODEL_CUSTOM
    ElseIf optKissA.value Then
        modelMode = AA_MODEL_KISS_A
    ElseIf optDieA.value Then
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
    txbAreaWidth.Enabled = optCustom.value
    txbAreaHeight.Enabled = optCustom.value
    cmdProcess.Enabled = True
    AARefreshHints
End Sub

Private Sub AAUpdateModeControls()
    If Not mModeReady Or mChangingMode Then Exit Sub
    mChangingMode = True
    If optMedium.value <> mMediumActive Then
        If optMedium.value Then
            mMinimumGapH = AAInputText(txbGapHorizontal)
            mMinimumGapV = AAInputText(txbGapVertical)
            mMinimumHPercent = chkHPercent.value
            mMinimumVPercent = chkVPercent.value
            ' Input literal diteruskan pada kunjungan pertama; persen tidak diubah ke mm.
            If Not mMediumVisited Then
                If Not mMinimumHPercent Then mMediumGapH = mMinimumGapH
                If Not mMinimumVPercent Then mMediumGapV = mMinimumGapV
                mMediumVisited = True
            End If
            chkHPercent.value = False
            chkVPercent.value = False
            AASetModeGapText mMediumGapH, mMediumGapV
        Else
            mMediumGapH = AAInputText(txbGapHorizontal)
            mMediumGapV = AAInputText(txbGapVertical)
            chkHPercent.value = mMinimumHPercent
            chkVPercent.value = mMinimumVPercent
            AASetModeGapText mMinimumGapH, mMinimumGapV
        End If
        mMediumActive = optMedium.value
    End If
    chkHPercent.Enabled = Not optMedium.value
    chkVPercent.Enabled = Not optMedium.value
    mChangingMode = False
    AARefreshHints
End Sub

Private Sub AASetModeGapText(ByVal horizontalText As String, ByVal verticalText As String)
    mUpdatingHint = True
    mHintActive(1) = False: mHintActive(2) = False
    txbGapHorizontal.ForeColor = mInputColor(1)
    txbGapVertical.ForeColor = mInputColor(2)
    txbGapHorizontal.Text = horizontalText
    txbGapVertical.Text = verticalText
    mUpdatingHint = False
End Sub

Private Sub AAUpdateQuantityEditor()
    Dim i As Long
    Dim commonText As String
    Dim mixedQuantity As Boolean

    If mUpdatingQuantity Or mPresenter Is Nothing Then Exit Sub
    mQuantityIndex = -1
    For i = 0 To lbxObjects.ListCount - 1
        If lbxObjects.Selected(i) Then
            If mPresenter.ObjectRoleAt(i) = AA_ROLE_DESIGN Then
                If mQuantityIndex < 0 Then
                    mQuantityIndex = i
                    commonText = mPresenter.ObjectQuantityTextAt(i)
                ElseIf commonText <> mPresenter.ObjectQuantityTextAt(i) Then
                    mixedQuantity = True
                End If
            End If
        End If
    Next i
    mUpdatingQuantity = True
    If mHintReady Then
        mHintActive(5) = False
        txbQuantity.ForeColor = mInputColor(5)
    End If
    txbQuantity.Enabled = (mQuantityIndex >= 0)
    If mQuantityIndex >= 0 And Not mixedQuantity Then
        txbQuantity.Text = commonText
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
    Dim layoutMode As Long
    Dim errorNumber As Long, errorSource As String, errorDescription As String

    If mPresenter Is Nothing Then Set mPresenter = New AAPresenter
    layoutMode = AA_MODE_MINIMUM
    If optMedium.value Then layoutMode = AA_MODE_MEDIUM
    MRNotifyProcessTiming True
    On Error GoTo ProcessFailed
    mPresenter.RunLayout modelMode, AAInputText(txbAreaWidth), AAInputText(txbAreaHeight), _
        AAInputText(txbGapHorizontal), AAInputText(txbGapVertical), chkHPercent.value, _
        chkVPercent.value, feedback, feedbackStyle, layoutMode
    AARefreshObjects
    On Error GoTo 0
    MRNotifyProcessTiming False
    If Len(feedback) > 0 Then MsgBox feedback, feedbackStyle, "Auto Align"
    Exit Sub

ProcessFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    MRNotifyProcessTiming False
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub MRNotifyProcessTiming(ByVal started As Boolean)
    If pMRObserver Is Nothing Then Exit Sub
    On Error GoTo NotifyFailed
    If started Then
        CallByName pMRObserver, "AutoAlignProcessStarted", VbMethod, pMRToken
    Else
        CallByName pMRObserver, "AutoAlignProcessFinished", VbMethod, pMRToken
    End If
    Exit Sub
NotifyFailed:
    MsgBox "Gagal mencatat durasi Auto Align di Macro Runner (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
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
    Dim gapText As String, zigzagText As String
    zigzagText = "88%"
    If optCustom.value Then
        gapText = "0"
    ElseIf optDieA.value Then
        gapText = "1,5": zigzagText = "89%"
    Else
        gapText = "1"
    End If
    Select Case AAInputIndex(box)
        Case 1
            AAHintText = gapText
        Case 2
            If optMedium.value Then
                AAHintText = gapText
            Else
                AAHintText = gapText & " / " & zigzagText
            End If
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
    txbGapHorizontal.ControlTipText = "Default: KissA 1 mm, DieA 1,5 mm, Custom 0 mm. KissA/DieA: rectangle bersudut runcing memakai 0. Input manual diutamakan."
    txbGapVertical.ControlTipText = "Default Straight / ZigZag: KissA 1 mm / 88%, DieA 1,5 mm / 89%, Custom 0 mm / 88%. KissA/DieA: rectangle bersudut runcing memakai gap 0 pada kedua pola."
    If optMedium.value Then txbGapVertical.ControlTipText = _
        "Medium: gap kontur dalam mm. Sisi miring memakai nilai terbesar H/V. Persen dan gap negatif belum tersedia."
    txbQuantity.ControlTipText = "Kosong (-): isi sisa area. 0: lewati Design. /n: bagi rata dalam satu container kelompok n. Berlaku untuk semua Design terpilih (Ctrl/Shift)."
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
    Dim i As Long
    On Error GoTo InvalidQuantity
    For i = 0 To lbxObjects.ListCount - 1
        If lbxObjects.Selected(i) Then
            If mPresenter.ObjectRoleAt(i) = AA_ROLE_DESIGN Then _
                mPresenter.ValidateObjectQuantity i
        End If
    Next i
    AAExitInput txbQuantity
    Exit Sub
InvalidQuantity:
    Cancel = True
    MsgBox Err.Description, vbExclamation, "Auto Align"
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

    mUpdatingQuantity = True
    lbxObjects.Clear
    For i = 0 To mPresenter.ObjectCount - 1
        lbxObjects.AddItem mPresenter.ObjectListText(i)
    Next i
    mUpdatingQuantity = False
    AAUpdateQuantityEditor
End Sub

' Called only by MRTargetBridge; normal menu entry points remain unchanged.
Public Sub MRBindRunner(ByVal observer As Object, ByVal token As String)
    Set pMRObserver = observer
    pMRToken = token
End Sub

Public Sub MRDetachRunner()
    Set pMRObserver = Nothing
    pMRToken = vbNullString
End Sub

Private Sub UserForm_Terminate()
    Dim observer As Object, token As String
    On Error GoTo NotifyFailed
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
    Exit Sub
NotifyFailed:
    MsgBox "Gagal memberitahu Macro Runner bahwa form sudah ditutup (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub
