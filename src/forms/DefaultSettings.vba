Option Explicit

Private mPreferences As AADefaultPreferences
Private mModel As Long, mUpdating As Boolean, mSaved As Boolean

Public Property Get Saved() As Boolean
    Saved = mSaved
End Property

Private Sub UserForm_Initialize()
    Set mPreferences = New AADefaultPreferences
    mPreferences.Load
    mUpdating = True
    optKissASet.GroupName = "AADefaultModels"
    optDieASet.GroupName = "AADefaultModels"
    SelectModel AA_MODEL_KISS_A
End Sub

Public Sub SelectModel(ByVal modelMode As Long)
    If modelMode <> AA_MODEL_DIE_A Then modelMode = AA_MODEL_KISS_A
    mUpdating = True
    mModel = modelMode
    optKissASet.Value = (modelMode = AA_MODEL_KISS_A)
    optDieASet.Value = (modelMode = AA_MODEL_DIE_A)
    ShowProfile
    mUpdating = False
End Sub

Private Sub ShowProfile()
    txbDefaultAreaWidth.Text = CStr(mPreferences.AreaWidthMM(mModel))
    txbDefaultAreaHeight.Text = CStr(mPreferences.AreaHeightMM(mModel))
    txbDefaultGapHorizontal.Text = CStr(mPreferences.GapHorizontalMM(mModel))
    txbDefaultGapVertical.Text = CStr(mPreferences.GapVerticalMM(mModel))
End Sub

Private Sub CaptureProfile()
    mPreferences.SetProfile mModel, txbDefaultAreaWidth.Text, txbDefaultAreaHeight.Text, _
        txbDefaultGapHorizontal.Text, txbDefaultGapVertical.Text
End Sub

Private Sub ChangeModel(ByVal modelMode As Long)
    Dim message As String
    If mUpdating Or modelMode = mModel Then Exit Sub
    On Error GoTo InvalidInput
    CaptureProfile
    SelectModel modelMode
    Exit Sub
InvalidInput:
    message = Err.Description
    mUpdating = True
    optKissASet.Value = (mModel = AA_MODEL_KISS_A)
    optDieASet.Value = (mModel = AA_MODEL_DIE_A)
    mUpdating = False
    MsgBox message, vbExclamation, "Default Settings"
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo SaveFailed
    CaptureProfile
    mPreferences.Save
    mSaved = True
    Me.Hide
    Exit Sub
SaveFailed:
    MsgBox Err.Description, vbExclamation, "Default Settings"
End Sub

Private Sub optDieASet_Click()
    If optDieASet.Value Then ChangeModel AA_MODEL_DIE_A
End Sub

Private Sub optKissASet_Click()
    If optKissASet.Value Then ChangeModel AA_MODEL_KISS_A
End Sub

Private Sub txbDefaultAreaHeight_Change()
' TextBox yang mengatur tinggi area default. Validasi input dilakukan di kelas AADefaultPreferences.
End Sub

Private Sub txbDefaultAreaWidth_Change()
' TextBox yang mengatur lebar area default. Validasi input dilakukan di kelas AADefaultPreferences.
End Sub

Private Sub txbDefaultGapHorizontal_Change()
' TextBox yang mengatur gap horizontal default. Validasi input dilakukan di kelas AADefaultPreferences.
End Sub

Private Sub txbDefaultGapVertical_Change()
' TextBox yang mengatur gap vertical default. Validasi input dilakukan di kelas AADefaultPreferences.
End Sub
