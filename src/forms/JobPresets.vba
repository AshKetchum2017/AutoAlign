Option Explicit

Private mOwner As Object
Private mStore As AAJobPresetStore
Private mRefreshing As Boolean

Public Sub Configure(ByVal owner As Object)
    Set mOwner = owner
    Set mStore = New AAJobPresetStore
    txbDetails.MultiLine = True
    txbDetails.Locked = True
    txbDetails.ScrollBars = fmScrollBarsVertical
    txbPresetName.MaxLength = 80
    lbxPresets.MultiSelect = fmMultiSelectSingle
    RefreshList vbNullString
End Sub

Private Sub RefreshList(ByVal selectedName As String)
    Dim i As Long, index As Long
    mStore.Load
    mRefreshing = True
    lbxPresets.Clear
    lbxPresets.AddItem "(Tanpa preset - DefaultSettings)"
    For i = 1 To mStore.Count
        lbxPresets.AddItem mStore.NameAt(i)
    Next i
    index = mStore.FindName(selectedName)
    lbxPresets.ListIndex = index
    mRefreshing = False
    RefreshDetails
End Sub

Private Sub lbxPresets_Click()
    If Not mRefreshing Then RefreshDetails
End Sub

Private Sub RefreshDetails()
    Dim preset As AAJobPreset, index As Long
    index = lbxPresets.ListIndex
    cmdApply.Enabled = (index >= 0)
    cmdDelete.Enabled = (index > 0)
    If index <= 0 Then
        txbPresetName.Text = vbNullString
        txbDetails.Text = "Terapkan pilihan ini untuk memakai DefaultSettings kembali." & vbCrLf & _
            "Model/mode, input manual, margin, zona, serta daftar Objects tetap." & vbCrLf & _
            "Save mengambil pengaturan pada form AA, bukan preset yang dipilih dalam daftar." & vbCrLf & _
            "Preset menyimpan model, mode, area, gap dan margin. Zona tidak ikut disimpan."
        Exit Sub
    End If
    txbPresetName.Text = mStore.NameAt(index)
    On Error GoTo InvalidPreset
    Set preset = mStore.Read(index)
    txbDetails.Text = preset.Details & vbCrLf & vbCrLf & _
        "Save mengambil pengaturan form AA; nama yang sama mengganti preset setelah konfirmasi."
    Exit Sub
InvalidPreset:
    cmdApply.Enabled = False
    txbDetails.Text = "Preset tidak dapat diterapkan: " & Err.Description & vbCrLf & _
        "Data belum diubah. Hapus atau timpa dengan pengaturan AA yang valid."
End Sub

Private Sub cmdSave_Click()
    Dim name As String, preset As AAJobPreset, excludedZones As Long
    On Error GoTo Failed
    name = mStore.ValidName(txbPresetName.Text)
    Set preset = mOwner.AACaptureJobPreset(excludedZones)
    If mStore.FindName(name) > 0 Then
        If MsgBox("Ganti preset '" & name & "' dengan pengaturan AA saat ini?", _
            vbYesNo + vbQuestion + vbDefaultButton2, "Job Presets") <> vbYes Then Exit Sub
    End If
    mStore.Save name, preset
    RefreshList name
    If excludedZones > 0 Then
        MsgBox "Preset disimpan. " & CStr(excludedZones) & _
            " zona terlarang pada pekerjaan ini tidak ikut disimpan.", vbInformation, "Job Presets"
    End If
    Exit Sub
Failed:
    MsgBox "Preset belum dapat disimpan: " & Err.Description, vbExclamation, "Job Presets"
End Sub

Private Sub cmdApply_Click()
    Dim preset As AAJobPreset, name As String
    If lbxPresets.ListIndex < 0 Then Exit Sub
    On Error GoTo Failed
    If lbxPresets.ListIndex > 0 Then
        name = mStore.NameAt(lbxPresets.ListIndex)
        Set preset = mStore.Read(lbxPresets.ListIndex)
    End If
    mOwner.AAApplyJobPreset preset, name
    Unload Me
    Exit Sub
Failed:
    MsgBox "Preset tidak dapat diterapkan: " & Err.Description, vbExclamation, "Job Presets"
End Sub

Private Sub cmdDelete_Click()
    If lbxPresets.ListIndex <= 0 Then Exit Sub
    On Error GoTo Failed
    If MsgBox("Hapus preset '" & mStore.NameAt(lbxPresets.ListIndex) & "'?", _
        vbYesNo + vbQuestion + vbDefaultButton2, "Job Presets") <> vbYes Then Exit Sub
    mStore.Remove lbxPresets.ListIndex
    RefreshList vbNullString
    Exit Sub
Failed:
    MsgBox "Preset belum dapat dihapus: " & Err.Description, vbExclamation, "Job Presets"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Set mOwner = Nothing
    Set mStore = Nothing
End Sub
