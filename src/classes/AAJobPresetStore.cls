Option Explicit

Private Const REG_APP As String = "RinCorelMacros"
Private Const REG_SECTION As String = "AutoAlignJobPresets"
Private mNames As Collection

Private Sub Class_Initialize()
    Set mNames = New Collection
End Sub

Public Sub Load()
    Dim entries As Variant, names As New Collection, i As Long
    entries = GetAllSettings(REG_APP, REG_SECTION)
    If Not IsEmpty(entries) Then
        For i = LBound(entries, 1) To UBound(entries, 1)
            names.Add CStr(entries(i, 0))
        Next i
    End If
    Set mNames = names
End Sub

Public Property Get Count() As Long
    Count = mNames.Count
End Property

Public Function NameAt(ByVal index As Long) As String
    NameAt = CStr(mNames(index))
End Function

Public Function FindName(ByVal text As String) As Long
    Dim i As Long
    For i = 1 To mNames.Count
        If StrComp(NameAt(i), Trim$(text), vbTextCompare) = 0 Then FindName = i: Exit Function
    Next i
End Function

Public Function ValidName(ByVal text As String) As String
    Dim name As String, i As Long, code As Long
    name = Trim$(text)
    If Len(name) = 0 Or Len(name) > 80 Then Err.Raise vbObjectError + 5481, _
        "AAJobPresetStore", "Nama preset harus berisi 1 sampai 80 karakter."
    If StrComp(name, "(Tanpa preset - DefaultSettings)", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 5486, "AAJobPresetStore", "Gunakan nama pekerjaan, bukan nama pilihan Tanpa preset."
    For i = 1 To Len(name)
        code = AscW(Mid$(name, i, 1))
        If (code >= 0 And code < 32) Or code = 127 Or Mid$(name, i, 1) = "\" Or Mid$(name, i, 1) = "/" Then _
            Err.Raise vbObjectError + 5482, "AAJobPresetStore", "Nama preset tidak boleh memuat slash atau karakter kontrol."
    Next i
    ValidName = name
End Function

Public Function Read(ByVal index As Long) As AAJobPreset
    Dim result As New AAJobPreset
    result.LoadSnapshot GetSetting(REG_APP, REG_SECTION, NameAt(index), vbNullString)
    Set Read = result
End Function

Public Sub Save(ByVal name As String, ByVal preset As AAJobPreset)
    Dim key As String, index As Long, snapshot As String
    key = ValidName(name)
    snapshot = preset.Serialize
    index = FindName(key)
    If index > 0 Then key = NameAt(index)
    ' Satu preset disimpan dalam satu nilai; profil DefaultSettings tidak disentuh.
    SaveSetting REG_APP, REG_SECTION, key, snapshot
    Load
End Sub

Public Sub Remove(ByVal index As Long)
    DeleteSetting REG_APP, REG_SECTION, NameAt(index)
    Load
End Sub
