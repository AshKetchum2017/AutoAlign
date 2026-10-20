Option Explicit

Private Const REG_APP As String = "RinCorelMacros"
Private Const REG_SECTION As String = "AutoAlignDefaults"
Private Const REG_KEY As String = "ProfilesV1"

Private mAreaWidth(1 To 2) As Double, mAreaHeight(1 To 2) As Double
Private mGapH(1 To 2) As Double, mGapV(1 To 2) As Double

Private Sub Class_Initialize()
    ResetDefaults
End Sub

Private Sub ResetDefaults()
    Dim i As Long
    For i = 1 To 2
        mAreaWidth(i) = 320#: mAreaHeight(i) = 470#
        mGapH(i) = 0#: mGapV(i) = 0#
    Next i
    mGapH(2) = 0#: mGapV(2) = 0#
End Sub

Public Sub Load()
    Dim saved As String, parts As Variant, values(0 To 7) As Double
    Dim parser As AASettings, i As Long, profile As Long, offset As Long
    ResetDefaults
    On Error GoTo InvalidSnapshot
    saved = GetSetting(REG_APP, REG_SECTION, REG_KEY, vbNullString)
    If Len(saved) = 0 Then Exit Sub
    parts = Split(saved, "|")
    If UBound(parts) <> 7 Then GoTo InvalidSnapshot
    ' ParseDecimal tidak memanggil Configure/Load, sehingga tidak rekursif.
    Set parser = New AASettings
    For i = 0 To 7
        values(i) = ReadStoredNumber(CStr(parts(i)), parser)
        If i Mod 4 < 2 Then
            If values(i) <= 0# Then GoTo InvalidSnapshot
        Else
            If values(i) < 0# Then GoTo InvalidSnapshot
        End If
    Next i
    For profile = 1 To 2
        offset = (profile - 1) * 4
        mAreaWidth(profile) = values(offset): mAreaHeight(profile) = values(offset + 1)
        mGapH(profile) = values(offset + 2): mGapV(profile) = values(offset + 3)
    Next profile
    Exit Sub
InvalidSnapshot:
    ResetDefaults
    AADebugWrite "Defaults", "registry unavailable/invalid; using factory defaults"
End Sub

Friend Function ReadStoredNumber(ByVal text As String, ByVal parser As AASettings) As Double
    Dim parts As Variant, exponent As Double, mantissa As Double
    ' Str$ dapat memakai notasi E untuk Double kecil/besar. Ini hanya format
    ' penyimpanan; editor tetap memakai parser desimal biasa tanpa notasi E.
    parts = Split(UCase$(text), "E")
    If UBound(parts) = 0 Then
        ReadStoredNumber = parser.ParseDecimal(text, "Preferensi default")
        Exit Function
    End If
    If UBound(parts) <> 1 Then Err.Raise vbObjectError + 5383, "AADefaultPreferences", "Format snapshot tidak valid."
    mantissa = parser.ParseDecimal(CStr(parts(0)), "Preferensi default")
    exponent = parser.ParseDecimal(CStr(parts(1)), "Eksponen preferensi")
    If exponent <> Fix(exponent) Or exponent < -324# Or exponent > 308# Then _
        Err.Raise vbObjectError + 5383, "AADefaultPreferences", "Eksponen snapshot tidak valid."
    ReadStoredNumber = Val(text)
End Function

Public Sub Save()
    Dim profile As Long, saved As String
    ' Satu snapshot: Save menyimpan kedua draft dalam satu operasi registry.
    ' Urutan: KissA W/H/Hgap/Vgap, kemudian DieA W/H/Hgap/Vgap.
    For profile = 1 To 2
        If profile > 1 Then saved = saved & "|"
        saved = saved & Trim$(Str$(mAreaWidth(profile))) & "|" & _
            Trim$(Str$(mAreaHeight(profile))) & "|" & Trim$(Str$(mGapH(profile))) & "|" & _
            Trim$(Str$(mGapV(profile)))
    Next profile
    SaveSetting REG_APP, REG_SECTION, REG_KEY, saved
End Sub

Public Sub SetProfile(ByVal modelMode As Long, ByVal widthText As String, _
                      ByVal heightText As String, ByVal gapHText As String, ByVal gapVText As String)
    Dim parser As AASettings, profile As Long
    Dim w As Double, h As Double, gh As Double, gv As Double
    profile = ProfileIndex(modelMode)
    Set parser = New AASettings
    w = parser.ParseDecimal(widthText, "Lebar area")
    h = parser.ParseDecimal(heightText, "Tinggi area")
    gh = parser.ParseDecimal(gapHText, "Gap horizontal")
    gv = parser.ParseDecimal(gapVText, "Gap vertikal")
    If w <= 0# Or h <= 0# Then Err.Raise vbObjectError + 5380, _
        "AADefaultPreferences", "Lebar dan tinggi area harus lebih besar dari nol."
    If gh < 0# Or gv < 0# Then Err.Raise vbObjectError + 5381, _
        "AADefaultPreferences", "Gap tidak boleh negatif. Gunakan nilai minimal 0 mm."
    ' Seluruh input valid sebelum satu nilai draft pun diganti.
    mAreaWidth(profile) = w: mAreaHeight(profile) = h
    mGapH(profile) = gh: mGapV(profile) = gv
End Sub

Public Property Get AreaWidthMM(ByVal modelMode As Long) As Double
    If modelMode = AA_MODEL_CUSTOM Then AreaWidthMM = 320#: Exit Property
    AreaWidthMM = mAreaWidth(ProfileIndex(modelMode))
End Property

Public Property Get AreaHeightMM(ByVal modelMode As Long) As Double
    If modelMode = AA_MODEL_CUSTOM Then AreaHeightMM = 470#: Exit Property
    AreaHeightMM = mAreaHeight(ProfileIndex(modelMode))
End Property

Public Property Get GapHorizontalMM(ByVal modelMode As Long) As Double
    If modelMode = AA_MODEL_CUSTOM Then GapHorizontalMM = 0#: Exit Property
    GapHorizontalMM = mGapH(ProfileIndex(modelMode))
End Property

Public Property Get GapVerticalMM(ByVal modelMode As Long) As Double
    If modelMode = AA_MODEL_CUSTOM Then GapVerticalMM = 0#: Exit Property
    GapVerticalMM = mGapV(ProfileIndex(modelMode))
End Property

Private Function ProfileIndex(ByVal modelMode As Long) As Long
    Select Case modelMode
        Case AA_MODEL_KISS_A: ProfileIndex = 1
        Case AA_MODEL_DIE_A: ProfileIndex = 2
        Case Else: Err.Raise vbObjectError + 5382, "AADefaultPreferences", "Model preferensi tidak dikenal."
    End Select
End Function
