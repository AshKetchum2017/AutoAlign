Option Explicit

Private Const AA_PRESET_PAGE_WIDTH_MM As Double = 325#
Private Const AA_PRESET_PAGE_HEIGHT_MM As Double = 485#

Private mAreaWidthMM As Double
Private mAreaHeightMM As Double
Private mHorizontalGap As Double
Private mVerticalGap As Double
Private mHasHorizontalGap As Boolean
Private mHasVerticalGap As Boolean
Private mModelMode As Long
Private mDefaultHorizontalGapMM As Double, mDefaultVerticalGapMM As Double
Private mCustomArea As AACustomArea

Public Sub Configure(ByVal modelMode As Long, ByVal areaWidthText As String, _
                     ByVal areaHeightText As String, ByVal horizontalGapText As String, _
                     ByVal verticalGapText As String)
    Dim preferences As AADefaultPreferences
    If modelMode <> AA_MODEL_CUSTOM And modelMode <> AA_MODEL_KISS_A And _
        modelMode <> AA_MODEL_DIE_A Then Err.Raise vbObjectError + 5324, _
        "AASettings", "Model area tidak dikenal."
    mModelMode = modelMode
    Set mCustomArea = Nothing
    Set preferences = New AADefaultPreferences
    preferences.Load
    mAreaWidthMM = preferences.AreaWidthMM(modelMode)
    mAreaHeightMM = preferences.AreaHeightMM(modelMode)
    mDefaultHorizontalGapMM = preferences.GapHorizontalMM(modelMode)
    mDefaultVerticalGapMM = preferences.GapVerticalMM(modelMode)
    If modelMode = AA_MODEL_CUSTOM And Len(Trim$(areaWidthText)) > 0 Then
        mAreaWidthMM = ParseDecimal(areaWidthText, "Lebar area")
    End If
    If modelMode = AA_MODEL_CUSTOM And Len(Trim$(areaHeightText)) > 0 Then
        mAreaHeightMM = ParseDecimal(areaHeightText, "Tinggi area")
    End If
    If mAreaWidthMM <= 0# Or mAreaHeightMM <= 0# Then Err.Raise _
        vbObjectError + 5320, "AASettings", "Lebar dan tinggi area harus lebih besar dari nol."

    mHasHorizontalGap = False
    mHasVerticalGap = False
    mHorizontalGap = 0#
    mVerticalGap = 0#
    mHasHorizontalGap = (Len(Trim$(horizontalGapText)) > 0)
    mHasVerticalGap = (Len(Trim$(verticalGapText)) > 0)
    If mHasHorizontalGap Then
        mHorizontalGap = ParseDecimal(horizontalGapText, "Gap horizontal")
    End If
    If mHasVerticalGap Then
        mVerticalGap = ParseDecimal(verticalGapText, "Gap vertikal")
    End If
    If (mHasHorizontalGap And mHorizontalGap < 0#) Or _
       (mHasVerticalGap And mVerticalGap < 0#) Then Err.Raise vbObjectError + 5328, _
        "AASettings", "Gap tidak boleh negatif. Gunakan nilai minimal 0 mm."
End Sub

Public Property Get ModelMode() As Long
    ModelMode = mModelMode
End Property

Public Property Get PresetPageWidthMM() As Double
    PresetPageWidthMM = AA_PRESET_PAGE_WIDTH_MM
End Property

Public Property Get PresetPageHeightMM() As Double
    PresetPageHeightMM = AA_PRESET_PAGE_HEIGHT_MM
End Property

Public Property Get DefaultHorizontalGapMM() As Double
    DefaultHorizontalGapMM = mDefaultHorizontalGapMM
End Property

Public Property Get DefaultVerticalGapMM() As Double
    DefaultVerticalGapMM = mDefaultVerticalGapMM
End Property

Public Property Get DefaultZigZagVerticalFactor() As Double
    ' Hanya seed pencarian Medium; bukan gap atau pitch grid Minimum.
    If mModelMode = AA_MODEL_DIE_A Then
        DefaultZigZagVerticalFactor = 0.89
    Else
        DefaultZigZagVerticalFactor = 0.88
    End If
End Property

Public Property Get AreaWidthMM() As Double
    AreaWidthMM = mAreaWidthMM
    If Not mCustomArea Is Nothing Then AreaWidthMM = AreaWidthMM - mCustomArea.MarginLeft - mCustomArea.MarginRight
End Property

Public Property Get AreaHeightMM() As Double
    AreaHeightMM = mAreaHeightMM
    If Not mCustomArea Is Nothing Then AreaHeightMM = AreaHeightMM - mCustomArea.MarginTop - mCustomArea.MarginBottom
End Property

Public Sub ApplyCustomArea(ByVal profile As AACustomArea, ByVal doc As Document, ByVal pageIndex As Long)
    If mModelMode <> AA_MODEL_CUSTOM Or profile Is Nothing Then Exit Sub
    profile.CheckSource doc, pageIndex
    profile.Validate mAreaWidthMM, mAreaHeightMM
    Set mCustomArea = profile.Copy
End Sub

Public Property Get FullAreaWidthMM() As Double
    FullAreaWidthMM = mAreaWidthMM
End Property
Public Property Get FullAreaHeightMM() As Double
    FullAreaHeightMM = mAreaHeightMM
End Property
Public Property Get AreaOffsetX() As Double
    If Not mCustomArea Is Nothing Then AreaOffsetX = mCustomArea.MarginLeft
End Property
Public Property Get AreaOffsetY() As Double
    If Not mCustomArea Is Nothing Then AreaOffsetY = mCustomArea.MarginBottom
End Property
Public Property Get HasCustomRestrictions() As Boolean
    If Not mCustomArea Is Nothing Then HasCustomRestrictions = mCustomArea.HasRestrictions
End Property

Public Property Get IncludeDesignZones() As Boolean
    If Not mCustomArea Is Nothing Then _
        IncludeDesignZones = (mCustomArea.IncludeDesign And mCustomArea.ZoneCount > 0)
End Property

Public Function CustomSummary() As String
    If mCustomArea Is Nothing Then Exit Function
    CustomSummary = "Custom penuh " & CStr(FullAreaWidthMM) & " x " & CStr(FullAreaHeightMM) & _
        " mm; margin kiri/kanan/atas/bawah " & CStr(mCustomArea.MarginLeft) & "/" & _
        CStr(mCustomArea.MarginRight) & "/" & CStr(mCustomArea.MarginTop) & "/" & CStr(mCustomArea.MarginBottom) & _
        " mm; zona=" & CStr(mCustomArea.ZoneCount) & "; Design ikut zona=" & CStr(mCustomArea.IncludeDesign) & "."
End Function

Public Function CustomForbidden() As Collection
    If mCustomArea Is Nothing Then
        Set CustomForbidden = New Collection
    Else
        Set CustomForbidden = mCustomArea.Forbidden
    End If
End Function

' Origin efektif dalam koordinat halaman; langkah grid memakai ukuran penuh.
Public Function AreaLeft(ByVal pageCenterX As Double, ByVal areaIndex As Long) As Double
    AreaLeft = pageCenterX - FullAreaWidthMM / 2# + _
        CDbl(areaIndex) * (FullAreaWidthMM + 50#) + AreaOffsetX
End Function

Public Function AreaBottom(ByVal pageCenterY As Double) As Double
    AreaBottom = pageCenterY - FullAreaHeightMM / 2# + AreaOffsetY
End Function

Public Function AreaColumn(ByVal rootCenterX As Double, ByVal pageCenterX As Double) As Long
    Dim column As Double
    column = Int((rootCenterX - pageCenterX) / (FullAreaWidthMM + 50#) + 0.5)
    If column < 0# Then column = 0#
    If column > 2147483646# Then Err.Raise vbObjectError + 5417, "AASettings", "Posisi area terlalu jauh."
    AreaColumn = CLng(column)
End Function

Public Property Get HasHorizontalGap() As Boolean
    HasHorizontalGap = mHasHorizontalGap
End Property

Public Property Get HasVerticalGap() As Boolean
    HasVerticalGap = mHasVerticalGap
End Property

Public Property Get HorizontalGap() As Double
    HorizontalGap = mHorizontalGap
End Property

Public Property Get VerticalGap() As Double
    VerticalGap = mVerticalGap
End Property

Friend Function ParseDecimal(ByVal rawText As String, ByVal fieldName As String) As Double
    Dim numberText As String
    Dim character As String
    Dim digitCount As Long
    Dim separatorCount As Long
    Dim i As Long

    numberText = Trim$(rawText)

    For i = 1 To Len(numberText)
        character = Mid$(numberText, i, 1)
        Select Case character
            Case "0" To "9"
                digitCount = digitCount + 1
            Case ".", ","
                separatorCount = separatorCount + 1
            Case "+", "-"
                If i <> 1 Then Exit For
            Case Else
                Exit For
        End Select
    Next i
    If i <= Len(numberText) Or digitCount = 0 Or separatorCount > 1 Then GoTo InvalidNumber

    On Error GoTo InvalidNumber
    ParseDecimal = Val(Replace$(numberText, ",", "."))
    Exit Function

InvalidNumber:
    Err.Raise vbObjectError + 5323, "AASettings", _
        fieldName & " harus berisi angka."
End Function
