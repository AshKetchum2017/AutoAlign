Option Explicit

Private Const AA_DEFAULT_AREA_WIDTH_MM As Double = 320#
Private Const AA_DEFAULT_AREA_HEIGHT_MM As Double = 470#
Private Const AA_PRESET_PAGE_WIDTH_MM As Double = 325#
Private Const AA_PRESET_PAGE_HEIGHT_MM As Double = 485#

Private mAreaWidthMM As Double
Private mAreaHeightMM As Double
Private mHorizontalGap As Double
Private mVerticalGap As Double
Private mHasHorizontalGap As Boolean
Private mHasVerticalGap As Boolean
Private mHorizontalPercent As Boolean
Private mVerticalPercent As Boolean
Private mModelMode As Long

Public Sub Configure(ByVal modelMode As Long, ByVal areaWidthText As String, _
                     ByVal areaHeightText As String, ByVal horizontalGapText As String, _
                     ByVal verticalGapText As String, ByVal horizontalPercent As Boolean, _
                     ByVal verticalPercent As Boolean)
    If modelMode <> AA_MODEL_CUSTOM And modelMode <> AA_MODEL_KISS_A And _
        modelMode <> AA_MODEL_DIE_A Then Err.Raise vbObjectError + 5324, _
        "AASettings", "Model area tidak dikenal."
    mModelMode = modelMode
    mAreaWidthMM = AA_DEFAULT_AREA_WIDTH_MM
    mAreaHeightMM = AA_DEFAULT_AREA_HEIGHT_MM
    If modelMode = AA_MODEL_CUSTOM And Len(Trim$(areaWidthText)) > 0 Then
        mAreaWidthMM = ParseDecimal(areaWidthText, "Lebar area", False)
    End If
    If modelMode = AA_MODEL_CUSTOM And Len(Trim$(areaHeightText)) > 0 Then
        mAreaHeightMM = ParseDecimal(areaHeightText, "Tinggi area", False)
    End If
    If mAreaWidthMM <= 0# Or mAreaHeightMM <= 0# Then Err.Raise _
        vbObjectError + 5320, "AASettings", "Lebar dan tinggi area harus lebih besar dari nol."

    mHorizontalPercent = horizontalPercent
    mVerticalPercent = verticalPercent
    mHasHorizontalGap = False
    mHasVerticalGap = False
    mHorizontalGap = 0#
    mVerticalGap = 0#
    mHasHorizontalGap = (Len(Trim$(horizontalGapText)) > 0)
    mHasVerticalGap = (Len(Trim$(verticalGapText)) > 0)
    If mHasHorizontalGap Then
        mHorizontalGap = ParseDecimal(horizontalGapText, "Gap horizontal", mHorizontalPercent)
        If mHorizontalPercent And mHorizontalGap <= 0# Then Err.Raise _
            vbObjectError + 5321, "AASettings", "Persen gap horizontal harus lebih besar dari nol."
    End If
    If mHasVerticalGap Then
        mVerticalGap = ParseDecimal(verticalGapText, "Gap vertikal", mVerticalPercent)
        If mVerticalPercent And mVerticalGap <= 0# Then Err.Raise _
            vbObjectError + 5322, "AASettings", "Persen gap vertikal harus lebih besar dari nol."
    End If
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
    If mModelMode = AA_MODEL_CUSTOM Then
        DefaultHorizontalGapMM = 0#
    Else
        DefaultHorizontalGapMM = 1#
    End If
End Property

Public Property Get DefaultVerticalGapMM() As Double
    DefaultVerticalGapMM = DefaultHorizontalGapMM
End Property

Public Property Get AreaWidthMM() As Double
    AreaWidthMM = mAreaWidthMM
End Property

Public Property Get AreaHeightMM() As Double
    AreaHeightMM = mAreaHeightMM
End Property

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

Public Property Get HorizontalPercent() As Boolean
    HorizontalPercent = mHorizontalPercent
End Property

Public Property Get VerticalPercent() As Boolean
    VerticalPercent = mVerticalPercent
End Property

Private Function ParseDecimal(ByVal rawText As String, ByVal fieldName As String, _
                              ByVal percentMode As Boolean) As Double
    Dim numberText As String
    Dim character As String
    Dim digitCount As Long
    Dim separatorCount As Long
    Dim i As Long

    numberText = Trim$(rawText)
    If percentMode And Right$(numberText, 1) = "%" Then
        numberText = Trim$(Left$(numberText, Len(numberText) - 1))
    End If

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
        fieldName & " harus berisi angka" & IIf(percentMode, " dengan % opsional.", ".")
End Function
