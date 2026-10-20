Option Explicit

' Nilai pekerjaan saja: tidak memuat ID objek, Document, atau geometri zona.
Private mModel As Long, mMode As Long
Private mWidth As Double, mHeight As Double, mDefaultH As Double, mDefaultV As Double
Private mManualH As Boolean, mManualV As Boolean, mGapH As Double, mGapV As Double
Private mLeft As Double, mRight As Double, mTop As Double, mBottom As Double

Public Sub Capture(ByVal settings As AASettings, ByVal layoutMode As Long, ByVal area As AACustomArea)
    mModel = settings.ModelMode: mMode = layoutMode
    mWidth = settings.FullAreaWidthMM: mHeight = settings.FullAreaHeightMM
    mDefaultH = settings.DefaultHorizontalGapMM: mDefaultV = settings.DefaultVerticalGapMM
    mManualH = settings.HasHorizontalGap: mManualV = settings.HasVerticalGap
    mGapH = settings.HorizontalGap: mGapV = settings.VerticalGap
    mLeft = 0#: mRight = 0#: mTop = 0#: mBottom = 0#
    If mModel = AA_MODEL_CUSTOM And Not area Is Nothing Then
        mLeft = area.MarginLeft: mRight = area.MarginRight
        mTop = area.MarginTop: mBottom = area.MarginBottom
    End If
    Validate
End Sub

Public Sub Validate()
    If mModel <> AA_MODEL_CUSTOM And mModel <> AA_MODEL_KISS_A And mModel <> AA_MODEL_DIE_A Then _
        Err.Raise vbObjectError + 5470, "AAJobPreset", "Model preset tidak dikenal."
    If mMode <> AA_MODE_MINIMUM And mMode <> AA_MODE_MEDIUM Then _
        Err.Raise vbObjectError + 5471, "AAJobPreset", "Mode preset harus Minimum atau Medium."
    If mWidth <= 0# Or mHeight <= 0# Then _
        Err.Raise vbObjectError + 5472, "AAJobPreset", "Ukuran area preset harus positif."
    If mDefaultH < 0# Or mDefaultV < 0# Or mGapH < 0# Or mGapV < 0# Then _
        Err.Raise vbObjectError + 5473, "AAJobPreset", "Gap preset tidak boleh negatif."
    If mLeft < 0# Or mRight < 0# Or mTop < 0# Or mBottom < 0# Then _
        Err.Raise vbObjectError + 5474, "AAJobPreset", "Margin preset tidak boleh negatif."
    If mLeft >= mWidth Or mRight >= mWidth - mLeft Or _
       mTop >= mHeight Or mBottom >= mHeight - mTop Then _
        Err.Raise vbObjectError + 5475, "AAJobPreset", "Margin harus menyisakan area susun."
    If mModel <> AA_MODEL_CUSTOM And (mLeft <> 0# Or mRight <> 0# Or mTop <> 0# Or mBottom <> 0#) Then _
        Err.Raise vbObjectError + 5476, "AAJobPreset", "Margin hanya berlaku pada Custom."
End Sub

Public Function Serialize() As String
    Validate
    Serialize = "1|" & CStr(mModel) & "|" & CStr(mMode) & "|" & _
        Trim$(Str$(mWidth)) & "|" & Trim$(Str$(mHeight)) & "|" & _
        Trim$(Str$(mDefaultH)) & "|" & Trim$(Str$(mDefaultV)) & "|" & _
        CStr(Abs(CLng(mManualH))) & "|" & Trim$(Str$(mGapH)) & "|" & _
        CStr(Abs(CLng(mManualV))) & "|" & Trim$(Str$(mGapV)) & "|" & _
        Trim$(Str$(mLeft)) & "|" & Trim$(Str$(mRight)) & "|" & _
        Trim$(Str$(mTop)) & "|" & Trim$(Str$(mBottom))
End Function

Public Sub LoadSnapshot(ByVal text As String)
    Dim parts As Variant, values(0 To 13) As Double, i As Long
    Dim parser As AASettings, reader As AADefaultPreferences
    parts = Split(text, "|")
    If UBound(parts) <> 14 Then Err.Raise vbObjectError + 5477, "AAJobPreset", "Data preset tidak lengkap."
    If CStr(parts(0)) <> "1" Then Err.Raise vbObjectError + 5478, "AAJobPreset", "Versi preset belum didukung."
    Set parser = New AASettings: Set reader = New AADefaultPreferences
    For i = 0 To 13
        values(i) = reader.ReadStoredNumber(CStr(parts(i + 1)), parser)
    Next i
    If (values(0) <> AA_MODEL_CUSTOM And values(0) <> AA_MODEL_KISS_A And values(0) <> AA_MODEL_DIE_A) Or _
       (values(1) <> AA_MODE_MINIMUM And values(1) <> AA_MODE_MEDIUM) Then _
        Err.Raise vbObjectError + 5479, "AAJobPreset", "Model/mode preset tidak valid."
    If (values(6) <> 0# And values(6) <> 1#) Or (values(8) <> 0# And values(8) <> 1#) Then _
        Err.Raise vbObjectError + 5480, "AAJobPreset", "Status gap preset tidak valid."
    mModel = CLng(values(0)): mMode = CLng(values(1))
    mWidth = values(2): mHeight = values(3): mDefaultH = values(4): mDefaultV = values(5)
    mManualH = CBool(values(6)): mGapH = values(7): mManualV = CBool(values(8)): mGapV = values(9)
    mLeft = values(10): mRight = values(11): mTop = values(12): mBottom = values(13)
    Validate
End Sub

Public Function Copy() As AAJobPreset
    Set Copy = New AAJobPreset
    Copy.LoadSnapshot Serialize
End Function

Public Property Get ModelMode() As Long
    ModelMode = mModel
End Property
Public Property Get LayoutMode() As Long
    LayoutMode = mMode
End Property
Public Property Get FullAreaWidthMM() As Double
    FullAreaWidthMM = mWidth
End Property
Public Property Get FullAreaHeightMM() As Double
    FullAreaHeightMM = mHeight
End Property
Public Property Get DefaultHorizontalGapMM() As Double
    DefaultHorizontalGapMM = mDefaultH
End Property
Public Property Get DefaultVerticalGapMM() As Double
    DefaultVerticalGapMM = mDefaultV
End Property
Public Property Get HorizontalText() As String
    If mManualH Then HorizontalText = DecimalText(mGapH)
End Property
Public Property Get VerticalText() As String
    If mManualV Then VerticalText = DecimalText(mGapV)
End Property
Public Property Get WidthText() As String
    WidthText = DecimalText(mWidth)
End Property
Public Property Get HeightText() As String
    HeightText = DecimalText(mHeight)
End Property

Public Sub ApplyMargins(ByVal draft As AACustomArea)
    Validate
    If mModel = AA_MODEL_CUSTOM Then
        draft.SetMargins mLeft, mRight, mTop, mBottom, mWidth, mHeight
        draft.Validate mWidth, mHeight
    End If
End Sub

' Form menerima desimal biasa; ekspansi E menghindari input tak valid setelah reload.
Private Function DecimalText(ByVal value As Double) As String
    Dim text As String, at As Long, exponent As Long, point As Long, digits As String
    text = Trim$(Str$(value))
    at = InStr(1, text, "E", vbTextCompare)
    If at = 0 Then
        DecimalText = text
        Exit Function
    End If
    exponent = CLng(Mid$(text, at + 1))
    text = Left$(text, at - 1)
    point = InStr(text, ".")
    If point = 0 Then point = Len(text) + 1
    digits = Replace$(text, ".", "")
    point = point - 1 + exponent
    If point <= 0 Then
        DecimalText = "0." & String$(-point, "0") & digits
    ElseIf point >= Len(digits) Then
        DecimalText = digits & String$(point - Len(digits), "0")
    Else
        DecimalText = Left$(digits, point) & "." & Mid$(digits, point + 1)
    End If
End Function

Public Property Get Details() As String
    Dim modelText As String, modeText As String, h As String, v As String
    modelText = "Custom"
    If mModel = AA_MODEL_KISS_A Then modelText = "Kiss Cut A3+"
    If mModel = AA_MODEL_DIE_A Then modelText = "Die Cut A3+"
    modeText = "Minimum"
    If mMode = AA_MODE_MEDIUM Then modeText = "Medium"
    h = CStr(mDefaultH) & " mm (default tersimpan)"
    v = CStr(mDefaultV) & " mm (default tersimpan)"
    If mManualH Then h = CStr(mGapH) & " mm (manual)"
    If mManualV Then v = CStr(mGapV) & " mm (manual)"
    Details = modelText & " / " & modeText & vbCrLf & _
        "Area: " & CStr(mWidth) & " x " & CStr(mHeight) & " mm." & vbCrLf & _
        "Gap H: " & h & "; V: " & v & "." & vbCrLf & _
        "Margin kiri/kanan/atas/bawah: " & CStr(mLeft) & "/" & CStr(mRight) & "/" & _
        CStr(mTop) & "/" & CStr(mBottom) & " mm." & vbCrLf & _
        "Nilai default tersimpan tidak mengubah DefaultSettings. Aturan gap default rectangle tajam tetap berlaku." & vbCrLf & _
        "Zona terlarang tidak disimpan. Zona dan checkbox Design pada pekerjaan aktif dipertahankan; " & _
        "penerapan Custom ditolak bila zona tidak muat dalam area baru." & vbCrLf & _
        "Objects, pasangan, Quantity dan rotasi per Design tidak diubah. Apply tidak menjalankan Process."
End Property
