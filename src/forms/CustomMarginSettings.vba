Option Explicit

Private mOwner As Object, mDraft As AACustomArea, mDocument As Document
Private mPageIndex As Long, mWidth As Double, mHeight As Double
Private mClosing As Boolean

Public Sub Configure(ByVal owner As Object, ByVal draft As AACustomArea, _
                     ByVal doc As Document, ByVal pageIndex As Long, _
                     ByVal widthMM As Double, ByVal heightMM As Double)
    Set mOwner = owner: Set mDraft = draft.copy: Set mDocument = doc
    mPageIndex = pageIndex: mWidth = widthMM: mHeight = heightMM
    Me.Caption = "Area Custom " & CStr(widthMM) & " x " & CStr(heightMM) & " mm"
    txbMarginLeft.text = CStr(mDraft.MarginLeft)
    txbMarginRight.text = CStr(mDraft.MarginRight)
    txbMarginTop.text = CStr(mDraft.MarginTop)
    txbMarginBottom.text = CStr(mDraft.MarginBottom)
    chkIncludeDesign.TripleState = False
    chkIncludeDesign.value = mDraft.IncludeDesign
    txbZoneX.text = "0": txbZoneY.text = "0"
    txbZoneWidth.text = "": txbZoneHeight.text = ""
    txbZoneX.ControlTipText = "mm dari kiri area pertama sebelum margin."
    txbZoneY.ControlTipText = "mm dari bawah area pertama sebelum margin."
    cmdAddFromSelection.ControlTipText = "Pilih objek tertutup di CorelDRAW, lalu ambil salinan geometrinya."
    RefreshZones
End Sub

Private Sub CheckSource()
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5455, _
        "CustomMarginSettings", "Dokumen sumber sudah ditutup. Batalkan editor dan buka ulang pada dokumen tujuan."
    mDraft.CheckSource ActiveDocument, ActiveDocument.ActivePage.index
End Sub

Private Function Number(ByVal value As String, ByVal field As String) As Double
    Dim parser As New AASettings
    Number = parser.ParseDecimal(value, field)
End Function

Private Sub CaptureMargins()
    mDraft.SetMargins Number(txbMarginLeft.text, "Margin kiri"), Number(txbMarginRight.text, "Margin kanan"), _
        Number(txbMarginTop.text, "Margin atas"), Number(txbMarginBottom.text, "Margin bawah"), mWidth, mHeight
End Sub

Private Sub RefreshZones()
    Dim i As Long
    lbxForbiddenAreas.Clear
    lbxForbiddenAreas.MultiSelect = fmMultiSelectSingle
    For i = 1 To mDraft.ZoneCount
        lbxForbiddenAreas.AddItem mDraft.ZoneText(i)
    Next i
    cmdRemoveZone.enabled = False
    cmdClearZones.enabled = (mDraft.ZoneCount > 0)
End Sub

Private Function RestoreUnits(ByVal unit As cdrUnit, ByVal horizontal As cdrUnit, ByVal vertical As cdrUnit) As String
    On Error Resume Next
    Err.Clear
    mDocument.unit = unit
    If Err.Number <> 0 Then RestoreUnits = " satuan dokumen"
    Err.Clear
    mDocument.Rulers.HUnits = horizontal
    If Err.Number <> 0 Then RestoreUnits = RestoreUnits & " ruler horizontal"
    Err.Clear
    mDocument.Rulers.VUnits = vertical
    If Err.Number <> 0 Then RestoreUnits = RestoreUnits & " ruler vertikal"
    On Error GoTo 0
End Function

Private Sub cmdAddFromSelection_Click()
    Dim selected As ShapeRange, item As Shape, zone As AAGeometry, localZone As AAGeometry
    Dim candidate As AACustomArea, part As Curve, path As SubPath, i As Long
    Dim unitSaved As Boolean, previousUnit As cdrUnit, previousH As cdrUnit, previousV As cdrUnit
    Dim message As String, restoreError As String, originX As Double, originY As Double
    On Error GoTo Failed
    CheckSource
    Set selected = ActiveSelectionRange
    If selected.Count = 0 Then Err.Raise vbObjectError + 5456, "CustomMarginSettings", _
        "Pilih objek zona pada halaman sumber, lalu tekan Ambil Seleksi."
    Set candidate = mDraft.copy
    previousUnit = mDocument.unit: previousH = mDocument.Rulers.HUnits: previousV = mDocument.Rulers.VUnits
    unitSaved = True: mDocument.unit = cdrMillimeter
    originX = mDocument.ActivePage.centerX - mWidth / 2#
    originY = mDocument.ActivePage.centerY - mHeight / 2#
    For Each item In selected
        Set zone = New AAGeometry
        zone.Capture item
        For i = 1 To zone.CurveCount
            Set part = zone.CurveAt(i)
            For Each path In part.SubPaths
                If Not path.Closed Then Err.Raise vbObjectError + 5457, "CustomMarginSettings", _
                    "Zona dari seleksi harus memakai kurva tertutup. Tidak ada zona dari seleksi ini yang ditambahkan."
            Next path
        Next i
        Set zone = zone.OuterBoundary
        Set localZone = zone.placed(zone.leftX - originX, zone.topY - originY)
        candidate.AddSnapshot localZone, "Obj. /" & CStr(item.staticID), mWidth, mHeight
    Next item
    restoreError = RestoreUnits(previousUnit, previousH, previousV): unitSaved = False
    If Len(restoreError) > 0 Then Err.Raise vbObjectError + 5460, "CustomMarginSettings", _
        "State belum pulih:" & restoreError & ". Zona belum ditambahkan."
    ' Seluruh seleksi harus valid sebelum draft diganti.
    Set mDraft = candidate
    RefreshZones
    Exit Sub
Failed:
    message = Err.description
    If unitSaved Then restoreError = RestoreUnits(previousUnit, previousH, previousV)
    If Len(restoreError) > 0 Then message = message & vbCrLf & "State belum pulih:" & restoreError
    MsgBox message, vbExclamation, "Ambil Seleksi"
End Sub

Private Sub cmdAddRectangle_Click()
    Dim previousUnit As cdrUnit, previousH As cdrUnit, previousV As cdrUnit, unitSaved As Boolean
    Dim message As String, restoreError As String
    Dim candidate As AACustomArea
    On Error GoTo Failed
    CheckSource
    Set candidate = mDraft.copy
    previousUnit = mDocument.unit: previousH = mDocument.Rulers.HUnits: previousV = mDocument.Rulers.VUnits
    unitSaved = True: mDocument.unit = cdrMillimeter
    candidate.AddRectangle Number(txbZoneX.text, "X zona"), Number(txbZoneY.text, "Y zona"), _
        Number(txbZoneWidth.text, "Lebar zona"), Number(txbZoneHeight.text, "Tinggi zona"), mWidth, mHeight
    restoreError = RestoreUnits(previousUnit, previousH, previousV): unitSaved = False
    If Len(restoreError) > 0 Then Err.Raise vbObjectError + 5460, "CustomMarginSettings", _
        "State belum pulih:" & restoreError & ". Zona belum ditambahkan."
    Set mDraft = candidate
    RefreshZones
    Exit Sub
Failed:
    message = Err.description
    If unitSaved Then restoreError = RestoreUnits(previousUnit, previousH, previousV)
    If Len(restoreError) > 0 Then message = message & vbCrLf & "State belum pulih:" & restoreError
    MsgBox message, vbExclamation, "Tambah Rectangle"
End Sub

Private Sub cmdApply_Click()
    On Error GoTo Failed
    CheckSource
    CaptureMargins
    mDraft.SetIncludeDesign CBool(chkIncludeDesign.value)
    mDraft.Validate mWidth, mHeight
    CloseEditor True
    Exit Sub
Failed:
    MsgBox Err.description, vbExclamation, "Area Custom"
End Sub

Private Sub cmdCancel_Click()
    CloseEditor False
End Sub

Private Sub cmdClearZones_Click()
    mDraft.ClearZones
    RefreshZones
End Sub

Private Sub cmdRemoveZone_Click()
    If lbxForbiddenAreas.ListIndex < 0 Then Exit Sub
    mDraft.RemoveZone lbxForbiddenAreas.ListIndex + 1
    RefreshZones
End Sub

Private Sub lbxForbiddenAreas_Click()
    cmdRemoveZone.enabled = (lbxForbiddenAreas.ListIndex >= 0)
End Sub

Private Sub CloseEditor(ByVal apply As Boolean)
    Dim message As String
    If mClosing Then Exit Sub
    On Error GoTo Failed
    mClosing = True
    Me.Hide
    mOwner.AACloseAreaEditor mDraft, apply
    Set mOwner = Nothing: Set mDraft = Nothing: Set mDocument = Nothing
    Unload Me
    Exit Sub
Failed:
    message = Err.description
    mClosing = False
    Me.Show vbModeless
    MsgBox message, vbExclamation, "Area Custom"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CloseEditor False
    End If
End Sub
