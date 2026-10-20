Option Explicit

' Snapshot milik sesi form, terikat dokumen/halaman. Tidak menyimpan Shape sumber.
Private mDocument As Document, mPageIndex As Long
Private mLeft As Double, mRight As Double, mTop As Double, mBottom As Double
Private mZones As Collection, mLabels As Collection
Private mIncludeDesign As Boolean

Private Sub Class_Initialize()
    Set mZones = New Collection
    Set mLabels = New Collection
    mIncludeDesign = True
End Sub

Public Property Get IncludeDesign() As Boolean
    IncludeDesign = mIncludeDesign
End Property

Public Sub SetIncludeDesign(ByVal value As Boolean)
    mIncludeDesign = value
End Sub

Public Sub Bind(ByVal doc As Document, ByVal pageIndex As Long)
    Set mDocument = doc
    mPageIndex = pageIndex
End Sub

Public Function IsSource(ByVal doc As Document, ByVal pageIndex As Long) As Boolean
    IsSource = (doc Is mDocument) And pageIndex = mPageIndex
End Function

Public Sub CheckSource(ByVal doc As Document, ByVal pageIndex As Long)
    If Not IsSource(doc, pageIndex) Then Err.Raise vbObjectError + 5450, "AACustomArea", _
        "Pengaturan area milik dokumen/halaman lain. Buka Pengaturan Area Custom untuk menetapkan pengaturan pada sumber ini."
End Sub

Public Property Get MarginLeft() As Double
    MarginLeft = mLeft
End Property
Public Property Get MarginRight() As Double
    MarginRight = mRight
End Property
Public Property Get MarginTop() As Double
    MarginTop = mTop
End Property
Public Property Get MarginBottom() As Double
    MarginBottom = mBottom
End Property
Public Property Get ZoneCount() As Long
    ZoneCount = mZones.Count
End Property
Public Property Get HasRestrictions() As Boolean
    HasRestrictions = (mLeft + mRight + mTop + mBottom > 0# Or mZones.Count > 0)
End Property

Public Sub SetMargins(ByVal leftMM As Double, ByVal rightMM As Double, _
                      ByVal topMM As Double, ByVal bottomMM As Double, _
                      ByVal widthMM As Double, ByVal heightMM As Double)
    If leftMM < 0# Or rightMM < 0# Or topMM < 0# Or bottomMM < 0# Then _
        Err.Raise vbObjectError + 5451, "AACustomArea", "Margin tidak boleh negatif."
    If widthMM <= leftMM + rightMM Or heightMM <= topMM + bottomMM Then _
        Err.Raise vbObjectError + 5452, "AACustomArea", "Margin harus menyisakan lebar dan tinggi ruang susun lebih dari nol."
    mLeft = leftMM: mRight = rightMM: mTop = topMM: mBottom = bottomMM
End Sub

Public Sub Validate(ByVal widthMM As Double, ByVal heightMM As Double)
    Dim zone As AAGeometry
    If widthMM <= mLeft + mRight Or heightMM <= mTop + mBottom Then _
        Err.Raise vbObjectError + 5452, "AACustomArea", "Margin harus menyisakan lebar dan tinggi ruang susun lebih dari nol."
    For Each zone In mZones
        ValidateZone zone, widthMM, heightMM
    Next zone
End Sub

Private Sub ValidateZone(ByVal zone As AAGeometry, ByVal widthMM As Double, ByVal heightMM As Double)
    If zone.Width <= 0# Or zone.Height <= 0# Then Err.Raise vbObjectError + 5453, _
        "AACustomArea", "Zona harus memiliki lebar dan tinggi lebih dari nol."
    If zone.LeftX < -0.0001 Or zone.BottomY < -0.0001 Or _
       zone.RightX > widthMM + 0.0001 Or zone.TopY > heightMM + 0.0001 Then _
        Err.Raise vbObjectError + 5454, "AACustomArea", _
            "Zona harus berada di dalam area Custom keseluruhan. Acuan X/Y adalah kiri bawah area pertama, sebelum margin."
End Sub

Public Sub AddRectangle(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                        ByVal widthMM As Double, ByVal heightMM As Double)
    Dim zone As New AAGeometry
    If w <= 0# Or h <= 0# Then Err.Raise vbObjectError + 5453, "AACustomArea", _
        "Lebar dan tinggi rectangle harus lebih dari nol."
    zone.AddRectangle x, y, x + w, y + h, mDocument
    AddSnapshot zone, "Rectangle", widthMM, heightMM
End Sub

Public Sub AddSnapshot(ByVal zone As AAGeometry, ByVal label As String, _
                       ByVal widthMM As Double, ByVal heightMM As Double)
    ValidateZone zone, widthMM, heightMM
    mZones.Add zone.Placed(zone.LeftX, zone.TopY)
    mLabels.Add label
End Sub

Public Sub RemoveZone(ByVal index As Long)
    mZones.Remove index
    mLabels.Remove index
End Sub

Public Sub ClearZones()
    Set mZones = New Collection
    Set mLabels = New Collection
End Sub

Public Function ZoneText(ByVal index As Long) As String
    Dim zone As AAGeometry
    Set zone = mZones(index)
    ZoneText = CStr(index) & ". " & CStr(mLabels(index)) & " | X/Y " & _
        Format$(zone.LeftX, "0.###") & "/" & Format$(zone.BottomY, "0.###") & _
        " | W/H " & Format$(zone.Width, "0.###") & "/" & Format$(zone.Height, "0.###") & " mm"
End Function

Public Function Copy() As AACustomArea
    Dim result As New AACustomArea
    result.LoadSnapshot mDocument, mPageIndex, mLeft, mRight, mTop, mBottom, mZones, mLabels
    result.SetIncludeDesign mIncludeDesign
    Set Copy = result
End Function

Friend Sub LoadSnapshot(ByVal doc As Document, ByVal pageIndex As Long, _
                        ByVal leftMM As Double, ByVal rightMM As Double, _
                        ByVal topMM As Double, ByVal bottomMM As Double, _
                        ByVal zones As Collection, ByVal labels As Collection)
    Dim i As Long, zone As AAGeometry
    Bind doc, pageIndex
    mLeft = leftMM: mRight = rightMM: mTop = topMM: mBottom = bottomMM
    ClearZones
    For i = 1 To zones.Count
        Set zone = zones(i)
        ' Snapshot geometri internal immutable; salin koleksi tanpa API kurva
        ' agar Copy aman sebelum dokumen beralih ke mm. Forbidden membuat salinan kerja.
        mZones.Add zone
        mLabels.Add CStr(labels(i))
    Next i
End Sub

Public Function Forbidden() As Collection
    Dim result As New Collection, zone As AAGeometry
    For Each zone In mZones
        result.Add zone.Placed(zone.LeftX - mLeft, zone.TopY - mBottom)
    Next zone
    Set Forbidden = result
End Function
