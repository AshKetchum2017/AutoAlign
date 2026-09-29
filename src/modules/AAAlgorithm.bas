Option Explicit
Option Private Module

Public Const AA_MODE_DENSE As Long = 1
Public Const AA_MODE_STRAIGHT As Long = 2
Public Const AA_MODE_ZIGZAG As Long = 3
Public Const AA_MODE_AUTO As Long = 4

Private Const AA_EDGE_TOLERANCE_MM As Double = 0.0001
Private Const AA_ROW_TOLERANCE_MM As Double = 10#

Public Function AAPlanPositions(ByVal shapeWidth As Double, ByVal shapeHeight As Double, _
                                ByVal areaWidth As Double, ByVal areaHeight As Double, _
                                ByVal layoutMode As Long, ByVal straightGap As Double, _
                                ByVal zigZagVerticalFactor As Double) As Collection
    Dim result As New Collection
    Dim horizontalStep As Double
    Dim verticalStep As Double
    Dim rowShift As Double
    Dim xOffset As Double
    Dim yOffset As Double
    Dim rowIndex As Long

    If shapeWidth <= 0# Or shapeHeight <= 0# Then Err.Raise vbObjectError + 5202, _
        "AAPlanPositions", "Ukuran ellipse harus lebih besar dari nol."
    If areaWidth <= 0# Or areaHeight <= 0# Then Err.Raise vbObjectError + 5204, _
        "AAPlanPositions", "Ukuran marker harus lebih besar dari nol."

    Select Case layoutMode
        Case AA_MODE_DENSE
            horizontalStep = shapeWidth
            verticalStep = shapeHeight
        Case AA_MODE_STRAIGHT
            horizontalStep = shapeWidth + straightGap
            verticalStep = shapeHeight + straightGap
        Case AA_MODE_ZIGZAG
            horizontalStep = shapeWidth + straightGap
            verticalStep = shapeHeight * zigZagVerticalFactor
            rowShift = horizontalStep / 2#
        Case Else
            Err.Raise vbObjectError + 5203, "AAPlanPositions", _
                "Mode Auto belum tersedia. Nantinya Auto memilih Straight atau ZigZag."
    End Select

    If horizontalStep <= 0# Or verticalStep <= 0# Then Err.Raise vbObjectError + 5205, _
        "AAPlanPositions", "Langkah penyusunan harus lebih besar dari nol."

    Do While yOffset + shapeHeight <= areaHeight + AA_EDGE_TOLERANCE_MM
        xOffset = 0#
        If layoutMode = AA_MODE_ZIGZAG And rowIndex Mod 2 = 1 Then xOffset = rowShift
        Do While xOffset + shapeWidth <= areaWidth + AA_EDGE_TOLERANCE_MM
            result.Add Array(xOffset, yOffset)
            xOffset = xOffset + horizontalStep
        Loop
        rowIndex = rowIndex + 1
        yOffset = rowIndex * verticalStep
    Loop

    Set AAPlanPositions = result
End Function

Public Function AAArrangeEllipse(ByVal sourceShape As Shape, ByVal marker As Shape, _
                                 ByVal targetPage As Page, ByVal positions As Collection) As Shape
    Dim duplicateShape As Shape
    Dim pageShape As Shape
    Dim insideMarker As ShapeRange
    Dim position As Variant
    Dim positionIndex As Long

    sourceShape.LeftX = marker.LeftX
    sourceShape.TopY = marker.TopY

    For positionIndex = 2 To positions.Count
        position = positions(positionIndex)
        Set duplicateShape = sourceShape.Duplicate(0#, 0#)
        duplicateShape.LeftX = marker.LeftX + CDbl(position(0))
        duplicateShape.TopY = marker.TopY - CDbl(position(1))
    Next positionIndex

    Set insideMarker = New ShapeRange
    ' Marker menentukan anggota group; pengguna menyiapkan area ini kosong.
    For Each pageShape In targetPage.Shapes
        If Not (pageShape Is marker) Then
            If AAIsInsideMarker(pageShape, marker) Then insideMarker.Add pageShape
        End If
    Next pageShape

    If insideMarker.Count < positions.Count Then Err.Raise vbObjectError + 5201, _
        "AAArrangeEllipse", "Sebagian ellipse hasil tidak ditemukan di dalam marker."
    Set AAArrangeEllipse = insideMarker.Group
End Function

Public Sub AAOrderGroupMembers(ByVal groupShape As Shape)
    Dim member As Shape
    Dim members() As Shape
    Dim centerX() As Double
    Dim centerY() As Double
    Dim order() As Long
    Dim previous As Shape
    Dim parent As Shape
    Dim x As Double
    Dim y As Double
    Dim width As Double
    Dim height As Double
    Dim memberCount As Long
    Dim i As Long

    memberCount = groupShape.Shapes.Count
    If memberCount < 2 Then Exit Sub

    ReDim members(1 To memberCount)
    ReDim centerX(1 To memberCount)
    ReDim centerY(1 To memberCount)

    i = 0
    For Each member In groupShape.Shapes
        i = i + 1
        Set members(i) = member
        member.GetBoundingBox x, y, width, height
        centerX(i) = x + width / 2#
        centerY(i) = y + height / 2#
    Next member

    AABuildZOrder centerX, centerY, memberCount, order

    ' Sama seperti SPO: kiri-atas paling belakang, lalu maju mengikuti baris.
    Set previous = members(order(1))
    For i = 2 To memberCount
        If Not members(order(i)).OrderIsInFrontOf(previous) Then
            members(order(i)).OrderFrontOf previous
        End If
        Set parent = members(order(i)).ParentGroup
        If parent Is Nothing Then Err.Raise vbObjectError + 5206, _
            "AAOrderGroupMembers", "Anggota keluar dari group setelah pengurutan Z-Order."
        If Not (parent Is groupShape) Then Err.Raise vbObjectError + 5206, _
            "AAOrderGroupMembers", "Anggota masuk ke group lain setelah pengurutan Z-Order."
        If Not members(order(i)).OrderIsInFrontOf(previous) Then Err.Raise vbObjectError + 5207, _
            "AAOrderGroupMembers", "Urutan Z-Order anggota group tidak sesuai."
        Set previous = members(order(i))
    Next i
End Sub

Private Function AAIsInsideMarker(ByVal candidate As Shape, ByVal marker As Shape) As Boolean
    AAIsInsideMarker = _
        (candidate.LeftX >= marker.LeftX - AA_EDGE_TOLERANCE_MM) And _
        (candidate.RightX <= marker.RightX + AA_EDGE_TOLERANCE_MM) And _
        (candidate.BottomY >= marker.BottomY - AA_EDGE_TOLERANCE_MM) And _
        (candidate.TopY <= marker.TopY + AA_EDGE_TOLERANCE_MM)
End Function

Private Sub AABuildZOrder(ByRef centerX() As Double, ByRef centerY() As Double, _
                          ByVal memberCount As Long, ByRef order() As Long)
    Dim i As Long
    Dim j As Long
    Dim firstInRow As Long
    Dim lastInRow As Long
    Dim candidate As Long
    Dim rowY As Double

    ReDim order(1 To memberCount)
    For i = 1 To memberCount
        order(i) = i
    Next i

    ' Y tertinggi lebih dulu; toleransi tiap baris diikat ke Y awal baris.
    For i = 2 To memberCount
        candidate = order(i)
        j = i - 1
        Do While j >= 1
            If centerY(order(j)) >= centerY(candidate) Then Exit Do
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = candidate
    Next i

    firstInRow = 1
    Do While firstInRow <= memberCount
        rowY = centerY(order(firstInRow))
        lastInRow = firstInRow
        Do While lastInRow < memberCount
            If rowY - centerY(order(lastInRow + 1)) > AA_ROW_TOLERANCE_MM Then Exit Do
            lastInRow = lastInRow + 1
        Loop

        For i = firstInRow + 1 To lastInRow
            candidate = order(i)
            j = i - 1
            Do While j >= firstInRow
                If centerX(order(j)) <= centerX(candidate) Then Exit Do
                order(j + 1) = order(j)
                j = j - 1
            Loop
            order(j + 1) = candidate
        Next i
        firstInRow = lastInRow + 1
    Loop
End Sub
