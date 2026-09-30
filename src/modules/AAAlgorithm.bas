Option Explicit
Option Private Module

Public Const AA_MODE_DENSE As Long = 1
Public Const AA_MODE_STRAIGHT As Long = 2
Public Const AA_MODE_ZIGZAG As Long = 3
Public Const AA_MODE_AUTO As Long = 4
Public Const AA_ROLE_DESIGN As String = "Design"
Public Const AA_ROLE_CUT_LINE As String = "Cut Line"

Private Const AA_EDGE_TOLERANCE_MM As Double = 0.0001
Private Const AA_MATCH_TOLERANCE_MM As Double = 10#
Private Const AA_MAX_OBJECTS_PER_SOURCE As Long = 5000

Public Function AAPlanPositions(ByVal shapeWidth As Double, ByVal shapeHeight As Double, _
                                ByVal areaWidth As Double, ByVal areaHeight As Double, _
                                ByVal layoutMode As Long, ByVal horizontalStep As Double, _
                                ByVal verticalStep As Double) As Collection
    Dim result As New Collection
    Dim rowShift As Double
    Dim xOffset As Double
    Dim yOffset As Double
    Dim rowIndex As Long

    If shapeWidth <= 0# Or shapeHeight <= 0# Then Err.Raise vbObjectError + 5202, _
        "AAPlanPositions", "Ukuran object harus lebih besar dari nol."
    If areaWidth <= 0# Or areaHeight <= 0# Then Err.Raise vbObjectError + 5204, _
        "AAPlanPositions", "Ukuran marker harus lebih besar dari nol."

    Select Case layoutMode
        Case AA_MODE_DENSE
            horizontalStep = shapeWidth
            verticalStep = shapeHeight
        Case AA_MODE_STRAIGHT, AA_MODE_ZIGZAG
            If layoutMode = AA_MODE_ZIGZAG Then rowShift = horizontalStep / 2#
        Case Else
            Err.Raise vbObjectError + 5203, "AAPlanPositions", "Mode penyusunan tidak dikenal."
    End Select

    If horizontalStep <= 0# Or verticalStep <= 0# Then Err.Raise vbObjectError + 5205, _
        "AAPlanPositions", "Langkah penyusunan harus lebih besar dari nol."

    Do While yOffset + shapeHeight <= areaHeight + AA_EDGE_TOLERANCE_MM
        xOffset = 0#
        If layoutMode = AA_MODE_ZIGZAG And rowIndex Mod 2 = 1 Then xOffset = rowShift
        Do While xOffset + shapeWidth <= areaWidth + AA_EDGE_TOLERANCE_MM
            If result.Count >= AA_MAX_OBJECTS_PER_SOURCE Then Err.Raise _
                vbObjectError + 5210, "AAPlanPositions", _
                "Satu sumber melebihi batas " & CStr(AA_MAX_OBJECTS_PER_SOURCE) & " object."
            result.Add Array(xOffset, yOffset)
            xOffset = xOffset + horizontalStep
        Loop
        rowIndex = rowIndex + 1
        yOffset = rowIndex * verticalStep
    Loop

    Set AAPlanPositions = result
End Function

Public Function AAArrangeShape(ByVal sourceShape As Shape, ByVal marker As Shape, _
                               ByVal positions As Collection, ByVal sourceLayer As Layer, _
                               ByRef plannedMembers As Collection) As Shape
    Dim duplicateShape As Shape
    Dim resultGroup As Shape
    Dim arrangedShapes As ShapeRange
    Dim position As Variant
    Dim positionIndex As Long

    sourceShape.LeftX = marker.LeftX
    sourceShape.TopY = marker.TopY
    Set arrangedShapes = New ShapeRange
    Set plannedMembers = New Collection
    arrangedShapes.Add sourceShape
    plannedMembers.Add sourceShape

    For positionIndex = 2 To positions.Count
        position = positions(positionIndex)
        Set duplicateShape = sourceShape.Duplicate(0#, 0#)
        duplicateShape.LeftX = marker.LeftX + CDbl(position(0))
        duplicateShape.TopY = marker.TopY - CDbl(position(1))
        arrangedShapes.Add duplicateShape
        plannedMembers.Add duplicateShape
    Next positionIndex

    If arrangedShapes.Count <> positions.Count Then Err.Raise vbObjectError + 5201, _
        "AAArrangeShape", "Jumlah object hasil tidak sesuai dengan rencana."
    Set resultGroup = arrangedShapes.Group
    If resultGroup.Layer.Index <> sourceLayer.Index Then resultGroup.MoveToLayer sourceLayer
    If resultGroup.Layer.Index <> sourceLayer.Index Then Err.Raise vbObjectError + 5208, _
        "AAArrangeShape", "Group hasil tidak berada pada layer sumber."
    Set AAArrangeShape = resultGroup
End Function

Public Function AAHasSameGridPattern(ByVal firstPositions As Collection, _
                                     ByVal secondPositions As Collection) As Boolean
    Dim i As Long
    Dim firstPreviousY As Double
    Dim secondPreviousY As Double
    Dim firstNewRow As Boolean
    Dim secondNewRow As Boolean
    Dim firstPosition As Variant
    Dim secondPosition As Variant

    If firstPositions.Count <> secondPositions.Count Then Exit Function
    For i = 2 To firstPositions.Count
        firstPosition = firstPositions(i - 1)
        secondPosition = secondPositions(i - 1)
        firstPreviousY = CDbl(firstPosition(1))
        secondPreviousY = CDbl(secondPosition(1))
        firstPosition = firstPositions(i)
        secondPosition = secondPositions(i)
        firstNewRow = Abs(CDbl(firstPosition(1)) - firstPreviousY) > AA_EDGE_TOLERANCE_MM
        secondNewRow = Abs(CDbl(secondPosition(1)) - secondPreviousY) > AA_EDGE_TOLERANCE_MM
        If firstNewRow <> secondNewRow Then Exit Function
    Next i
    AAHasSameGridPattern = True
End Function

Public Function AAMatchesDesign(ByVal cutGroup As Shape, ByVal designGroup As Shape, _
                                ByVal cutModel As AAModel, ByVal designModel As AAModel) As Boolean
    If cutModel.PositionCount <> designModel.PositionCount Then Exit Function
    If Abs(cutGroup.SizeWidth - designGroup.SizeWidth) > AA_MATCH_TOLERANCE_MM Then Exit Function
    If Abs(cutGroup.SizeHeight - designGroup.SizeHeight) > AA_MATCH_TOLERANCE_MM Then Exit Function
    AAMatchesDesign = AAHasSameGridPattern(cutModel.Positions, designModel.Positions)
End Function

Public Sub AAPlaceCutToLeft(ByVal cutGroup As Shape, ByRef resultGroups() As Shape, _
                            ByRef designIndexes() As Long, ByRef cutIndexes() As Long, _
                            ByVal cutRoleIndex As Long, ByVal groupGap As Double)
    If cutRoleIndex = 0 Then
        cutGroup.RightX = resultGroups(designIndexes(0)).LeftX - groupGap
    Else
        cutGroup.RightX = resultGroups(cutIndexes(cutRoleIndex - 1)).LeftX - groupGap
    End If
End Sub

Public Sub AAOrderRoles(ByVal targetPage As Page, ByRef resultGroups() As Shape, _
                        ByRef designIndexes() As Long, ByVal designCount As Long, _
                        ByRef cutIndexes() As Long, ByVal cutCount As Long)
    Dim passIndex As Long
    Dim designIndex As Long
    Dim cutIndex As Long
    Dim pairedCount As Long
    Dim designGroup As Shape
    Dim cutGroup As Shape
    Dim designLayer As Layer
    Dim cutLayer As Layer
    Dim orderChanged As Boolean

    pairedCount = designCount
    If cutCount < pairedCount Then pairedCount = cutCount

    ' Pasangan ditentukan oleh urutan pendaftaran di masing-masing peran.
    For passIndex = 1 To pairedCount * targetPage.Layers.Count + 1
        orderChanged = False
        For cutIndex = 0 To pairedCount - 1
            Set cutGroup = resultGroups(cutIndexes(cutIndex))
            Set designGroup = resultGroups(designIndexes(cutIndex))
            Set cutLayer = cutGroup.Layer
            Set designLayer = designGroup.Layer
            If AAIsNumberedLayer(cutLayer.Name) And AAIsNumberedLayer(designLayer.Name) Then
                If Val(Mid$(cutLayer.Name, 7)) < Val(Mid$(designLayer.Name, 7)) Then
                    ' Layer bernomor memakai nomor nama, bukan Index urutan layer.
                    designGroup.MoveToLayer cutLayer
                    cutGroup.MoveToLayer designLayer
                    If designGroup.Layer.Index <> cutLayer.Index Or _
                        cutGroup.Layer.Index <> designLayer.Index Then Err.Raise _
                        vbObjectError + 5316, "AAOrderRoles", _
                        "Pertukaran layer group Design dan Cut Line tidak selesai."
                    orderChanged = True
                End If
            ElseIf cutLayer.Index > designLayer.Index Then
                ' Layer bernama khusus mempertahankan keanggotaan group.
                cutLayer.MoveAbove designLayer
                orderChanged = True
            End If
        Next cutIndex
        If Not orderChanged Then Exit For
    Next passIndex

    For cutIndex = 0 To pairedCount - 1
        Set cutGroup = resultGroups(cutIndexes(cutIndex))
        Set designGroup = resultGroups(designIndexes(cutIndex))
        Set cutLayer = cutGroup.Layer
        Set designLayer = designGroup.Layer
        If AAIsNumberedLayer(cutLayer.Name) And AAIsNumberedLayer(designLayer.Name) Then
            If Val(Mid$(cutLayer.Name, 7)) < Val(Mid$(designLayer.Name, 7)) Then Err.Raise _
                vbObjectError + 5315, "AAOrderRoles", _
                "Cut Line belum berada pada Layer bernomor di atas pasangan Design."
        ElseIf cutLayer.Index > designLayer.Index Then
            Err.Raise vbObjectError + 5315, "AAOrderRoles", _
                "Urutan layer tidak dapat menempatkan Cut Line di atas pasangan Design."
        End If
    Next cutIndex

    ' Pada layer yang sama, semua Cut Line berada di depan semua Design.
    For cutIndex = 0 To cutCount - 1
        Set cutGroup = resultGroups(cutIndexes(cutIndex))
        For designIndex = 0 To designCount - 1
            Set designGroup = resultGroups(designIndexes(designIndex))
            If cutGroup.Layer.Index = designGroup.Layer.Index Then
                If Not cutGroup.OrderIsInFrontOf(designGroup) Then cutGroup.OrderFrontOf designGroup
            End If
        Next designIndex
    Next cutIndex
End Sub

Private Function AAIsNumberedLayer(ByVal layerName As String) As Boolean
    Dim numberText As String

    If LCase$(Left$(layerName, 6)) <> "layer " Then Exit Function
    numberText = Mid$(layerName, 7)
    If Len(numberText) = 0 Then Exit Function
    AAIsNumberedLayer = Not (numberText Like "*[!0-9]*")
End Function

Public Sub AAOrderGroupMembers(ByVal groupShape As Shape, ByVal plannedMembers As Collection)
    Dim member As Shape
    Dim previous As Shape
    Dim parent As Shape
    Dim memberCount As Long
    Dim i As Long

    If plannedMembers Is Nothing Then Err.Raise vbObjectError + 5209, _
        "AAOrderGroupMembers", "Urutan object dari rencana tidak tersedia."
    memberCount = plannedMembers.Count
    If groupShape.Shapes.Count <> memberCount Then Err.Raise vbObjectError + 5209, _
        "AAOrderGroupMembers", "Jumlah anggota group tidak sesuai dengan rencana."

    ' AAArrangeShape menambahkan source dan duplikat menurut urutan AAPlanPositions.
    For i = 1 To memberCount
        Set member = plannedMembers(i)
        Set parent = member.ParentGroup
        If parent Is Nothing Then Err.Raise vbObjectError + 5206, _
            "AAOrderGroupMembers", "Anggota keluar dari group setelah pengurutan Z-Order."
        If Not (parent Is groupShape) Then Err.Raise vbObjectError + 5206, _
            "AAOrderGroupMembers", "Anggota masuk ke group lain setelah pengurutan Z-Order."
    Next i

    If memberCount < 2 Then Exit Sub
    Set previous = plannedMembers(1)
    For i = 2 To memberCount
        Set member = plannedMembers(i)
        If Not member.OrderIsInFrontOf(previous) Then member.OrderFrontOf previous
        Set parent = member.ParentGroup
        If parent Is Nothing Then Err.Raise vbObjectError + 5206, _
            "AAOrderGroupMembers", "Anggota keluar dari group setelah pengurutan Z-Order."
        If Not (parent Is groupShape) Then Err.Raise vbObjectError + 5206, _
            "AAOrderGroupMembers", "Anggota masuk ke group lain setelah pengurutan Z-Order."
        If Not member.OrderIsInFrontOf(previous) Then Err.Raise vbObjectError + 5207, _
            "AAOrderGroupMembers", "Urutan Z-Order anggota group tidak sesuai."
        Set previous = member
    Next i
End Sub
