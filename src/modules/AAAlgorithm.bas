Option Explicit
Option Private Module

Public Const AA_MODE_DENSE As Long = 1
Public Const AA_MODE_STRAIGHT As Long = 2
Public Const AA_MODE_ZIGZAG As Long = 3
Public Const AA_MODE_AUTO As Long = 4
Public Const AA_MODE_MINIMUM As Long = 4
Public Const AA_MODE_MEDIUM As Long = 5
Public Const AA_MODEL_CUSTOM As Long = 1
Public Const AA_MODEL_KISS_A As Long = 2
Public Const AA_MODEL_DIE_A As Long = 3
Public Const AA_ROLE_DESIGN As String = "Design"
Public Const AA_ROLE_CUT_LINE As String = "Cut Line"

Public Sub AAOrderPlacementPair(ByVal designShape As Shape, ByVal cutShape As Shape)
    Dim designLayer As Layer
    Dim cutLayer As Layer

    Set designLayer = designShape.Layer
    Set cutLayer = cutShape.Layer
    If AAIsNumberedLayer(designLayer.Name) And AAIsNumberedLayer(cutLayer.Name) Then
        If Val(Mid$(cutLayer.Name, 7)) < Val(Mid$(designLayer.Name, 7)) Then
            designShape.MoveToLayer cutLayer
            cutShape.MoveToLayer designLayer
            If designShape.Layer.Index <> cutLayer.Index Or _
                cutShape.Layer.Index <> designLayer.Index Then Err.Raise _
                vbObjectError + 5316, "AAOrderPlacementPair", _
                "Pertukaran layer Design dan Cut Line tidak selesai."
        End If
    ElseIf cutLayer.Index > designLayer.Index Then
        cutLayer.MoveAbove designLayer
    End If
    If cutShape.Layer.Index = designShape.Layer.Index Then
        If Not cutShape.OrderIsInFrontOf(designShape) Then _
            cutShape.OrderFrontOf designShape
    End If
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

    ' Koleksi anggota mengikuti urutan slot dari AAPlanner.
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
