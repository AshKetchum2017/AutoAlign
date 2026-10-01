Option Explicit

Private mAreaIndex As Long
Private mEntryIndex As Long
Private mRoleLabel As String
Private mLayer As Layer
Private mShapes As ShapeRange
Private mMembers As Collection
Private mResult As Shape

Private Sub Class_Initialize()
    Set mShapes = New ShapeRange
    Set mMembers = New Collection
End Sub

Public Sub Configure(ByVal areaIndex As Long, ByVal entryIndex As Long, _
                     ByVal roleLabel As String, ByVal targetLayer As Layer)
    mAreaIndex = areaIndex
    mEntryIndex = entryIndex
    mRoleLabel = roleLabel
    Set mLayer = targetLayer
End Sub

Public Sub Add(ByVal member As Shape)
    mShapes.Add member
    mMembers.Add member
End Sub

Public Sub Finish(ByVal sortMembers As Boolean)
    If mShapes.Count = 0 Then Err.Raise vbObjectError + 5240, _
        "AAGroupBucket", "Group hasil kosong."
    If mShapes.Count = 1 And sortMembers Then
        Set mResult = mShapes(1)
    Else
        Set mResult = mShapes.Group
        If sortMembers Then AAOrderGroupMembers mResult, mMembers
    End If
    If mResult.Layer.Index <> mLayer.Index Then mResult.MoveToLayer mLayer
    If mResult.Layer.Index <> mLayer.Index Then Err.Raise vbObjectError + 5241, _
        "AAGroupBucket", "Group hasil tidak berada pada layer tujuan."
End Sub

Public Property Get AreaIndex() As Long
    AreaIndex = mAreaIndex
End Property

Public Property Get EntryIndex() As Long
    EntryIndex = mEntryIndex
End Property

Public Property Get RoleLabel() As String
    RoleLabel = mRoleLabel
End Property

Public Property Get TargetLayer() As Layer
    Set TargetLayer = mLayer
End Property

Public Property Get Result() As Shape
    Set Result = mResult
End Property
