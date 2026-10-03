Option Explicit

Private mStore As AAObjectStore
Private mLastProcessCompleted As Boolean

Private Sub Class_Initialize()
    Set mStore = New AAObjectStore
End Sub

Public Property Get ObjectCount() As Long
    ObjectCount = mStore.Count
End Property

Public Function ObjectListText(ByVal index As Long) As String
    ObjectListText = mStore.RoleAt(index) & " | Obj. ID: /" & CStr(mStore.StaticIDAt(index))
    If mStore.RoleAt(index) = AA_ROLE_DESIGN Then
        If mStore.QuantityAt(index) >= 0 Then _
            ObjectListText = ObjectListText & " | Qty: " & mStore.QuantityTextAt(index)
    End If
End Function

Public Function ObjectRoleAt(ByVal index As Long) As String
    ObjectRoleAt = mStore.RoleAt(index)
End Function

Public Function ObjectQuantityTextAt(ByVal index As Long) As String
    ObjectQuantityTextAt = mStore.QuantityTextAt(index)
End Function

Public Sub SetObjectQuantityText(ByVal index As Long, ByVal valueText As String)
    mStore.SetQuantityTextAt index, valueText
End Sub

Public Sub ClearObjects()
    mStore.Clear
    mLastProcessCompleted = False
End Sub

Public Sub RemoveObjectAt(ByVal index As Long)
    mStore.RemoveAt index
End Sub

Public Sub RegisterSelected(ByVal roleLabel As String, ByRef feedback As String, _
                            ByRef feedbackStyle As VbMsgBoxStyle)
    Dim doc As Document
    Dim targetPage As Page
    Dim selectedShapes As ShapeRange
    Dim addedCount As Long

    feedback = vbNullString
    feedbackStyle = vbExclamation
    On Error GoTo RegisterFailed
    If Application.Documents.Count = 0 Then Err.Raise vbObjectError + 5310, _
        "AAPresenter", "Tidak ada dokumen aktif."
    Set doc = ActiveDocument
    Set targetPage = doc.ActivePage
    Set selectedShapes = ActiveSelectionRange
    addedCount = mStore.RegisterSelection(roleLabel, selectedShapes, doc, targetPage)
    If addedCount = 0 Then
        feedback = "Tidak ada object baru yang ditambahkan."
        feedbackStyle = vbInformation
    End If
    Exit Sub

RegisterFailed:
    feedback = Err.Description
End Sub

Public Sub RunLayout(ByVal modelMode As Long, ByVal areaWidthText As String, _
                     ByVal areaHeightText As String, ByVal horizontalGapText As String, _
                     ByVal verticalGapText As String, ByVal horizontalPercent As Boolean, _
                     ByVal verticalPercent As Boolean, ByRef feedback As String, _
                     ByRef feedbackStyle As VbMsgBoxStyle, _
                     Optional ByVal layoutMode As Long = AA_MODE_MINIMUM)
    Dim engine As AAProcessEngine

    Set engine = New AAProcessEngine
    engine.Run mStore, modelMode, areaWidthText, areaHeightText, _
        horizontalGapText, verticalGapText, horizontalPercent, verticalPercent, _
        mLastProcessCompleted, feedback, feedbackStyle, layoutMode
End Sub
