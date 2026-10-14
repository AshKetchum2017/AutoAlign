Option Explicit
Option Private Module

' Satu sesi Process per GMS. Tidak bergantung pada Debug dan tidak menyimpan
' referensi dokumen/shape. DoEvents hanya berjalan di checkpoint caller.
#If VBA7 Then
Private Declare PtrSafe Function AAEnableWindow Lib "user32" Alias "EnableWindow" _
    (ByVal hwnd As LongPtr, ByVal enabled As Long) As Long
Private Declare PtrSafe Function AAIsWindowEnabled Lib "user32" Alias "IsWindowEnabled" _
    (ByVal hwnd As LongPtr) As Long
Private mHostWindow As LongPtr
#Else
Private Declare Function AAEnableWindow Lib "user32" Alias "EnableWindow" _
    (ByVal hwnd As Long, ByVal enabled As Long) As Long
Private Declare Function AAIsWindowEnabled Lib "user32" Alias "IsWindowEnabled" _
    (ByVal hwnd As Long) As Long
Private mHostWindow As Long
#End If

Private Const AA_PROGRESS_INTERVAL As Double = 0.25
Public Const AA_ERROR_CANCELLED As Long = vbObjectError + 5390
Private mView As Object
Private mPumping As Boolean
Private mLockedDocumentWindows As Collection
Private mStage As String, mDetail As String, mGroup As String
Private mTrial As Long, mCandidate As Long
Private mDone As Long, mTotal As Long
Private mClock As Double, mElapsed As Double, mLastPaint As Double
Private mCancelRequested As Boolean, mFinishing As Boolean, mWasCancelled As Boolean

Public Function AAProgressRequestCancel() As Boolean
    If mView Is Nothing Or mFinishing Then Exit Function
    mCancelRequested = True
    AADebugWrite "Progress", "Cancel requested"
    AAProgressRequestCancel = True
End Function

Public Sub AAProgressCheckCancelled()
    If mCancelRequested And Not mFinishing Then Err.Raise AA_ERROR_CANCELLED, _
        "AAProgress", "Pemrosesan dibatalkan oleh pengguna."
End Sub

Public Sub AAProgressBeginCleanup(Optional ByVal cancelled As Boolean = False)
    ' Cleanup tidak yield/raise lagi, agar pemulihan tidak dibatalkan ulang.
    mFinishing = True
    mWasCancelled = cancelled
End Sub

Public Function AAProgressRunning() As Boolean
    AAProgressRunning = Not mView Is Nothing
End Function

Public Sub AAProgressBegin(ByVal view As Object)
    Dim doc As Document, documentWindow As Window
#If VBA7 Then
    Dim documentHandle As LongPtr
#Else
    Dim documentHandle As Long
#End If
    If AAProgressRunning Then Err.Raise vbObjectError + 5350, "AAProgress", _
        "Auto Align masih berjalan."
    Set mView = view
    Set mLockedDocumentWindows = New Collection
    mCancelRequested = False: mFinishing = False: mWasCancelled = False
    mClock = Timer: mElapsed = 0#: mLastPaint = -AA_PROGRESS_INTERVAL
    mTrial = 0: mCandidate = 0: mGroup = vbNullString
    mStage = "Memeriksa input": mDetail = vbNullString
    mDone = 0: mTotal = 0
    ' Owner modeless harus tetap aktif agar host meneruskan event klik form.
    ' Kunci hanya window dokumen, bukan AppWindow yang memiliki UserForm.
    ' Owner modal yang sudah disabled tidak diubah.
    mHostWindow = Application.AppWindow.Handle
    If mHostWindow = 0 Then Err.Raise vbObjectError + 5351, "AAProgress", _
        "Jendela CorelDRAW tidak ditemukan."
    If AAIsWindowEnabled(mHostWindow) <> 0 Then
        For Each doc In Application.Documents
            For Each documentWindow In doc.Windows
                documentHandle = documentWindow.Handle
                If documentHandle = 0 Or documentHandle = mHostWindow Then Err.Raise _
                    vbObjectError + 5352, "AAProgress", "Handle jendela dokumen tidak valid."
                If AAIsWindowEnabled(documentHandle) <> 0 Then
                    ' Catat sebelum memanggil API agar partial lock dapat dipulihkan.
                    mLockedDocumentWindows.Add documentHandle
                    AAEnableWindow documentHandle, 0
                    If AAIsWindowEnabled(documentHandle) <> 0 Then Err.Raise _
                        vbObjectError + 5352, "AAProgress", "Jendela dokumen tidak dapat dikunci."
                End If
            Next documentWindow
        Next doc
    End If
    AADebugWrite "Progress", "document windows locked=" & CStr(mLockedDocumentWindows.Count) & _
        "; host enabled=" & CStr(AAIsWindowEnabled(mHostWindow) <> 0)
    AAProgressPulse True
End Sub

Public Sub AAProgressStage(ByVal stage As String, Optional ByVal detail As String = "", _
                           Optional ByVal done As Long = 0, Optional ByVal total As Long = 0)
    If mView Is Nothing Then Exit Sub
    mStage = stage: mDetail = detail: mDone = done: mTotal = total
    ' Perubahan tahap tidak mem-bypass throttle: trial singkat tetap murah.
    AAProgressPulse
End Sub

Public Sub AAProgressStep(ByVal done As Long, ByVal total As Long)
    If mView Is Nothing Then Exit Sub
    mDone = done: mTotal = total
    AAProgressPulse
End Sub

Public Sub AAProgressGroup(ByVal groupLabel As String)
    If mView Is Nothing Then Exit Sub
    mGroup = groupLabel: mTrial = 0: mCandidate = 0
    AAProgressStage "Mencari susunan terbaik"
End Sub

Public Sub AAProgressTrial()
    If mView Is Nothing Then Exit Sub
    mTrial = mTrial + 1: mCandidate = 0
    AAProgressStage "Mencari susunan terbaik"
End Sub

Public Sub AAProgressCandidate()
    If mView Is Nothing Then Exit Sub
    mCandidate = mCandidate + 1
    AAProgressStage "Mencari susunan terbaik"
End Sub

Public Sub AAProgressPulse(Optional ByVal force As Boolean = False)
    Dim detail As String
    If mView Is Nothing Or mPumping Or mFinishing Then Exit Sub
    AAProgressCheckCancelled
    AAProgressReadClock
    If Not force And mElapsed - mLastPaint < AA_PROGRESS_INTERVAL Then Exit Sub
    mLastPaint = mElapsed
    detail = mDetail
    If mStage = "Mencari susunan terbaik" Or mStage = "Merapatkan susunan" Or _
       mStage = "Menyelaraskan susunan" Then
        detail = mGroup
        If mTrial > 0 Then detail = AAProgressJoin(detail, "Percobaan " & CStr(mTrial))
        If mCandidate > 0 Then detail = AAProgressJoin(detail, "Pola " & CStr(mCandidate))
    End If
    mPumping = True
    ' Kesalahan presentasi tidak mengubah hasil/penanganan error geometri.
    On Error GoTo PaintFailed
    CallByName mView, "AAProgressRender", VbMethod, mStage, detail, mDone, mTotal, mElapsed
PumpEvents:
    On Error GoTo PumpFinished
    DoEvents
PumpFinished:
    mPumping = False
    On Error GoTo 0
    ' Di luar handler presentasi: sinyal Cancel harus mencapai engine.
    AAProgressCheckCancelled
    Exit Sub
PaintFailed:
    Debug.Print "[AA][Progress] UI error " & CStr(Err.Number) & ": " & Err.Description
    Resume PumpEvents
End Sub

Public Sub AAProgressEnd(ByVal succeeded As Boolean)
    Dim view As Object
    Dim i As Long
    If mView Is Nothing Then Exit Sub
    Set view = mView
    ' Tidak yield lagi: engine sudah selesai memulihkan dokumen/CommandGroup.
    On Error Resume Next
    AAProgressReadClock
    If Not mLockedDocumentWindows Is Nothing Then
        For i = 1 To mLockedDocumentWindows.Count
            AAEnableWindow mLockedDocumentWindows(i), 1
        Next i
    End If
    Set mLockedDocumentWindows = Nothing
    mHostWindow = 0
    Set mView = Nothing
    If mWasCancelled Then
        CallByName view, "AAProgressRender", VbMethod, "Dibatalkan", "Pemrosesan dibatalkan.", 0&, 1&, mElapsed
    ElseIf succeeded Then
        CallByName view, "AAProgressRender", VbMethod, "Selesai", "", 1&, 1&, mElapsed
    Else
        CallByName view, "AAProgressRender", VbMethod, "Proses dihentikan", "Periksa pesan Auto Align.", 0&, 1&, mElapsed
    End If
    Set view = Nothing
    mPumping = False
End Sub

Private Sub AAProgressReadClock()
    Dim now As Double, delta As Double
    now = Timer
    delta = now - mClock
    If delta < 0# Then delta = delta + 86400# ' Timer melewati tengah malam.
    mElapsed = mElapsed + delta: mClock = now
End Sub

Private Function AAProgressJoin(ByVal first As String, ByVal second As String) As String
    If Len(first) > 0 Then first = first & " | "
    AAProgressJoin = first & second
End Function
