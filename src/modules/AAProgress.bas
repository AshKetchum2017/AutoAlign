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
Private mView As Object
Private mPumping As Boolean, mRestoreHost As Boolean
Private mStage As String, mDetail As String, mGroup As String
Private mTrial As Long, mCandidate As Long
Private mDone As Long, mTotal As Long
Private mClock As Double, mElapsed As Double, mLastPaint As Double

Public Function AAProgressRunning() As Boolean
    AAProgressRunning = Not mView Is Nothing
End Function

Public Sub AAProgressBegin(ByVal view As Object)
    If AAProgressRunning Then Err.Raise vbObjectError + 5350, "AAProgress", _
        "Auto Align masih berjalan."
    Set mView = view
    mClock = Timer: mElapsed = 0#: mLastPaint = -AA_PROGRESS_INTERVAL
    mTrial = 0: mCandidate = 0: mGroup = vbNullString
    mStage = "Memeriksa input": mDetail = vbNullString
    mDone = 0: mTotal = 0
    ' Form dapat modeless melalui Macro Runner. Cegah edit dokumen selama
    ' DoEvents; owner modal yang sudah disabled tidak boleh di-enable kembali.
    mHostWindow = Application.AppWindow.Handle
    If mHostWindow = 0 Then Err.Raise vbObjectError + 5351, "AAProgress", _
        "Jendela CorelDRAW tidak ditemukan."
    mRestoreHost = (AAIsWindowEnabled(mHostWindow) <> 0)
    If mRestoreHost Then
        AAEnableWindow mHostWindow, 0
        If AAIsWindowEnabled(mHostWindow) <> 0 Then Err.Raise _
            vbObjectError + 5352, "AAProgress", "Jendela dokumen tidak dapat dikunci."
    End If
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
    If mView Is Nothing Or mPumping Then Exit Sub
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
    Exit Sub
PaintFailed:
    Debug.Print "[AA][Progress] UI error " & CStr(Err.Number) & ": " & Err.Description
    Resume PumpEvents
End Sub

Public Sub AAProgressEnd(ByVal succeeded As Boolean)
    Dim view As Object
    If mView Is Nothing Then Exit Sub
    Set view = mView
    ' Tidak yield lagi: engine sudah selesai memulihkan dokumen/CommandGroup.
    On Error Resume Next
    AAProgressReadClock
    If mRestoreHost Then AAEnableWindow mHostWindow, 1
    mRestoreHost = False: mHostWindow = 0
    Set mView = Nothing
    If succeeded Then
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
