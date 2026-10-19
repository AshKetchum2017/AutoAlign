Option Explicit

Private Const AA_MAX_INSPECTION_ROWS As Long = 1000
Private mIssues As Collection
Private mDocument As Document
Private mPageIndex As Long
Private mComplete As Boolean, mIncomplete As Boolean
Private mTotalIssues As Long, mReadCount As Long, mPairCount As Long
Private mStatus As String, mContext As String, mCreated As String

Private Sub Class_Initialize()
    Set mIssues = New Collection
    mCreated = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Friend Sub Bind(ByVal doc As Document, ByVal pageIndex As Long, ByVal context As String)
    Set mDocument = doc
    mPageIndex = pageIndex: mContext = context
End Sub

Friend Sub AddIssue(ByVal category As String, ByVal areaNumber As Long, _
                    ByVal firstID As Long, ByVal secondID As Long, ByVal detail As String)
    Dim item As AAInspectionIssue
    mTotalIssues = mTotalIssues + 1
    If mIssues.Count < AA_MAX_INSPECTION_ROWS Then
        Set item = New AAInspectionIssue
        item.Category = category: item.AreaNumber = areaNumber
        item.FirstID = firstID: item.SecondID = secondID: item.Detail = detail
        mIssues.Add item
    End If
    AADebugWrite "Inspect", category & "; area=" & CStr(areaNumber) & "; ID=" & _
        CStr(firstID) & "/" & CStr(secondID) & "; " & detail
End Sub

Friend Sub MarkIncomplete(ByVal reason As String)
    mIncomplete = True: mComplete = False
    If Len(mStatus) = 0 Then mStatus = reason
End Sub

Friend Sub CountRead()
    mReadCount = mReadCount + 1
End Sub

Friend Sub CountPair()
    mPairCount = mPairCount + 1
End Sub

Friend Sub Finish()
    mComplete = Not mIncomplete
End Sub

Public Property Get Completed() As Boolean
    Completed = mComplete
End Property

Public Property Get Count() As Long
    Count = mIssues.Count
End Property

Public Property Get TotalIssues() As Long
    TotalIssues = mTotalIssues
End Property

Public Property Get Issue(ByVal index As Long) As AAInspectionIssue
    Set Issue = mIssues(index)
End Property

Public Property Get Summary() As String
    If Not mComplete Then
        Summary = "Pemeriksaan belum selesai. " & mStatus
    ElseIf mTotalIssues > 0 Then
        Summary = "Pemeriksaan selesai: " & CStr(mTotalIssues) & " temuan."
    Else
        Summary = "Pemeriksaan selesai. Tidak ditemukan masalah pada objek yang diperiksa."
    End If
    Summary = Summary & vbCrLf & CStr(mReadCount) & " objek dibaca; " & _
        CStr(mPairCount) & " pasangan dikenali."
    If mTotalIssues > mIssues.Count Then Summary = Summary & vbCrLf & _
        "Daftar menampilkan " & CStr(mIssues.Count) & " temuan pertama dari " & CStr(mTotalIssues) & "."
End Property

Public Property Get Context() As String
    Context = mCreated & vbCrLf & mContext & vbCrLf & _
        "Hanya objek terdaftar yang diperiksa. Hasil berlaku saat pemeriksaan; jalankan ulang setelah mengedit susunan." & _
        vbCrLf & "Gap memakai validator kontur AA, bukan pengukuran jarak eksak atau verifikasi independen."
End Property

Friend Function Resolve(ByVal id As Long) As Shape
    Dim page As Page
    If mDocument Is Nothing Then Exit Function
    Set page = mDocument.Pages(mPageIndex)
    Set Resolve = FindIn(page.Shapes, id)
End Function

Private Function FindIn(ByVal items As Shapes, ByVal id As Long) As Shape
    Dim candidate As Shape, found As Shape
    For Each candidate In items
        If CLng(candidate.StaticID) = id Then Set FindIn = candidate: Exit Function
        If candidate.Type = cdrGroupShape Then
            Set found = FindIn(candidate.Shapes, id)
            If Not found Is Nothing Then Set FindIn = found: Exit Function
        End If
    Next candidate
End Function

Public Sub SelectIssue(ByVal index As Long)
    Dim item As AAInspectionIssue, first As Shape, second As Shape, selected As ShapeRange
    If AAProgressRunning Then Err.Raise vbObjectError + 5410, "AAInspectionReport", _
        "Tunggu sampai proses Auto Align selesai."
    If Application.Documents.Count = 0 Or mDocument Is Nothing Then GoTo StaleReport
    ' Baca ulang dokumen/page dan ID; jangan memakai reference Shape dari hasil lama.
    If Not ActiveDocument Is mDocument Then GoTo StaleReport
    If mDocument.ActivePage.Index <> mPageIndex Then GoTo StaleReport
    Set item = mIssues(index)
    Set selected = New ShapeRange
    If item.FirstID > 0 Then
        Set first = Resolve(item.FirstID)
        If first Is Nothing Then GoTo StaleReport
        selected.Add first
    End If
    If item.SecondID > 0 And item.SecondID <> item.FirstID Then
        Set second = Resolve(item.SecondID)
        If second Is Nothing Then GoTo StaleReport
        selected.Add second
    End If
    If selected.Count > 0 Then selected.CreateSelection
    Exit Sub
StaleReport:
    Err.Raise vbObjectError + 5411, "AAInspectionReport", _
        "Aktifkan dokumen/halaman sumber. Jika objek atau struktur group berubah, jalankan pemeriksaan ulang."
End Sub
