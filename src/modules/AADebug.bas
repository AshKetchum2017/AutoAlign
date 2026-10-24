Option Explicit
Option Private Module

' Pola Currency sama dengan MRStopwatch; tidak membutuhkan referensi GMS lain.
#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Public Const AA_PERF_CONTOUR As Long = 1
Public Const AA_PERF_GAP As Long = 2
Public Const AA_PERF_OVERLAP As Long = 3
Public Const AA_PERF_PLACEMENT As Long = 4
Public Const AA_PERF_WELD As Long = 5
Public Const AA_PERF_SLOT_PREP As Long = 6
Public Const AA_PERF_SHIFT_PREP As Long = 7
Public Const AA_PERF_SLOT_SORT As Long = 8
Private Const AA_PERF_KIND_COUNT As Long = 8
' Flag dibaca langsung oleh fungsi geometri yang sering dipanggil.
Public AADebugProfiling As Boolean
' Aktif hanya saat replay diagnostik sinkron; tidak memengaruhi keputusan gap.
Public AADebugGapWitnessActive As Boolean
' Counter ditambah langsung di query agar tidak ada call/printing per query.
Public AADebugQueryCacheHits As Double, AADebugQueryCacheMisses As Double
Public AADebugEdgeBatchGroups As Double, AADebugEdgeBatchFallbacks As Double
Public AADebugEdgeBatchQueries As Double, AADebugEdgeBatchRechecks As Double
Public AADebugEdgeBatchPrechecks As Double
Public AADebugEdgeBatchEdges As Double, AADebugEdgeBatchOffered As Double
Public AADebugEdgeBatchWide As Double, AADebugEdgeBatchSpread As Double, AADebugEdgeBatchDense As Double
Private mPerfSeconds(1 To AA_PERF_KIND_COUNT) As Double, mPerfCalls(1 To AA_PERF_KIND_COUNT) As Double
Private mClockFrequency As Currency, mClockStart As Currency
Private mClockTimerStart As Double, mClockUsesCounter As Boolean
Private mCandidateStarted As Double, mCandidateTiming As Boolean
Private mGroupsChecked As Double, mGroupsSkipped As Double
Private mGroupEdgesSkipped As Double, mGroupEdgesDetailed As Double
Private mOccupiedQueries As Double, mOccupiedAvailable As Double, mOccupiedOffered As Double
Private mOccupiedFallbacks As Double, mOccupiedCacheHits As Double
Private mOccupiedBuilds As Double, mOccupiedAdded As Double

' Immediate Window: AADebugOn (ringkas), AADebugOn True (detail), AADebugOff.
' Tidak disimpan: reset project mengembalikan debug ke kondisi nonaktif.
Private mFileNumber As Integer
Private mFilePath As String

Private mEnabled As Boolean
Private mDetailed As Boolean
Private mSlotsActive As Boolean
Private mCandidate As String
Private mBoundsRejected As Long, mGapRejected As Long
Private mMarkRejected As Long, mOverlapRejected As Long
Private mMarkSamples As Long
Private mMinimumGapSamples As Long

Private Type PROCESS_MEMORY_COUNTERS_EX
    cb As Long
    PageFaultCount As Long
#If VBA7 Then
    PeakWorkingSetSize As LongPtr
    WorkingSetSize As LongPtr
    QuotaPeakPagedPoolUsage As LongPtr
    QuotaPagedPoolUsage As LongPtr
    QuotaPeakNonPagedPoolUsage As LongPtr
    QuotaNonPagedPoolUsage As LongPtr
    PagefileUsage As LongPtr
    PeakPagefileUsage As LongPtr
    PrivateUsage As LongPtr
#Else
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
    PrivateUsage As Long
#End If
End Type

#If VBA7 Then
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetProcessMemoryInfo Lib "psapi.dll" ( _
    ByVal hProcess As LongPtr, _
    ByRef counters As PROCESS_MEMORY_COUNTERS_EX, _
    ByVal cb As Long) As Long
#Else
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" ( _
    ByVal hProcess As Long, _
    ByRef counters As PROCESS_MEMORY_COUNTERS_EX, _
    ByVal cb As Long) As Long
#End If

Private mMemoryTracking As Boolean, mMemoryTakeBaseline As Boolean
Private mMemoryHasBaseline As Boolean
Private mMemoryStartPrivate As Double
Private mMemoryMaxPrivate As Double, mMemoryMaxWorking As Double
Private mMemorySamples As Long

Public Sub AADebugOn(Optional ByVal detailed As Boolean = False)
    mEnabled = True
    mDetailed = detailed
    mSlotsActive = False
    AADebugProfiling = False
    AADebugGapWitnessActive = False
    mCandidateTiming = False
    AADebugMemoryReset
    AADebugWrite "Debug", "ON; detailed=" & CStr(detailed)
End Sub

Public Sub AADebugOff()
    AADebugFileWrite "[AA][Debug] OFF"
    AADebugCloseFile
    mEnabled = False
    mDetailed = False
    mSlotsActive = False
    AADebugProfiling = False
    AADebugGapWitnessActive = False
    mCandidateTiming = False
    AADebugMemoryReset
    Debug.Print "[AA][Debug] OFF"
End Sub

' Append agar trace run sebelumnya tetap tersedia; path kosong memakai TEMP.
' File Shared dapat dibaca saat Process berlangsung. AADebugOff menutup handle.
Public Sub AADebugFileOn(Optional ByVal filePath As String = vbNullString)
    Dim number As Long, description As String, folder As String, fileNumber As Integer
    On Error GoTo FileFailed
    If Len(Trim$(filePath)) = 0 Then
        folder = Environ$("TEMP")
        If Len(folder) = 0 Then folder = Environ$("TMP")
        If Len(folder) = 0 Then Err.Raise vbObjectError + 5340, "AADebug", "Folder TEMP tidak tersedia."
        If Right$(folder, 1) <> "\" Then folder = folder & "\"
        filePath = folder & "AutoAlign-Debug.txt"
    End If
    AADebugCloseFile
    fileNumber = FreeFile
    Open filePath For Append Access Write Shared As #fileNumber
    mFileNumber = fileNumber
    mFilePath = filePath
    If Not mEnabled Then AADebugOn
    If mFileNumber <> 0 Then AADebugWrite "Debug", "file ON; " & _
        Format$(Now, "yyyy-mm-dd hh:nn:ss") & "; path=" & mFilePath
    Exit Sub
FileFailed:
    number = Err.Number: description = Err.Description
    AADebugCloseFile
    If Not mEnabled Then AADebugOn
    Debug.Print "[AA][Debug] file unavailable; error=" & CStr(number) & "; " & description
    Err.Clear
End Sub

Public Sub AADebugFileOff()
    AADebugWrite "Debug", "file OFF"
    AADebugCloseFile
End Sub

Private Sub AADebugCloseFile()
    On Error Resume Next
    If mFileNumber <> 0 Then Close #mFileNumber
    mFileNumber = 0
    mFilePath = vbNullString
    Err.Clear
End Sub

Private Sub AADebugFileWrite(ByVal line As String)
    Dim number As Long, description As String
    If mFileNumber = 0 Then Exit Sub
    On Error GoTo WriteFailed
    Print #mFileNumber, line
    Exit Sub
WriteFailed:
    number = Err.Number: description = Err.Description
    AADebugCloseFile
    Debug.Print "[AA][Debug] file write failed; error=" & CStr(number) & "; " & description
    Err.Clear
End Sub

Public Function AADebugEnabled() As Boolean
    AADebugEnabled = mEnabled
End Function

Public Function AADebugDetailsEnabled() As Boolean
    AADebugDetailsEnabled = mEnabled And mDetailed
End Function

Public Sub AADebugWrite(ByVal section As String, ByVal message As String)
    Dim line As String
    If Not mEnabled Then Exit Sub
    On Error Resume Next
    line = "[AA][" & section & "] " & message
    Debug.Print line
    AADebugFileWrite line
    Err.Clear
End Sub

Public Sub AADebugProcessStart(ByVal layoutMode As Long, ByVal modelMode As Long)
    Dim layoutName As String, modelName As String
    If Not mEnabled Then Exit Sub
    mSlotsActive = False
    AADebugPerfReset
    layoutName = "Minimum"
    If layoutMode = AA_MODE_MEDIUM Then layoutName = "Medium"
    modelName = "Custom"
    If modelMode = AA_MODEL_KISS_A Then modelName = "KissA"
    If modelMode = AA_MODEL_DIE_A Then modelName = "DieA"
    AADebugWrite "Process", layoutName & " / " & modelName
    AADebugMemoryStart
End Sub

Public Sub AADebugOptimizeStart(ByVal modelMode As Long)
    Dim modelName As String
    If Not mEnabled Then Exit Sub
    mSlotsActive = False
    AADebugPerfReset
    modelName = "Custom"
    If modelMode = AA_MODEL_KISS_A Then modelName = "KissA"
    If modelMode = AA_MODEL_DIE_A Then modelName = "DieA"
    AADebugWrite "Optimize", "start / " & modelName
    AADebugMemoryStart
End Sub

Public Sub AADebugCandidateStart(ByVal label As String)
    mCandidateTiming = AADebugProfiling
    If mCandidateTiming Then mCandidateStarted = AADebugPerfBegin()
    If mEnabled Then mCandidate = label
    mSlotsActive = mEnabled And mDetailed
    If Not mSlotsActive Then Exit Sub
    mCandidate = label
    mBoundsRejected = 0: mGapRejected = 0
    mMarkRejected = 0: mOverlapRejected = 0
    mMarkSamples = 0
    mMinimumGapSamples = 0
End Sub

Public Sub AADebugCandidateEnd()
    If mCandidateTiming Then
        AADebugStageEnd "candidate build: " & mCandidate, mCandidateStarted
        mCandidateTiming = False
    End If
    If Not mSlotsActive Then Exit Sub
    mSlotsActive = False
    AADebugWrite "Slots", mCandidate & "; rejected bounds=" & CStr(mBoundsRejected) & _
        "; gap=" & CStr(mGapRejected) & "; mark=" & CStr(mMarkRejected) & _
        "; overlap=" & CStr(mOverlapRejected)
End Sub

Private Sub AADebugPerfReset()
    Dim i As Long
    On Error Resume Next
    AADebugProfiling = False
    mCandidateTiming = False
    For i = 1 To AA_PERF_KIND_COUNT
        mPerfCalls(i) = 0#: mPerfSeconds(i) = 0#
    Next i
    AADebugQueryCacheHits = 0#: AADebugQueryCacheMisses = 0#
    AADebugEdgeBatchGroups = 0#: AADebugEdgeBatchFallbacks = 0#
    AADebugEdgeBatchQueries = 0#: AADebugEdgeBatchRechecks = 0#
    AADebugEdgeBatchPrechecks = 0#
    AADebugEdgeBatchEdges = 0#: AADebugEdgeBatchOffered = 0#
    AADebugEdgeBatchWide = 0#: AADebugEdgeBatchSpread = 0#: AADebugEdgeBatchDense = 0#
    mGroupsChecked = 0#: mGroupsSkipped = 0#
    mGroupEdgesSkipped = 0#: mGroupEdgesDetailed = 0#
    mOccupiedQueries = 0#: mOccupiedAvailable = 0#: mOccupiedOffered = 0#
    mOccupiedFallbacks = 0#: mOccupiedCacheHits = 0#
    mOccupiedBuilds = 0#: mOccupiedAdded = 0#
    mClockTimerStart = Timer
    mClockUsesCounter = False
    If QueryPerformanceFrequency(mClockFrequency) <> 0 Then
        If mClockFrequency > 0 Then
            If QueryPerformanceCounter(mClockStart) <> 0 Then mClockUsesCounter = True
        End If
    End If
    AADebugProfiling = True
    If mClockUsesCounter Then
        AADebugWrite "Perf", "clock=QPC; times inclusive (do not sum nested stages/counters)"
    Else
        AADebugWrite "Perf", "clock=Timer fallback; times inclusive (do not sum nested stages/counters)"
    End If
    If mDetailed Then AADebugWrite "Perf", "detail mode includes extra sample diagnostics"
End Sub

Private Function AADebugPerfClock() As Double
    Dim finish As Currency
    On Error GoTo TimerOnly
    If mClockUsesCounter Then
        If QueryPerformanceCounter(finish) <> 0 Then
            AADebugPerfClock = CDbl(finish - mClockStart) / CDbl(mClockFrequency)
            Exit Function
        End If
    End If
TimerOnly:
    AADebugPerfClock = Timer - mClockTimerStart
    If AADebugPerfClock < 0# Then AADebugPerfClock = AADebugPerfClock + 86400#
End Function

Public Function AADebugPerfBegin() As Double
    If Not AADebugProfiling Then Exit Function
    AADebugPerfBegin = AADebugPerfClock()
End Function

Public Sub AADebugPerfEnd(ByVal kind As Long, ByVal started As Double)
    Dim elapsed As Double
    If Not AADebugProfiling Then Exit Sub
    On Error Resume Next
    elapsed = AADebugPerfClock() - started
    If elapsed < 0# Then elapsed = 0#
    mPerfCalls(kind) = mPerfCalls(kind) + 1#
    mPerfSeconds(kind) = mPerfSeconds(kind) + elapsed
End Sub

Public Sub AADebugPerfCount(ByVal kind As Long)
    If Not AADebugProfiling Then Exit Sub
    On Error Resume Next
    mPerfCalls(kind) = mPerfCalls(kind) + 1#
End Sub

Public Sub AADebugStageEnd(ByVal label As String, ByVal started As Double)
    Dim elapsed As Double
    If Not AADebugProfiling Then Exit Sub
    On Error Resume Next
    elapsed = AADebugPerfClock() - started
    If elapsed < 0# Then elapsed = 0#
    AADebugWrite "Perf", label & "=" & Format$(elapsed, "0.000") & " s"
End Sub

' Snapshot angka per scope: tidak mengubah counter global atau menahan object.
Public Function AADebugPerfSnapshot() As Variant
    Dim values() As Double, i As Long
    If Not AADebugProfiling Then Exit Function
    On Error GoTo SnapshotUnavailable
    ReDim values(0 To AA_PERF_KIND_COUNT * 2 - 1)
    For i = 1 To AA_PERF_KIND_COUNT
        values((i - 1) * 2) = mPerfCalls(i)
        values((i - 1) * 2 + 1) = mPerfSeconds(i)
    Next i
    AADebugPerfSnapshot = values
    Exit Function
SnapshotUnavailable:
    Err.Clear
End Function

Public Sub AADebugPerfDelta(ByVal label As String, ByVal snapshot As Variant)
    Dim i As Long, counterLabel As String, message As String
    If Not AADebugProfiling Then Exit Sub
    If Not IsArray(snapshot) Then Exit Sub
    On Error GoTo InvalidSnapshot
    If LBound(snapshot) <> 0 Or UBound(snapshot) <> AA_PERF_KIND_COUNT * 2 - 1 Then Exit Sub
    message = label
    For i = 1 To AA_PERF_KIND_COUNT
        Select Case i
            Case AA_PERF_CONTOUR: counterLabel = "contour"
            Case AA_PERF_GAP: counterLabel = "gap"
            Case AA_PERF_OVERLAP: counterLabel = "overlap"
            Case AA_PERF_PLACEMENT: counterLabel = "placement"
            Case AA_PERF_WELD: counterLabel = "weld"
            Case AA_PERF_SLOT_PREP: counterLabel = "slot prep"
            Case AA_PERF_SHIFT_PREP: counterLabel = "shift prep"
            Case AA_PERF_SLOT_SORT: counterLabel = "slot sort"
        End Select
        message = message & "; " & counterLabel & " calls=" & _
            Format$(mPerfCalls(i) - CDbl(snapshot((i - 1) * 2)), "0")
        If i <> AA_PERF_WELD Then message = message & "; " & counterLabel & " s=" & _
            Format$(mPerfSeconds(i) - CDbl(snapshot((i - 1) * 2 + 1)), "0.000")
    Next i
    AADebugWrite "Phase work", message
    Exit Sub
InvalidSnapshot:
    Err.Clear
End Sub

Public Sub AADebugContourCaptured(ByVal nodes As Long, ByVal edges As Long, ByVal flatnessMM As Double)
    If Not AADebugProfiling Then Exit Sub
    AADebugWrite "Contour", "nodes=" & CStr(nodes) & "; edges=" & CStr(edges) & _
        "; flatness=" & CStr(flatnessMM) & " mm"
End Sub

' Satu update per ClearOf yang memakai kelompok, bukan satu call per segmen.
Public Sub AADebugEdgeGroups(ByVal checked As Long, ByVal skipped As Long, _
                            ByVal skippedEdges As Long, ByVal detailEdges As Long)
    If Not AADebugProfiling Then Exit Sub
    mGroupsChecked = mGroupsChecked + checked
    mGroupsSkipped = mGroupsSkipped + skipped
    mGroupEdgesSkipped = mGroupEdgesSkipped + skippedEdges
    mGroupEdgesDetailed = mGroupEdgesDetailed + detailEdges
End Sub

Public Sub AADebugOccupiedBuild(ByVal building As Boolean, ByVal added As Long)
    If Not AADebugProfiling Then Exit Sub
    If building Then mOccupiedBuilds = mOccupiedBuilds + 1#
    mOccupiedAdded = mOccupiedAdded + added
End Sub

' offered menghitung kandidat yang disediakan, bukan pemeriksaan detail aktual.
Public Sub AADebugOccupiedQuery(ByVal available As Long, ByVal offered As Long, _
                               ByVal fallback As Boolean, ByVal cacheHit As Boolean)
    If Not AADebugProfiling Then Exit Sub
    mOccupiedQueries = mOccupiedQueries + 1#
    mOccupiedAvailable = mOccupiedAvailable + available
    mOccupiedOffered = mOccupiedOffered + offered
    If fallback Then mOccupiedFallbacks = mOccupiedFallbacks + 1#
    If cacheHit Then mOccupiedCacheHits = mOccupiedCacheHits + 1#
End Sub

Public Sub AADebugPerfFinish(ByVal status As String)
    Dim i As Long, label As String
    If Not AADebugProfiling Then Exit Sub
    On Error Resume Next
    AADebugStageEnd "Process total (" & status & ")", 0#
    For i = 1 To AA_PERF_KIND_COUNT
        Select Case i
            Case AA_PERF_CONTOUR: label = "contour capture"
            Case AA_PERF_GAP: label = "gap checks"
            Case AA_PERF_OVERLAP: label = "curve overlap"
            Case AA_PERF_PLACEMENT: label = "placements"
            Case AA_PERF_WELD: label = "overlap WeldWith"
            Case AA_PERF_SLOT_PREP: label = "initial slot preparation"
            Case AA_PERF_SHIFT_PREP: label = "refinement shift preparation"
            Case AA_PERF_SLOT_SORT: label = "initial slot sorting (inside preparation)"
        End Select
        If i = AA_PERF_WELD Then
            AADebugWrite "Perf", label & ": calls=" & Format$(mPerfCalls(i), "0")
        Else
            AADebugWrite "Perf", label & ": calls=" & Format$(mPerfCalls(i), "0") & _
                "; time=" & Format$(mPerfSeconds(i), "0.000") & " s"
        End If
    Next i
    AADebugWrite "Perf", "edge group filter: checked=" & Format$(mGroupsChecked, "0") & _
        "; skipped=" & Format$(mGroupsSkipped, "0") & _
        "; edges skipped=" & Format$(mGroupEdgesSkipped, "0") & _
        "; edges detailed=" & Format$(mGroupEdgesDetailed, "0")
    AADebugWrite "Perf", "edge query cache (indexed): hits=" & Format$(AADebugQueryCacheHits, "0") & _
        "; misses=" & Format$(AADebugQueryCacheMisses, "0")
    AADebugWrite "Perf", "edge batches: groups=" & Format$(AADebugEdgeBatchGroups, "0") & _
        "; fallback=" & Format$(AADebugEdgeBatchFallbacks, "0") & _
        "; queries=" & Format$(AADebugEdgeBatchQueries, "0") & _
        "; prechecks=" & Format$(AADebugEdgeBatchPrechecks, "0") & _
        "; rechecks=" & Format$(AADebugEdgeBatchRechecks, "0") & _
        "; source edges=" & Format$(AADebugEdgeBatchEdges, "0") & _
        "; offered=" & Format$(AADebugEdgeBatchOffered, "0") & _
        "; fallback cells/spread/dense=" & Format$(AADebugEdgeBatchWide, "0") & "/" & _
        Format$(AADebugEdgeBatchSpread, "0") & "/" & Format$(AADebugEdgeBatchDense, "0")
    AADebugWrite "Perf", "occupied index: queries=" & Format$(mOccupiedQueries, "0") & _
        "; builds=" & Format$(mOccupiedBuilds, "0") & "; added=" & Format$(mOccupiedAdded, "0") & _
        "; cache hits=" & Format$(mOccupiedCacheHits, "0") & "; fallbacks=" & Format$(mOccupiedFallbacks, "0") & _
        "; available=" & Format$(mOccupiedAvailable, "0") & "; offered=" & Format$(mOccupiedOffered, "0")
    AADebugMemoryFinish status
    AADebugProfiling = False
    mCandidateTiming = False
End Sub

Public Sub AADebugRejectSlot(ByVal reason As String)
    If Not mSlotsActive Then Exit Sub
    Select Case reason
        Case "bounds": mBoundsRejected = mBoundsRejected + 1
        Case "gap": mGapRejected = mGapRejected + 1
        Case "mark": mMarkRejected = mMarkRejected + 1
        Case "overlap": mOverlapRejected = mOverlapRejected + 1
    End Select
End Sub

Public Function AADebugSlotSnapshot() As Variant
    AADebugSlotSnapshot = Array(mBoundsRejected, mGapRejected, mMarkRejected, mOverlapRejected)
End Function

Public Sub AADebugMinimumSlotGap(ByVal original As AAPlacement, ByVal occupied As Collection)
    Dim placed As AAPlacement
    If Not mSlotsActive Then Exit Sub
    On Error GoTo TraceFailed
    For Each placed In occupied
        If Not original.HasClearance(placed) Then
            AADebugWrite "Minimum", "original mark slot gap check=False; raw L/T=" & _
                Format$(original.Footprint.LeftX, "0.000000") & "/" & Format$(original.Footprint.TopY, "0.000000")
            AADebugMinimumGapRejected original, placed, "Original mark slot"
            Exit Sub
        End If
    Next placed
    AADebugWrite "Minimum", "original mark slot gap check=True; raw L/T=" & _
        Format$(original.Footprint.LeftX, "0.000000") & "/" & Format$(original.Footprint.TopY, "0.000000")
TraceFailed:
    Err.Clear
End Sub

Public Sub AADebugMinimumGapRejected(ByVal candidate As AAPlacement, ByVal blocker As AAPlacement, _
                                     ByVal label As String)
    Dim clear As Boolean
    If Not mSlotsActive Then Exit Sub
    If mMinimumGapSamples >= 3 Then Exit Sub
    mMinimumGapSamples = mMinimumGapSamples + 1
    On Error GoTo TraceFailed
    AADebugWrite "Gap", label & "; first rejected edge pair (not global minimum distance)"
    AADebugValidationPlacement "candidate", candidate
    AADebugValidationPlacement "blocker", blocker
    AADebugGapWitnessActive = True
    clear = candidate.HasClearance(blocker)
    AADebugGapWitnessActive = False
    AADebugWrite "Gap", "replay clearance=" & CStr(clear)
TraceFailed:
    AADebugGapWitnessActive = False
    Err.Clear
End Sub

Public Sub AADebugGapEdgeRejected(ByVal distance As Double, ByVal gap As Double, _
                                  ByVal firstError As Double, ByVal secondError As Double, _
                                  ByVal tolerance As Double, ByVal ax As Double, ByVal ay As Double, _
                                  ByVal bx As Double, ByVal by As Double, ByVal cx As Double, _
                                  ByVal cy As Double, ByVal ex As Double, ByVal ey As Double)
    If Not AADebugGapWitnessActive Then Exit Sub
    AADebugWrite "Gap", "chord distance=" & Format$(distance, "0.000000000") & _
        "; required gap=" & Format$(gap, "0.000000000") & _
        "; chord errors=" & Format$(firstError, "0.000000000") & "/" & Format$(secondError, "0.000000000") & _
        "; tolerance=" & Format$(tolerance, "0.000000000") & _
        "; rejection cutoff=" & Format$(gap + firstError + secondError - tolerance, "0.000000000")
    AADebugWrite "Gap", "candidate chord raw A/B=" & Format$(ax, "0.000000000") & "/" & _
        Format$(ay, "0.000000000") & " -> " & Format$(bx, "0.000000000") & "/" & Format$(by, "0.000000000")
    AADebugWrite "Gap", "blocker chord raw A/B=" & Format$(cx, "0.000000000") & "/" & _
        Format$(cy, "0.000000000") & " -> " & Format$(ex, "0.000000000") & "/" & Format$(ey, "0.000000000")
End Sub

' Saat tracing detail aktif, hanya tiga contoh per kandidat yang membaca
' ulang kurva. Counter tetap mencakup semua penolakan dalam kandidat itu.
Public Sub AADebugMarkRejected(ByVal label As String, ByVal original As AAGeometry, _
                               ByVal moved As AAGeometry, ByVal obstacle As AAGeometry, _
                               ByVal dx As Double, ByVal dy As Double, ByVal markIndex As Long)
    If Not mSlotsActive Then Exit Sub
    If mMarkSamples >= 3 Then Exit Sub
    mMarkSamples = mMarkSamples + 1
    On Error GoTo TraceFailed
    AADebugWrite "Mark", label & "; obstacle=" & CStr(markIndex) & _
        "; shift X/Y=" & Format$(dx, "0.000000000") & "/" & Format$(dy, "0.000000000")
    AADebugWrite "Mark", "raw L/B/R/T=" & Format$(original.LeftX, "0.000000000") & "/" & _
        Format$(original.BottomY, "0.000000000") & "/" & Format$(original.RightX, "0.000000000") & "/" & _
        Format$(original.TopY, "0.000000000")
    AADebugWrite "Mark", "moved L/B/R/T=" & Format$(moved.LeftX, "0.000000000") & "/" & _
        Format$(moved.BottomY, "0.000000000") & "/" & Format$(moved.RightX, "0.000000000") & "/" & _
        Format$(moved.TopY, "0.000000000") & "; local center X/Y=" & _
        Format$((moved.LeftX + moved.RightX) / 2#, "0.000000000") & "/" & _
        Format$((moved.BottomY + moved.TopY) / 2#, "0.000000000")
    AADebugWrite "Mark", "obstacle L/B/R/T=" & Format$(obstacle.LeftX, "0.000000000") & "/" & _
        Format$(obstacle.BottomY, "0.000000000") & "/" & Format$(obstacle.RightX, "0.000000000") & "/" & _
        Format$(obstacle.TopY, "0.000000000")
    AADebugMarkGeometry moved, obstacle
TraceFailed:
    Err.Clear
End Sub

' Diagnostik terbatas oleh sample cap AADebugMarkRejected. Hasil pemeriksaan
' tambahan ini tidak dipakai untuk menerima atau menolak posisi.
Private Sub AADebugMarkGeometry(ByVal moved As AAGeometry, ByVal obstacle As AAGeometry)
    Dim first As Curve, second As Curve, joined As Curve
    AADebugWrite "Mark", "curve count object/mark=" & CStr(moved.CurveCount) & "/" & _
        CStr(obstacle.CurveCount)
    If moved.CurveCount <> 1 Or obstacle.CurveCount <> 1 Then Exit Sub
    Set first = moved.CurveAt(1)
    Set second = obstacle.CurveAt(1)
    Set joined = first.WeldWith(second)
    AADebugWrite "Mark", "area object/mark/union mm2=" & _
        Format$(Abs(first.Area), "0.000000000000") & "/" & _
        Format$(Abs(second.Area), "0.000000000000") & "/" & _
        Format$(Abs(joined.Area), "0.000000000000") & "; difference=" & _
        Format$(Abs(first.Area) + Abs(second.Area) - Abs(joined.Area), "0.000000000000")
    AADebugWrite "Mark", "IntersectsWith=" & CStr(first.IntersectsWith(second))
    AADebugMarkInsideNode "object node inside mark", first, second
    AADebugMarkInsideNode "mark node inside object", second, first
End Sub

Private Sub AADebugMarkInsideNode(ByVal label As String, ByVal source As Curve, ByVal region As Curve)
    Dim vertex As Node
    Dim i As Long
    For Each vertex In source.Nodes
        i = i + 1
        If region.IsPointInside(vertex.PositionX, vertex.PositionY) Then
            AADebugWrite "Mark", label & ": index=" & CStr(i) & "; local X/Y=" & _
                Format$(vertex.PositionX, "0.000000000") & "/" & _
                Format$(vertex.PositionY, "0.000000000")
            Exit Sub
        End If
    Next vertex
    AADebugWrite "Mark", label & ": none"
End Sub

Public Sub AADebugGapRejected(ByVal orientation As AARotationVariant, _
                              ByVal x As Double, ByVal y As Double, _
                              ByVal blocker As AAPlacement, _
                              ByVal h As Double, ByVal v As Double)
    Dim candidate As AAPlacement
    If Not mSlotsActive Then Exit Sub
    mGapRejected = mGapRejected + 1
    If mGapRejected > 3 Then Exit Sub
    On Error GoTo TraceFailed
    Set candidate = orientation.MakePlacement(0, -1, 0, x, y, h, v)
    AADebugWrite "Slot", "gap angle=" & CStr(orientation.Angle) & _
        "; x=" & Format$(x, "0.000000000") & "; y=" & Format$(y, "0.000000000") & _
        "; blocker angle=" & CStr(blocker.RotationDelta) & _
        "; placed clearance=" & CStr(candidate.HasClearance(blocker)) & _
        "; curve shift error=" & Format$(candidate.CollisionFootprint.LeftX - _
            x - orientation.InsetX, "0.000000000") & "/" & _
        Format$(candidate.CollisionFootprint.TopY - y + orientation.InsetY, "0.000000000")
TraceFailed:
    Err.Clear
End Sub

Public Sub AADebugOverlapRejected(ByVal candidate As AAPlacement, _
                                  ByVal blocker As AAPlacement)
    Dim first As Curve, second As Curve, joined As Curve
    If Not mSlotsActive Then Exit Sub
    mOverlapRejected = mOverlapRejected + 1
    If mOverlapRejected > 3 Then Exit Sub
    On Error GoTo TraceFailed
    AADebugWrite "Slot", "overlap angle=" & CStr(candidate.RotationDelta) & _
        "; x=" & Format$(candidate.Footprint.LeftX, "0.000000000") & _
        "; y=" & Format$(candidate.Footprint.TopY, "0.000000000") & _
        "; blocker angle=" & CStr(blocker.RotationDelta)
    If candidate.CollisionFootprint.CurveCount <> 1 Or _
       blocker.CollisionFootprint.CurveCount <> 1 Then Exit Sub
    Set first = candidate.CollisionFootprint.CurveAt(1)
    Set second = blocker.CollisionFootprint.CurveAt(1)
    Set joined = first.WeldWith(second)
    AADebugWrite "Slot", "union area difference=" & _
        Format$(Abs(first.Area) + Abs(second.Area) - Abs(joined.Area), "0.000000000000")
TraceFailed:
    Err.Clear
End Sub

' Hanya kegagalan pertama per ValidPlan; tersedia juga pada debug ringkas.
Public Sub AADebugValidationRejected(ByVal label As String, ByVal reason As String, _
                                     ByVal candidate As AAPlacement, ByVal blocker As AAPlacement)
    If Not mEnabled Then Exit Sub
    On Error GoTo TraceFailed
    AADebugWrite "Validate", label & "; rejected=" & reason
    AADebugValidationPlacement "candidate", candidate
    If Not blocker Is Nothing Then AADebugValidationPlacement "blocker", blocker
TraceFailed:
    Err.Clear
End Sub

Private Sub AADebugValidationPlacement(ByVal role As String, ByVal placed As AAPlacement)
    Dim box As AAGeometry, collision As AAGeometry
    Dim shiftErrorX As Double, shiftErrorY As Double
    Set box = placed.Footprint
    Set collision = placed.CollisionFootprint
    AADebugWrite "Validate", role & "; design index=" & CStr(placed.DesignIndex) & _
        "; cut index=" & CStr(placed.CutIndex) & "; area=" & CStr(placed.AreaIndex) & _
        "; angle=" & CStr(placed.RotationDelta) & _
        "; bounds L/B/R/T=" & Format$(box.LeftX, "0.000000000") & "/" & _
        Format$(box.BottomY, "0.000000000") & "/" & _
        Format$(box.RightX, "0.000000000") & "/" & Format$(box.TopY, "0.000000000")
    AADebugWrite "Validate", role & "; slot X/Y=" & _
        Format$(placed.SlotLeft, "0.000000000") & "/" & Format$(placed.SlotTop, "0.000000000") & _
        "; gap H/V=" & Format$(placed.GapH, "0.000000000") & "/" & Format$(placed.GapV, "0.000000000")
    If placed.UsesCenter Then
        shiftErrorX = collision.LeftX - placed.Contour.LeftX - placed.CenterX
        shiftErrorY = collision.TopY - placed.Contour.TopY - placed.CenterY
        AADebugWrite "Validate", role & "; curve shift error X/Y=" & _
            Format$(shiftErrorX, "0.000000000") & "/" & Format$(shiftErrorY, "0.000000000")
    End If
End Sub

Public Sub AADebugPlan(ByVal label As String, ByVal plan As AALayoutPlan)
    Dim widthMM As Double, heightMM As Double, boundsArea As Double
    Dim collisionWidth As Double, collisionHeight As Double, collisionArea As Double
    If Not mEnabled Then Exit Sub
    On Error GoTo TraceFailed
    If plan Is Nothing Then
        AADebugWrite "Plan", label & ": rejected (no valid plan)"
        Exit Sub
    End If
    PlanBounds plan, False, widthMM, heightMM, boundsArea
    PlanBounds plan, True, collisionWidth, collisionHeight, collisionArea
    AADebugWrite "Plan", label & ": areas=" & CStr(plan.AreaCount) & _
        "; objects=" & CStr(plan.DesignCount) & _
        "; regular=" & CStr(plan.HasRegularPattern) & _
        "; max footprint=" & Format$(widthMM, "0.000") & "x" & Format$(heightMM, "0.000") & _
        " mm; collision area=" & Format$(collisionArea, "0.000") & " mm2"
TraceFailed:
    Err.Clear
End Sub

Public Sub AADebugSettings(ByVal settings As AASettings)
    Dim h As String, v As String
    If Not mEnabled Then Exit Sub
    On Error GoTo TraceFailed
    h = "default": v = "default"
    If settings.HasHorizontalGap Then h = CStr(settings.HorizontalGap)
    If settings.HasVerticalGap Then v = CStr(settings.VerticalGap)
    If settings.HasHorizontalGap Then h = h & " mm"
    If settings.HasVerticalGap Then v = v & " mm"
    AADebugWrite "Input", "area=" & CStr(settings.AreaWidthMM) & "x" & _
        CStr(settings.AreaHeightMM) & " mm; gap H=" & h & "; V=" & v
    If settings.HasCustomRestrictions Then AADebugWrite "Custom Area", settings.CustomSummary
TraceFailed:
    Err.Clear
End Sub

Public Sub AADebugModel(ByVal index As Long, ByVal model As AAModel, ByVal quantity As Long)
    Dim orientation As AARotationVariant, h As Double, v As Double, angle As Double
    If Not AADebugDetailsEnabled Then Exit Sub
    On Error GoTo TraceFailed
    model.GetBoundaryGaps AA_MODE_STRAIGHT, h, v
    angle = 0#
    If Not model.AllowsRotation(angle) Then angle = 90#
    Set orientation = model.RotationVariant(angle)
    AADebugWrite "Model", "index=" & CStr(index) & "; quantity=" & CStr(quantity) & _
        "; angle=" & CStr(angle) & _
        "; size=" & Format$(orientation.Width, "0.000000") & "x" & _
        Format$(orientation.Height, "0.000000") & "; edges=" & CStr(orientation.Contour.EdgeCount) & _
        "; rectangle=" & CStr(model.IsRectangle) & "; triangle=" & CStr(model.IsTriangle) & _
        "; straight gap=" & Format$(h, "0.000000") & "/" & Format$(v, "0.000000")
TraceFailed:
    Err.Clear
End Sub

Private Sub PlanBounds(ByVal plan As AALayoutPlan, ByVal collision As Boolean, _
                       ByRef widthMM As Double, ByRef heightMM As Double, _
                       ByRef boundsArea As Double)
    Dim lefts() As Double, bottoms() As Double, rights() As Double, tops() As Double
    Dim seen() As Boolean, i As Long
    Dim placed As AAPlacement, box As AAGeometry
    If plan.AreaCount = 0 Then Exit Sub
    ReDim lefts(0 To plan.AreaCount - 1): ReDim bottoms(0 To plan.AreaCount - 1)
    ReDim rights(0 To plan.AreaCount - 1): ReDim tops(0 To plan.AreaCount - 1)
    ReDim seen(0 To plan.AreaCount - 1)
    For Each placed In plan.Placements
        i = placed.AreaIndex
        If collision Then Set box = placed.CollisionFootprint Else Set box = placed.Footprint
        If Not seen(i) Then
            lefts(i) = box.LeftX: bottoms(i) = box.BottomY
            rights(i) = box.RightX: tops(i) = box.TopY
            seen(i) = True
        Else
            If box.LeftX < lefts(i) Then lefts(i) = box.LeftX
            If box.BottomY < bottoms(i) Then bottoms(i) = box.BottomY
            If box.RightX > rights(i) Then rights(i) = box.RightX
            If box.TopY > tops(i) Then tops(i) = box.TopY
        End If
    Next placed
    For i = 0 To plan.AreaCount - 1
        If seen(i) Then
            If rights(i) - lefts(i) > widthMM Then widthMM = rights(i) - lefts(i)
            If tops(i) - bottoms(i) > heightMM Then heightMM = tops(i) - bottoms(i)
            boundsArea = boundsArea + (rights(i) - lefts(i)) * (tops(i) - bottoms(i))
        End If
    Next i
End Sub

' SIZE_T tidak bertanda; pada host 32-bit VBA membacanya sebagai Long bertanda.
#If VBA7 Then
Private Function AADebugMemoryBytes(ByVal value As LongPtr) As Double
#Else
Private Function AADebugMemoryBytes(ByVal value As Long) As Double
#End If
    AADebugMemoryBytes = CDbl(value)
#If Win64 Then
#Else
    If AADebugMemoryBytes < 0# Then _
        AADebugMemoryBytes = AADebugMemoryBytes + 4294967296#
#End If
End Function

Private Sub AADebugMemoryReset()
    mMemoryTracking = False: mMemoryTakeBaseline = False
    mMemoryHasBaseline = False
    mMemoryStartPrivate = 0#
    mMemoryMaxPrivate = 0#: mMemoryMaxWorking = 0#
    mMemorySamples = 0
End Sub

Private Sub AADebugMemoryStart()
    If Not mEnabled Then Exit Sub
    On Error Resume Next
    AADebugMemoryReset
    mMemoryTracking = True
    AADebugWrite "Mem", "scope=CorelDRAW process; unit=MiB; max=sampled per AA run; peakWS=process lifetime"
    ' Baseline hanya berasal dari awal Process, bukan sampel berikutnya jika gagal.
    mMemoryTakeBaseline = True
    AADebugMemory "Process start"
    mMemoryTakeBaseline = False
End Sub

Private Sub AADebugMemoryFinish(ByVal status As String)
    Dim delta As String
    If Not mMemoryTracking Then Exit Sub
    On Error Resume Next
    ' Engine masih memegang model/planner; ini bukan snapshot setelah pelepasan data.
    AADebugMemory "Process end (" & status & "; engine data may still be live)"
    delta = "n/a"
    If mMemoryHasBaseline Then delta = _
        Format$((mMemoryMaxPrivate - mMemoryStartPrivate) / 1048576#, "0.00") & " MiB"
    If mMemorySamples > 0 Then
        AADebugWrite "Mem", "summary: samples=" & CStr(mMemorySamples) & _
            "; max sampled private=" & Format$(mMemoryMaxPrivate / 1048576#, "0.00") & " MiB" & _
            "; max sampled working=" & Format$(mMemoryMaxWorking / 1048576#, "0.00") & " MiB" & _
            "; max sampled delta private=" & delta
    Else
        AADebugWrite "Mem", "summary: unavailable (no valid samples)"
    End If
    mMemoryTracking = False: mMemoryTakeBaseline = False
End Sub

Public Sub AADebugMemory(ByVal label As String)
    Dim mem As PROCESS_MEMORY_COUNTERS_EX
    Dim expectedSize As Long, apiError As Long, traceError As Long
    Dim traceDescription As String, delta As String
    Dim privateBytes As Double, workingBytes As Double, peakWorkingBytes As Double
    If Not mEnabled Then Exit Sub
    On Error GoTo MemoryUnavailable
    mem.cb = LenB(mem)
#If Win64 Then
    expectedSize = 80
#Else
    expectedSize = 44
#End If
    If mem.cb <> expectedSize Then
        AADebugWrite "Mem", label & "; unavailable; structure size=" & CStr(mem.cb) & _
            "; expected=" & CStr(expectedSize)
        Exit Sub
    End If
    If GetProcessMemoryInfo(GetCurrentProcess(), mem, mem.cb) = 0 Then
        apiError = Err.LastDllError
        AADebugWrite "Mem", label & "; unavailable; Win32 error=" & CStr(apiError)
        Exit Sub
    End If
    privateBytes = AADebugMemoryBytes(mem.PrivateUsage)
    workingBytes = AADebugMemoryBytes(mem.WorkingSetSize)
    peakWorkingBytes = AADebugMemoryBytes(mem.PeakWorkingSetSize)
    delta = "n/a"
    If mMemoryTracking Then
        mMemorySamples = mMemorySamples + 1
        If mMemoryTakeBaseline Then
            mMemoryStartPrivate = privateBytes
            mMemoryHasBaseline = True
        End If
        If privateBytes > mMemoryMaxPrivate Then mMemoryMaxPrivate = privateBytes
        If workingBytes > mMemoryMaxWorking Then mMemoryMaxWorking = workingBytes
        If mMemoryHasBaseline Then delta = _
            Format$((privateBytes - mMemoryStartPrivate) / 1048576#, "0.00") & " MiB"
    End If
    AADebugWrite "Mem", label & _
        "; private=" & Format$(privateBytes / 1048576#, "0.00") & " MiB" & _
        "; working=" & Format$(workingBytes / 1048576#, "0.00") & " MiB" & _
        "; process peakWS=" & Format$(peakWorkingBytes / 1048576#, "0.00") & " MiB" & _
        "; delta private=" & delta
    Exit Sub
MemoryUnavailable:
    traceError = Err.Number: traceDescription = Err.Description
    AADebugWrite "Mem", label & "; unavailable; VBA error=" & CStr(traceError) & " " & traceDescription
    Err.Clear
End Sub
