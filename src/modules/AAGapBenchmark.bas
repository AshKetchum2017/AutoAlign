Option Explicit
Option Private Module

' Diagnostik mandiri: tidak menyentuh dokumen, selection, registry atau profiler.
' Kedua kernel adalah salinan rumus; sinkronisasinya diperiksa harness.
#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private Const AA_BENCHMARK_CASES As Long = 4096
Private Const AA_BENCHMARK_BATCH_SECONDS As Double = 0.15
Private Const AA_BENCHMARK_MAX_PASSES As Long = 512
Private mBenchmarkRunning As Boolean, mUsesCounter As Boolean
Private mFrequency As Currency, mOrigin As Currency, mTimerOrigin As Double

' Immediate Window: AADebugBenchmarkGap
' A = empat akar, B = satu akar. Median ratio B/A < 1 berarti B lebih cepat.
Public Sub AADebugBenchmarkGap(Optional ByVal rounds As Long = 8)
    Dim data() As Double, timesA() As Double, timesB() As Double, ratios() As Double
    Dim passes As Long, i As Long, checked As Long
    Dim secondsA As Double, secondsB As Double, checksumA As Double, checksumB As Double
    Dim medianA As Double, medianB As Double, medianRatio As Double, order As String
    Dim errorNumber As Long, errorDescription As String
    If mBenchmarkRunning Or AAProgressRunning Then Exit Sub
    On Error GoTo BenchmarkFailed
    If rounds < 2 Or rounds > 20 Or rounds Mod 2 <> 0 Then
        Err.Raise vbObjectError + 5391, "AAGapBenchmark", "Jumlah putaran harus genap, antara 2 dan 20."
    End If
    mBenchmarkRunning = True
    BenchmarkStartClock
    BenchmarkCases data
    checked = BenchmarkValidate(data)
    BenchmarkWrite "scope=synthetic numeric kernel; A=four roots; B=one root; cases=" & _
        CStr(AA_BENCHMARK_CASES) & "; exact checks=" & CStr(checked)
    BenchmarkWrite "excludes contour/index/native overlap/Process; clock=" & BenchmarkClockName()
    ' Pemanasan kedua jalur dilakukan sebelum kalibrasi dan pengukuran.
    secondsA = BenchmarkRunA(data, 1, checksumA)
    secondsB = BenchmarkRunB(data, 1, checksumB)
    passes = 1
    Do
        secondsA = BenchmarkRunA(data, passes, checksumA)
        If secondsA >= AA_BENCHMARK_BATCH_SECONDS Or passes >= AA_BENCHMARK_MAX_PASSES Then Exit Do
        passes = passes * 2
    Loop
    BenchmarkWrite "rounds=" & CStr(rounds) & "; passes=" & CStr(passes) & _
        "; calls per sample=" & CStr(passes * AA_BENCHMARK_CASES)
    ReDim timesA(1 To rounds): ReDim timesB(1 To rounds): ReDim ratios(1 To rounds)
    For i = 1 To rounds
        If i Mod 2 = 1 Then
            order = "AB"
            secondsA = BenchmarkRunA(data, passes, checksumA)
            secondsB = BenchmarkRunB(data, passes, checksumB)
        Else
            order = "BA"
            secondsB = BenchmarkRunB(data, passes, checksumB)
            secondsA = BenchmarkRunA(data, passes, checksumA)
        End If
        If checksumA <> checksumB Then Err.Raise vbObjectError + 5392, _
            "AAGapBenchmark", "Checksum kernel berbeda; hasil waktu tidak valid."
        If secondsA <= 0# Or secondsB <= 0# Then Err.Raise vbObjectError + 5393, _
            "AAGapBenchmark", "Resolusi clock tidak cukup; hasil waktu tidak valid."
        timesA(i) = secondsA: timesB(i) = secondsB: ratios(i) = secondsB / secondsA
        BenchmarkWrite "round=" & CStr(i) & "; order=" & order & _
            "; A=" & Format$(secondsA, "0.000000") & " s; B=" & Format$(secondsB, "0.000000") & _
            " s; ratio B/A=" & Format$(ratios(i), "0.0000")
    Next i
    medianA = BenchmarkMedian(timesA): medianB = BenchmarkMedian(timesB)
    medianRatio = BenchmarkMedian(ratios)
    BenchmarkWrite "median A=" & Format$(medianA, "0.000000") & _
        " s; median B=" & Format$(medianB, "0.000000") & " s; median paired ratio B/A=" & _
        Format$(medianRatio, "0.0000") & "; B time change=" & Format$((medianRatio - 1#) * 100#, "0.00") & "%"
    BenchmarkWrite "OK; all distances/checksums equal; timing is not a Process speedup measurement"
    mBenchmarkRunning = False
    Exit Sub
BenchmarkFailed:
    errorNumber = Err.Number: errorDescription = Err.Description
    mBenchmarkRunning = False
    BenchmarkWrite "ERROR=" & CStr(errorNumber) & "; " & errorDescription
End Sub

Private Sub BenchmarkWrite(ByVal message As String)
    If AADebugEnabled Then
        AADebugWrite "Gap benchmark", message
    Else
        Debug.Print "[AA][Gap benchmark] " & message
    End If
End Sub

Private Sub BenchmarkStartClock()
    mUsesCounter = False
    mTimerOrigin = Timer
    On Error GoTo TimerOnly
    If QueryPerformanceFrequency(mFrequency) <> 0 Then
        If QueryPerformanceCounter(mOrigin) <> 0 Then mUsesCounter = True
    End If
TimerOnly:
End Sub

Private Function BenchmarkClock() As Double
    Dim finish As Currency
    If mUsesCounter Then
        If QueryPerformanceCounter(finish) = 0 Then Err.Raise vbObjectError + 5394, _
            "AAGapBenchmark", "Pembacaan clock QPC gagal."
        BenchmarkClock = CDbl(finish - mOrigin) / CDbl(mFrequency)
    Else
        BenchmarkClock = Timer - mTimerOrigin
        If BenchmarkClock < 0# Then BenchmarkClock = BenchmarkClock + 86400#
    End If
End Function

Private Function BenchmarkClockName() As String
    If mUsesCounter Then BenchmarkClockName = "QPC" Else BenchmarkClockName = "Timer fallback"
End Function

Private Function BenchmarkRandom(ByRef seed As Double) As Double
    seed = seed * 16807#
    seed = seed - Int(seed / 2147483647#) * 2147483647#
    BenchmarkRandom = seed / 2147483647#
End Function

Private Sub BenchmarkCases(ByRef data() As Double)
    Dim i As Long, column As Long, seed As Double
    Dim x As Double, y As Double, length As Double, gap As Double
    seed = 127#
    ReDim data(0 To 7, 0 To AA_BENCHMARK_CASES - 1)
    For i = 0 To AA_BENCHMARK_CASES - 1
        x = BenchmarkRandom(seed) * 2000# - 1000#
        y = BenchmarkRandom(seed) * 2000# - 1000#
        length = 0.01 + BenchmarkRandom(seed) * 10#
        gap = 0.00001 + BenchmarkRandom(seed) * 4#
        data(0, i) = x: data(1, i) = y: data(2, i) = x + length: data(3, i) = y
        data(4, i) = x: data(5, i) = y + gap
        data(6, i) = x + length: data(7, i) = y + gap
        Select Case i Mod 8
            Case 1 ' Kolinear terpisah.
                data(4, i) = x + length + gap: data(5, i) = y
                data(6, i) = x + 2# * length + gap: data(7, i) = y
            Case 2 ' Berpotongan; kedua kernel keluar sebelum proyeksi.
                data(4, i) = x + length / 2#: data(5, i) = y - gap
                data(6, i) = x + length / 2#: data(7, i) = y + gap
            Case 3 ' Salah satu segmen degenerat.
                data(2, i) = x: data(3, i) = y
            Case 4 ' Hampir paralel.
                data(7, i) = data(7, i) + 0.000000000001
            Case 5 ' Arah umum, tetap deterministik tanpa Rnd/Randomize.
                For column = 0 To 7
                    data(column, i) = x + BenchmarkRandom(seed) * length
                Next column
            Case 6 ' Kedua segmen degenerat.
                data(2, i) = x: data(3, i) = y
                data(6, i) = data(4, i): data(7, i) = data(5, i)
            Case 7 ' Dekat ambang nominal 2 mm setelah toleransi.
                data(5, i) = y + 1.9999 + (BenchmarkRandom(seed) - 0.5) * 0.00000000001
                data(7, i) = data(5, i)
        End Select
    Next i
End Sub

Private Function BenchmarkValidate(ByRef data() As Double) As Long
    Dim i As Long, a As Double, b As Double, vx As Double, vy As Double
    For i = 0 To UBound(data, 2)
        a = BenchmarkOldDistance(data(0, i), data(1, i), data(2, i), data(3, i), _
            data(4, i), data(5, i), data(6, i), data(7, i), vx, vy)
        b = BenchmarkNewDistance(data(0, i), data(1, i), data(2, i), data(3, i), _
            data(4, i), data(5, i), data(6, i), data(7, i), vx, vy)
        If a <> b Then Err.Raise vbObjectError + 5395, "AAGapBenchmark", _
            "Jarak kernel berbeda pada case " & CStr(i) & "; hasil waktu tidak valid."
        BenchmarkValidate = BenchmarkValidate + 1
    Next i
End Function

' Dua loop sengaja sama; tidak ada branch pemilihan kernel di loop terdalam.
' Logging/validasi/median di luar interval yang diukur. Tidak memompa event.
Private Function BenchmarkRunA(ByRef data() As Double, ByVal passes As Long, _
                               ByRef checksum As Double) As Double
    Dim pass As Long, i As Long, last As Long, started As Double
    Dim vx As Double, vy As Double
    checksum = 0#: last = UBound(data, 2)
    started = BenchmarkClock()
    For pass = 1 To passes
        For i = 0 To last
            checksum = checksum + BenchmarkOldDistance(data(0, i), data(1, i), data(2, i), data(3, i), _
                data(4, i), data(5, i), data(6, i), data(7, i), vx, vy)
        Next i
    Next pass
    BenchmarkRunA = BenchmarkClock() - started
End Function

Private Function BenchmarkRunB(ByRef data() As Double, ByVal passes As Long, _
                               ByRef checksum As Double) As Double
    Dim pass As Long, i As Long, last As Long, started As Double
    Dim vx As Double, vy As Double
    checksum = 0#: last = UBound(data, 2)
    started = BenchmarkClock()
    For pass = 1 To passes
        For i = 0 To last
            checksum = checksum + BenchmarkNewDistance(data(0, i), data(1, i), data(2, i), data(3, i), _
                data(4, i), data(5, i), data(6, i), data(7, i), vx, vy)
        Next i
    Next pass
    BenchmarkRunB = BenchmarkClock() - started
End Function

Private Function BenchmarkMedian(ByRef values() As Double) As Double
    Dim sorted() As Double, i As Long, j As Long, count As Long, value As Double
    count = UBound(values)
    ReDim sorted(1 To count)
    For i = 1 To count
        sorted(i) = values(i)
    Next i
    For i = 2 To count
        value = sorted(i): j = i - 1
        Do While j >= 1
            If sorted(j) <= value Then Exit Do
            sorted(j + 1) = sorted(j): j = j - 1
        Loop
        sorted(j + 1) = value
    Next i
    If count Mod 2 = 0 Then
        BenchmarkMedian = (sorted(count \ 2) + sorted(count \ 2 + 1)) / 2#
    Else
        BenchmarkMedian = sorted((count + 1) \ 2)
    End If
End Function

' BEGIN FROZEN KERNELS

Private Function BenchmarkOldDistance(ByVal ax As Double, ByVal ay As Double, _
                              ByVal bx As Double, ByVal by As Double, _
                              ByVal cx As Double, ByVal cy As Double, _
                              ByVal dx As Double, ByVal dy As Double, _
                              ByRef vx As Double, ByRef vy As Double) As Double
    Dim denominator As Double, t As Double, u As Double
    Dim best As Double, value As Double, px As Double, py As Double
    denominator = (bx - ax) * (dy - cy) - (by - ay) * (dx - cx)
    If Abs(denominator) > 0.000000000001 Then
        t = ((cx - ax) * (dy - cy) - (cy - ay) * (dx - cx)) / denominator
        u = ((cx - ax) * (by - ay) - (cy - ay) * (bx - ax)) / denominator
        If t >= 0# And t <= 1# And u >= 0# And u <= 1# Then
            vx = 0#: vy = 0#
            Exit Function
        End If
    End If
    best = BenchmarkOldPoint(ax, ay, cx, cy, dx, dy, px, py)
    vx = ax - px: vy = ay - py
    value = BenchmarkOldPoint(bx, by, cx, cy, dx, dy, px, py)
    If value < best Then best = value: vx = bx - px: vy = by - py
    value = BenchmarkOldPoint(cx, cy, ax, ay, bx, by, px, py)
    If value < best Then best = value: vx = px - cx: vy = py - cy
    value = BenchmarkOldPoint(dx, dy, ax, ay, bx, by, px, py)
    If value < best Then best = value: vx = px - dx: vy = py - dy
    BenchmarkOldDistance = best
End Function

Private Function BenchmarkOldPoint(ByVal x As Double, ByVal y As Double, _
                               ByVal ax As Double, ByVal ay As Double, _
                               ByVal bx As Double, ByVal by As Double, _
                               ByRef px As Double, ByRef py As Double) As Double
    Dim length2 As Double, t As Double
    length2 = (bx - ax) ^ 2 + (by - ay) ^ 2
    If length2 > 0# Then t = ((x - ax) * (bx - ax) + (y - ay) * (by - ay)) / length2
    If t < 0# Then t = 0#
    If t > 1# Then t = 1#
    px = ax + t * (bx - ax): py = ay + t * (by - ay)
    BenchmarkOldPoint = Sqr((x - px) ^ 2 + (y - py) ^ 2)
End Function

Private Function BenchmarkNewDistance(ByVal ax As Double, ByVal ay As Double, _
                              ByVal bx As Double, ByVal by As Double, _
                              ByVal cx As Double, ByVal cy As Double, _
                              ByVal dx As Double, ByVal dy As Double, _
                              ByRef vx As Double, ByRef vy As Double) As Double
    Dim denominator As Double, t As Double, u As Double
    Dim best As Double, value As Double, px As Double, py As Double
    denominator = (bx - ax) * (dy - cy) - (by - ay) * (dx - cx)
    If Abs(denominator) > 0.000000000001 Then
        t = ((cx - ax) * (dy - cy) - (cy - ay) * (dx - cx)) / denominator
        u = ((cx - ax) * (by - ay) - (cy - ay) * (bx - ax)) / denominator
        If t >= 0# And t <= 1# And u >= 0# And u <= 1# Then
            vx = 0#: vy = 0#
            Exit Function
        End If
    End If
    ' Akar bersifat monoton: min(Sqr(d2)) = Sqr(min(d2)). Pilih jarak
    ' kuadrat dahulu agar hanya satu akar diperlukan; hasil tetap dalam mm.
    best = BenchmarkNewPointSquared(ax, ay, cx, cy, dx, dy, px, py)
    vx = ax - px: vy = ay - py
    value = BenchmarkNewPointSquared(bx, by, cx, cy, dx, dy, px, py)
    If value < best Then best = value: vx = bx - px: vy = by - py
    value = BenchmarkNewPointSquared(cx, cy, ax, ay, bx, by, px, py)
    If value < best Then best = value: vx = px - cx: vy = py - cy
    value = BenchmarkNewPointSquared(dx, dy, ax, ay, bx, by, px, py)
    If value < best Then best = value: vx = px - dx: vy = py - dy
    BenchmarkNewDistance = Sqr(best)
End Function

Private Function BenchmarkNewPointSquared(ByVal x As Double, ByVal y As Double, _
                                      ByVal ax As Double, ByVal ay As Double, _
                                      ByVal bx As Double, ByVal by As Double, _
                                      ByRef px As Double, ByRef py As Double) As Double
    Dim length2 As Double, t As Double
    length2 = (bx - ax) ^ 2 + (by - ay) ^ 2
    If length2 > 0# Then t = ((x - ax) * (bx - ax) + (y - ay) * (by - ay)) / length2
    If t < 0# Then t = 0#
    If t > 1# Then t = 1#
    px = ax + t * (bx - ax): py = ay + t * (by - ay)
    BenchmarkNewPointSquared = (x - px) ^ 2 + (y - py) ^ 2
End Function
