Option Explicit

Public Sub RunLayout(ByVal layoutMode As Long, ByRef feedback As String, _
                     ByRef feedbackStyle As VbMsgBoxStyle)
    Dim doc As Document
    Dim targetPage As Page
    Dim selectedShapes As ShapeRange
    Dim sourceShape As Shape
    Dim pageShape As Shape
    Dim marker As Shape
    Dim resultGroup As Shape
    Dim model As AAModel
    Dim previousUnit As cdrUnit
    Dim unitChanged As Boolean
    Dim commandGroupOpen As Boolean
    Dim mutationStarted As Boolean
    Dim rollbackSucceeded As Boolean
    Dim sourceOnPage As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim operation As String

    feedback = vbNullString
    feedbackStyle = vbExclamation
    On Error GoTo RunFailed

    operation = "Membaca dokumen aktif"
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo RunFailed
    If doc Is Nothing Then
        feedback = "Tidak ada dokumen aktif."
        Exit Sub
    End If

    Set targetPage = doc.ActivePage
    On Error Resume Next
    Set selectedShapes = ActiveSelectionRange
    On Error GoTo RunFailed
    If selectedShapes Is Nothing Then
        feedback = "Pilih tepat satu ellipse terlebih dahulu."
        Exit Sub
    End If
    If selectedShapes.Count <> 1 Then
        feedback = "Pilih tepat satu ellipse terlebih dahulu."
        Exit Sub
    End If

    Set sourceShape = selectedShapes(1)
    If sourceShape.Type <> cdrEllipseShape Then
        feedback = "Objek yang dipilih harus berupa ellipse."
        Exit Sub
    End If

    For Each pageShape In targetPage.Shapes
        If pageShape Is sourceShape Then
            sourceOnPage = True
            Exit For
        End If
    Next pageShape
    If Not sourceOnPage Then
        feedback = "Pilih ellipse langsung pada halaman aktif, bukan anggota group."
        Exit Sub
    End If

    operation = "Membaca unit dokumen"
    previousUnit = doc.Unit

    operation = "Memulai CommandGroup"
    doc.BeginCommandGroup "Auto Align"
    commandGroupOpen = True

    operation = "Mengatur unit milimeter"
    If previousUnit <> cdrMillimeter Then
        doc.Unit = cdrMillimeter
        unitChanged = True
    End If

    operation = "Merencanakan posisi ellipse"
    Set model = New AAModel
    model.Prepare layoutMode, sourceShape.SizeWidth, sourceShape.SizeHeight
    If model.PositionCount < 2 Then
        feedback = "Ukuran ellipse tidak memungkinkan duplikat di dalam batas 320 x 470 mm."
        GoTo CleanExit
    End If

    operation = "Membuat marker"
    Set marker = sourceShape.Layer.CreateRectangle( _
        targetPage.CenterX - model.MarkerWidthMM / 2#, _
        targetPage.CenterY + model.MarkerHeightMM / 2#, _
        targetPage.CenterX + model.MarkerWidthMM / 2#, _
        targetPage.CenterY - model.MarkerHeightMM / 2#)
    mutationStarted = True

    operation = "Menyusun dan menggroup ellipse"
    Set resultGroup = AAArrangeEllipse(sourceShape, marker, targetPage, model.Positions)

    operation = "Menghapus marker"
    marker.Delete
    Set marker = Nothing

    operation = "Memusatkan group pada halaman"
    resultGroup.CenterX = targetPage.CenterX
    resultGroup.CenterY = targetPage.CenterY

    operation = "Mengurutkan Z-Order anggota group"
    AAOrderGroupMembers resultGroup

    operation = "Memilih group hasil"
    resultGroup.CreateSelection

CleanExit:
    operation = "Memulihkan unit dokumen"
    If unitChanged Then
        doc.Unit = previousUnit
        unitChanged = False
    End If

    operation = "Menutup CommandGroup"
    If commandGroupOpen Then
        doc.EndCommandGroup
        commandGroupOpen = False
    End If
    Exit Sub

RunFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If commandGroupOpen Then
        Err.Clear
        doc.EndCommandGroup
        If Err.Number = 0 Then
            commandGroupOpen = False
            If mutationStarted Then
                Err.Clear
                doc.Undo
                rollbackSucceeded = (Err.Number = 0)
            Else
                rollbackSucceeded = True
            End If
        End If
    End If
    If unitChanged Then doc.Unit = previousUnit
    On Error GoTo 0

    feedbackStyle = vbCritical
    feedback = "Auto Align gagal saat " & operation & "." & vbCrLf & _
        "Error " & CStr(errorNumber) & " (" & errorSource & "): " & errorDescription
    If commandGroupOpen Then
        feedback = feedback & vbCrLf & _
            "CommandGroup tidak berhasil ditutup; periksa Undo dan dokumen."
    ElseIf rollbackSucceeded Then
        feedback = feedback & vbCrLf & "Perubahan telah dibatalkan."
    Else
        feedback = feedback & vbCrLf & "Periksa hasil di dokumen sebelum melanjutkan."
    End If
End Sub
