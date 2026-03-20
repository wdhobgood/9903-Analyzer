Attribute VB_Name = "mod9903Analyzer"
Option Explicit

' ============================================================
' 9903 Analyzer v5 — Full Ship-PC Analysis
'   - Outputs ALL rows (not just reciprocal)
'   - Streams detail to sheets during processing (low memory)
'   - Per-entry aggregation on Summary
'   - Log sheet for non-reciprocal 98/99 with value > 0
' ============================================================

Private Const SHIPPC_PATH  As String = "C:\Data\9903 Analyzer\Input Reports\Ship-PC\"
Private Const ARCHIVE_PATH As String = "C:\Data\9903 Analyzer\Input Reports\Ship-PC\Archive\"
Private Const OUTPUT_PATH  As String = "C:\Data\9903 Analyzer\Output Reports\"

Private Const DETAIL_COLS  As Long = 18
Private Const CHUNK_SIZE   As Long = 50000
Private Const MAX_DATA_ROWS As Long = 1048575

' --- Settings ---
Private dictHTS        As Object   ' reciprocal HTS lookup
Private gLog           As String

' --- Running totals ---
Private totalWrongEV   As Double
Private totalWrongDuty As Double
Private totalCorrectEV As Double
Private totalCorrectDuty As Double
Private totalEVDiff    As Double
Private totalDutyDiff  As Double

' --- Per-entry aggregation: key=EntryID, val=Array(wrongEV, wrongDuty, correctEV, correctDuty, diffEV, diffDuty) ---
Private dictEntryAgg   As Object

' --- Log for non-reciprocal 98/99 with value > 0 ---
Private colLog         As Collection   ' each item = Array(entry, hts, mid, note)

' --- Streaming detail write state ---
Private wsDetail       As Worksheet    ' current detail sheet
Private wbDetail       As Workbook     ' current detail workbook (main or overflow)
Private wbMainOut      As Workbook     ' primary output wb — never close during overflow
Private detailWsRow    As Long         ' next row to write
Private detailRowsInFile As Long       ' data rows in current file
Private detailFileNum  As Long         ' file counter
Private detailTotalRows As Long        ' grand total detail rows
Private gOutStamp      As String       ' shared timestamp for all files

' --- Detail chunk buffer ---
Private chunkBuf()     As Variant
Private chunkPos       As Long         ' rows filled in current chunk
Private chunkSize      As Long         ' actual size of current chunk

' ================================================================
'  ENTRY POINT
' ================================================================
Public Sub Run_9903_Analyzer()

    Dim calcMode As XlCalculation
    On Error GoTo SafeExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    gLog = vbNullString
    Set dictHTS = CreateObject("Scripting.Dictionary")
    Set dictEntryAgg = CreateObject("Scripting.Dictionary")
    Set colLog = New Collection
    dictHTS.CompareMode = vbTextCompare
    dictEntryAgg.CompareMode = vbTextCompare

    totalWrongEV = 0#: totalWrongDuty = 0#
    totalCorrectEV = 0#: totalCorrectDuty = 0#
    totalEVDiff = 0#: totalDutyDiff = 0#
    detailTotalRows = 0

    ' progress form
    Dim frm As frmProgress
    Set frm = New frmProgress
    frm.Show vbModeless
    DoEvents

    ' 1 — load reciprocal HTS from Settings
    frm.UpdateStatus "Loading Settings...", 0
    LoadReciprocalHTS
    If dictHTS.Count = 0 Then
        gLog = gLog & "No HTS codes found in Settings sheet." & vbNewLine
        GoTo Finish
    End If

    ' 2 — collect Ship-PC files
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(SHIPPC_PATH) Then
        gLog = gLog & "Ship-PC folder not found: " & SHIPPC_PATH & vbNewLine
        GoTo Finish
    End If
    If Not fso.FolderExists(ARCHIVE_PATH) Then fso.CreateFolder ARCHIVE_PATH
    If Not fso.FolderExists(OUTPUT_PATH) Then fso.CreateFolder OUTPUT_PATH

    Dim folder As Object, f As Object
    Set folder = fso.GetFolder(SHIPPC_PATH)

    Dim filePaths As New Collection
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) Like "xls*" Then
            filePaths.Add f.Path
        End If
    Next f

    Dim totalFiles As Long
    totalFiles = filePaths.Count
    If totalFiles = 0 Then
        gLog = gLog & "No Ship-PC files found." & vbNewLine
        GoTo Finish
    End If

    ' 3 — create main output workbook and first Detail sheet
    gOutStamp = Format$(Now, "YYYYMMDD\_HHMM")

    Dim wbOut As Workbook
    Set wbOut = Workbooks.Add(xlWBATWorksheet)

    ' init detail streaming into main workbook
    Set wbDetail = wbOut
    Set wbMainOut = wbOut
    detailFileNum = 1
    Set wsDetail = wbOut.Sheets(1)
    wsDetail.Name = "Detail"
    WriteDetailHeaders wsDetail
    detailWsRow = 2
    detailRowsInFile = 0
    chunkPos = 0

    ' 4 — process each Ship-PC then archive
    Dim i As Long
    For i = 1 To totalFiles
        Dim fp As String, fn As String
        fp = CStr(filePaths(i))
        fn = fso.GetFileName(fp)

        frm.UpdateStatus "Ship-PC " & i & "/" & totalFiles & ": " & fn, _
                         CLng(((i - 1) / (totalFiles + 1)) * 100)

        ProcessShipPC fp

        ' archive
        On Error Resume Next
        If fso.FileExists(ARCHIVE_PATH & fn) Then fso.DeleteFile ARCHIVE_PATH & fn, True
        fso.MoveFile fp, ARCHIVE_PATH & fn
        If Err.Number <> 0 Then
            gLog = gLog & "Archive failed: " & fn & " — " & Err.Description & vbNewLine
            Err.Clear
        End If
        On Error GoTo SafeExit
    Next i

    ' 5 — flush any remaining chunk
    FlushChunk
    FormatDetailSheet wsDetail

    ' close last overflow workbook if not main
    If detailFileNum > 1 And Not wbDetail Is wbMainOut Then
        Application.DisplayAlerts = False
        wbDetail.SaveAs OUTPUT_PATH & "9903_Output_" & gOutStamp & "_Detail_" & detailFileNum & ".xlsx", xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        wbDetail.Close False
        Set wbDetail = Nothing
    End If

    ' 6 — write Summary (into main workbook)
    frm.UpdateStatus "Writing Summary...", CLng((totalFiles / (totalFiles + 1)) * 100)
    WriteSummary wbOut

    ' 7 — write Log sheet if any
    If colLog.Count > 0 Then WriteLog wbOut

    ' 8 — delete default blank sheets
    On Error Resume Next
    Application.DisplayAlerts = False
    Dim wsTemp As Worksheet
    For Each wsTemp In wbOut.Sheets
        If wsTemp.Name = "Sheet1" Then wsTemp.Delete: Exit For
    Next wsTemp
    Application.DisplayAlerts = True
    On Error GoTo SafeExit

    ' 9 — save main workbook
    Dim outName As String
    outName = "9903_Output_" & gOutStamp & ".xlsx"

    Application.DisplayAlerts = False
    wbOut.SaveAs OUTPUT_PATH & outName, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbOut.Close False

    frm.UpdateStatus "Complete!", 100
    Application.Wait Now + TimeSerial(0, 0, 1)

Finish:
    On Error Resume Next
    Unload frm: Set frm = Nothing
    On Error GoTo 0

    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Dim summaryMsg As String
    summaryMsg = "Detail rows: " & detailTotalRows & vbNewLine & _
                 "Unique entries: " & dictEntryAgg.Count & vbNewLine & _
                 "Log entries: " & colLog.Count
    If detailFileNum > 1 Then
        summaryMsg = summaryMsg & vbNewLine & "Detail split across " & detailFileNum & " files"
    End If

    If Len(gLog) > 0 Then
        MsgBox "Completed with warnings:" & vbNewLine & vbNewLine & _
               summaryMsg & vbNewLine & vbNewLine & _
               Left$(gLog, 600), vbExclamation, "9903 Analyzer"
    Else
        MsgBox "Completed successfully." & vbNewLine & vbNewLine & _
               summaryMsg & vbNewLine & _
               "Output: " & OUTPUT_PATH & outName, vbInformation, "9903 Analyzer"
    End If

    Set dictHTS = Nothing
    Set dictEntryAgg = Nothing
    Set colLog = Nothing
    Exit Sub

SafeExit:
    Dim errMsg As String
    errMsg = Err.Description
    On Error Resume Next
    Unload frm: Set frm = Nothing
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Fatal error: " & errMsg, vbCritical, "9903 Analyzer"
    Set dictHTS = Nothing
    Set dictEntryAgg = Nothing
    Set colLog = Nothing
End Sub

' ================================================================
'  LOAD RECIPROCAL HTS FROM SETTINGS
' ================================================================
Private Sub LoadReciprocalHTS()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then
        gLog = gLog & "Settings sheet not found." & vbNewLine
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    Dim arr As Variant
    If lastRow = 1 Then
        ReDim arr(1 To 1, 1 To 1): arr(1, 1) = ws.Cells(1, 1).Value
    Else
        arr = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Value
    End If

    Dim i As Long, htsVal As String
    For i = LBound(arr, 1) To UBound(arr, 1)
        htsVal = Trim$(CStr(arr(i, 1)))
        If Len(htsVal) > 0 Then
            If Not dictHTS.Exists(htsVal) Then dictHTS.Add htsVal, 1
        End If
    Next i
End Sub

' ================================================================
'  HELPERS
' ================================================================
Private Function SafeDbl(v As Variant) As Double
    If IsNumeric(v) Then SafeDbl = CDbl(v) Else SafeDbl = 0#
End Function

Private Function SafeCol(data As Variant, rowIdx As Long, colIdx As Long) As Variant
    If colIdx > 0 Then SafeCol = data(rowIdx, colIdx) Else SafeCol = Empty
End Function

' ================================================================
'  DETAIL STREAMING — flush chunk buffer to worksheet
' ================================================================
Private Sub FlushChunk()
    If chunkPos = 0 Then Exit Sub

    ' trim to actual rows filled
    Dim outArr() As Variant
    ReDim outArr(1 To chunkPos, 1 To DETAIL_COLS)
    Dim r As Long, c As Long
    For r = 1 To chunkPos
        For c = 1 To DETAIL_COLS
            outArr(r, c) = chunkBuf(r, c)
        Next c
    Next r

    wsDetail.Range(wsDetail.Cells(detailWsRow, 1), _
                   wsDetail.Cells(detailWsRow + chunkPos - 1, DETAIL_COLS)).Value = outArr
    detailWsRow = detailWsRow + chunkPos
    Erase outArr
    chunkPos = 0
End Sub

' ================================================================
'  DETAIL STREAMING — add one row, handle overflow
' ================================================================
Private Sub AddDetailRow(rowArr() As Variant)
    ' check if current file is full
    If detailRowsInFile >= MAX_DATA_ROWS Then
        ' flush remaining chunk
        FlushChunk
        FormatDetailSheet wsDetail

        ' save overflow workbook if not main
        If Not wbDetail Is wbMainOut Then
            Application.DisplayAlerts = False
            wbDetail.SaveAs OUTPUT_PATH & "9903_Output_" & gOutStamp & "_Detail_" & detailFileNum & ".xlsx", xlOpenXMLWorkbook
            Application.DisplayAlerts = True
            wbDetail.Close False
        End If

        ' start new overflow file
        detailFileNum = detailFileNum + 1
        Set wbDetail = Workbooks.Add(xlWBATWorksheet)
        Set wsDetail = wbDetail.Sheets(1)
        wsDetail.Name = "Detail (cont. " & detailFileNum & ")"
        WriteDetailHeaders wsDetail
        detailWsRow = 2
        detailRowsInFile = 0
    End If

    ' allocate chunk if needed
    If chunkPos = 0 Then
        Dim remaining As Long
        remaining = MAX_DATA_ROWS - detailRowsInFile
        If remaining > CHUNK_SIZE Then remaining = CHUNK_SIZE
        chunkSize = remaining
        ReDim chunkBuf(1 To chunkSize, 1 To DETAIL_COLS)
    End If

    ' write row into chunk
    chunkPos = chunkPos + 1
    Dim c As Long
    For c = 1 To DETAIL_COLS
        chunkBuf(chunkPos, c) = rowArr(c)
    Next c
    detailRowsInFile = detailRowsInFile + 1
    detailTotalRows = detailTotalRows + 1

    ' flush chunk if full
    If chunkPos >= chunkSize Then FlushChunk
End Sub

' ================================================================
'  ACCUMULATE PER-ENTRY TOTALS
' ================================================================
Private Sub AccumulateEntry(entryID As String, lnWrongEV As Double, lnWrongDuty As Double, _
                            lnCorrectEV As Double, lnCorrectDuty As Double, _
                            lnEVDiff As Double, lnDutyDiff As Double)
    If dictEntryAgg.Exists(entryID) Then
        Dim ex As Variant: ex = dictEntryAgg(entryID)
        ex(0) = ex(0) + lnWrongEV
        ex(1) = ex(1) + lnWrongDuty
        ex(2) = ex(2) + lnCorrectEV
        ex(3) = ex(3) + lnCorrectDuty
        ex(4) = ex(4) + lnEVDiff
        ex(5) = ex(5) + lnDutyDiff
        dictEntryAgg(entryID) = ex
    Else
        dictEntryAgg.Add entryID, Array(lnWrongEV, lnWrongDuty, lnCorrectEV, lnCorrectDuty, lnEVDiff, lnDutyDiff)
    End If
End Sub

' ================================================================
'  PROCESS ONE SHIP-PC FILE — ALL rows
' ================================================================
Private Sub ProcessShipPC(filePath As String)
    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
    Dim ws As Worksheet: Set ws = wbSrc.Sheets(1)

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then
        wbSrc.Close False: Exit Sub
    End If

    Dim data As Variant
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
    wbSrc.Close False

    ' --- detect headers ---
    Dim cExitDoc As Long, cReceiptDate As Long, cTxnDate As Long
    Dim cReceiptDocID As Long, cProductNum As Long, cOrderNumReceipt As Long
    Dim cMID As Long, cCountry As Long, cTxnQty As Long

    Dim seqHTSCol(1 To 4) As Long, seqValCol(1 To 4) As Long, seqRateCol(1 To 4) As Long
    Dim seqCount As Long: seqCount = 0

    Dim j As Long, hdr As String
    For j = 1 To lastCol
        hdr = Trim$(LCase$(CStr(data(1, j))))
        Select Case hdr
            Case "exitdocid", "exit doc id", "exitdoc_id"
                cExitDoc = j
            Case "receipt date", "receiptdate", "receipt_date"
                cReceiptDate = j
            Case "txndate", "txn date", "txn_date", "transactiondate", "transaction date"
                cTxnDate = j
            Case "receiptdocid", "receipt doc id", "receiptdoc_id", "receipt document id"
                cReceiptDocID = j
            Case "productnum", "product num", "product_num", "productnumber", "product number", "materialnumber", "material number"
                cProductNum = j
            Case "ordernumreceipt", "order num receipt", "ordernumber receipt", "ordernumberreceipt", "order_num_receipt"
                cOrderNumReceipt = j
            Case "manufacturerid", "manufacturer id", "manufacturer_id", "mid"
                cMID = j
            Case "countryoforigin", "country of origin", "country_of_origin", "countryorigin", "origin"
                cCountry = j
            Case "txnqty", "txn qty", "txn_qty", "transactionqty", "transaction qty", "quantity"
                cTxnQty = j
            Case "htssequence1_htsnum", "htssequence1_htsindex": seqHTSCol(1) = j
            Case "htssequence1_value":     seqValCol(1) = j
            Case "htssequence1_advaloremrate": seqRateCol(1) = j
            Case "htssequence2_htsnum", "htssequence2_htsindex": seqHTSCol(2) = j
            Case "htssequence2_value":     seqValCol(2) = j
            Case "htssequence2_advaloremrate": seqRateCol(2) = j
            Case "htssequence3_htsnum", "htssequence3_htsindex": seqHTSCol(3) = j
            Case "htssequence3_value":     seqValCol(3) = j
            Case "htssequence3_advaloremrate": seqRateCol(3) = j
            Case "htssequence4_htsnum", "htssequence4_htsindex": seqHTSCol(4) = j
            Case "htssequence4_value":     seqValCol(4) = j
            Case "htssequence4_advaloremrate": seqRateCol(4) = j
        End Select
    Next j

    ' determine available sequences
    Dim s As Long
    For s = 1 To 4
        If seqHTSCol(s) > 0 And seqValCol(s) > 0 And seqRateCol(s) > 0 Then
            seqCount = s
        Else
            Exit For
        End If
    Next s

    ' validate required
    If cExitDoc = 0 Or cMID = 0 Or seqCount < 1 Then
        gLog = gLog & "Missing required columns in " & Dir(filePath) & vbNewLine
        Exit Sub
    End If

    ' warn missing optional
    Dim missOpt As String
    If cReceiptDate = 0 Then missOpt = missOpt & "ReceiptDate "
    If cTxnDate = 0 Then missOpt = missOpt & "TxnDate "
    If cReceiptDocID = 0 Then missOpt = missOpt & "ReceiptDocID "
    If cProductNum = 0 Then missOpt = missOpt & "ProductNum "
    If cOrderNumReceipt = 0 Then missOpt = missOpt & "OrderNumReceipt "
    If cCountry = 0 Then missOpt = missOpt & "CountryOfOrigin "
    If cTxnQty = 0 Then missOpt = missOpt & "TxnQty "
    If Len(missOpt) > 0 Then
        gLog = gLog & "Optional cols not found in " & Dir(filePath) & ": " & missOpt & vbNewLine
    End If

    ' --- iterate ALL rows ---
    Dim i As Long
    For i = 2 To lastRow
        Dim entryID As String
        entryID = Trim$(CStr(data(i, cExitDoc)))
        If Len(entryID) = 0 Then GoTo NextRow

        Dim txnQty As Double
        If cTxnQty > 0 Then txnQty = SafeDbl(data(i, cTxnQty)) Else txnQty = 1

        ' --- scan ALL sequences: classify as reciprocal-tariff or MFN-merchandise ---
        ' Any HTS in Settings list ? known reciprocal (accumulate val/rate)
        ' Any HTS starting with 99/98 NOT in Settings ? other tariff provision (accumulate val/rate)
        ' First HTS not starting with 99/98 ? MFN merchandise
        Dim recipCount As Long: recipCount = 0
        Dim tariffCount As Long: tariffCount = 0
        Dim mfnIdx As Long: mfnIdx = 0

        ' reciprocal/tariff accumulators
        Dim sumRecipVal As Double: sumRecipVal = 0#
        Dim sumRecipRate As Double: sumRecipRate = 0#
        Dim recipHTS1 As String: recipHTS1 = vbNullString
        Dim recipHTS2 As String: recipHTS2 = vbNullString

        ' MFN values
        Dim mfnVal As Double: mfnVal = 0#
        Dim mfnRate As Double: mfnRate = 0#
        Dim mfnHTSStr As String: mfnHTSStr = vbNullString

        For s = 1 To seqCount
            Dim hStr As String: hStr = vbNullString
            If Not IsEmpty(data(i, seqHTSCol(s))) Then hStr = Trim$(CStr(data(i, seqHTSCol(s))))
            If Len(hStr) = 0 Then GoTo NextSeq

            If dictHTS.Exists(hStr) Then
                ' known reciprocal from Settings
                recipCount = recipCount + 1
                sumRecipVal = sumRecipVal + SafeDbl(data(i, seqValCol(s)))
                sumRecipRate = sumRecipRate + SafeDbl(data(i, seqRateCol(s)))
                If recipCount = 1 Then recipHTS1 = hStr
                If recipCount = 2 Then recipHTS2 = hStr

            ElseIf Left$(hStr, 2) = "99" Or Left$(hStr, 2) = "98" Then
                ' other tariff provision (S301, etc.) — treat as tariff, not merchandise
                tariffCount = tariffCount + 1
                sumRecipVal = sumRecipVal + SafeDbl(data(i, seqValCol(s)))
                sumRecipRate = sumRecipRate + SafeDbl(data(i, seqRateCol(s)))
                ' store in recipHTS slots if empty
                If recipCount + tariffCount = 1 Then recipHTS1 = hStr
                If recipCount + tariffCount = 2 Then recipHTS2 = hStr

                ' LOG: non-reciprocal 98/99 with value > 0
                Dim tVal As Double: tVal = SafeDbl(data(i, seqValCol(s)))
                If tVal > 0 Then
                    Dim midLog As String: midLog = vbNullString
                    If cMID > 0 Then midLog = Trim$(CStr(data(i, cMID)))
                    colLog.Add Array(entryID, hStr, midLog, _
                               "Value > 0 found on non-reciprocal HTS starting with " & Left$(hStr, 2))
                End If
            Else
                ' first real merchandise HTS = MFN
                If mfnIdx = 0 Then
                    mfnIdx = s
                    mfnHTSStr = hStr
                    mfnVal = SafeDbl(data(i, seqValCol(s)))
                    mfnRate = SafeDbl(data(i, seqRateCol(s)))
                End If
            End If
NextSeq:
        Next s

        ' fallback: if no non-99/98 MFN found, use first available sequence
        If mfnIdx = 0 Then
            For s = 1 To seqCount
                Dim fbStr As String: fbStr = vbNullString
                If Not IsEmpty(data(i, seqHTSCol(s))) Then fbStr = Trim$(CStr(data(i, seqHTSCol(s))))
                If Len(fbStr) > 0 Then
                    mfnIdx = s
                    mfnHTSStr = fbStr
                    mfnVal = SafeDbl(data(i, seqValCol(s)))
                    mfnRate = SafeDbl(data(i, seqRateCol(s)))
                    Exit For
                End If
            Next s
        End If

        ' --- calculations ---
        ' WrongEV  = (sum of all reciprocal values + MFN value)
        ' WrongDuty = WrongEV * (sumRecipRate + mfnRate)
        ' CorrectEV = MFN value only
        ' CorrectDuty = MFN value * (sumRecipRate + mfnRate)
        ' The rates still apply — only the entered value was wrong
        Dim totalVal As Double:    totalVal = sumRecipVal + mfnVal
        Dim totalRate As Double:   totalRate = sumRecipRate + mfnRate

        Dim wrongEV As Double:     wrongEV = totalVal
        Dim wrongDuty As Double:   wrongDuty = totalVal * totalRate
        Dim correctEV As Double:   correctEV = mfnVal
        Dim correctDuty As Double: correctDuty = mfnVal * totalRate

        Dim lnWrongEV As Double:     lnWrongEV = txnQty * wrongEV
        Dim lnWrongDuty As Double:   lnWrongDuty = txnQty * wrongDuty
        Dim lnCorrectEV As Double:   lnCorrectEV = txnQty * correctEV
        Dim lnCorrectDuty As Double: lnCorrectDuty = txnQty * correctDuty
        Dim lnEVDiff As Double:      lnEVDiff = lnWrongEV - lnCorrectEV
        Dim lnDutyDiff As Double:    lnDutyDiff = lnWrongDuty - lnCorrectDuty

        ' --- accumulate totals ---
        totalWrongEV = totalWrongEV + lnWrongEV
        totalWrongDuty = totalWrongDuty + lnWrongDuty
        totalCorrectEV = totalCorrectEV + lnCorrectEV
        totalCorrectDuty = totalCorrectDuty + lnCorrectDuty
        totalEVDiff = totalEVDiff + lnEVDiff
        totalDutyDiff = totalDutyDiff + lnDutyDiff

        ' --- accumulate per-entry ---
        AccumulateEntry entryID, lnWrongEV, lnWrongDuty, lnCorrectEV, lnCorrectDuty, lnEVDiff, lnDutyDiff

        ' --- build detail row and stream to sheet ---
        Dim rowArr(1 To 18) As Variant
        rowArr(1) = entryID
        rowArr(2) = SafeCol(data, i, cReceiptDocID)
        rowArr(3) = SafeCol(data, i, cTxnDate)
        rowArr(4) = SafeCol(data, i, cReceiptDate)
        rowArr(5) = SafeCol(data, i, cProductNum)
        rowArr(6) = SafeCol(data, i, cOrderNumReceipt)
        rowArr(7) = SafeCol(data, i, cMID)
        rowArr(8) = SafeCol(data, i, cCountry)
        rowArr(9) = txnQty
        rowArr(10) = recipHTS1
        rowArr(11) = recipHTS2
        rowArr(12) = mfnHTSStr
        rowArr(13) = lnWrongEV
        rowArr(14) = lnWrongDuty
        rowArr(15) = lnCorrectEV
        rowArr(16) = lnCorrectDuty
        rowArr(17) = lnEVDiff
        rowArr(18) = lnDutyDiff

        AddDetailRow rowArr

NextRow:
    Next i
End Sub

' ================================================================
'  WRITE SUMMARY SHEET
' ================================================================
Private Sub WriteSummary(wbOut As Workbook)
    Dim ws As Worksheet
    Set ws = wbOut.Sheets.Add(Before:=wbOut.Sheets(1))
    ws.Name = "Summary"

    ' --- SECTION 1: Grand totals ---
    ws.Cells(1, 1).Value = "9903 Reciprocal Tariff Analysis — Summary"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 13

    ws.Cells(3, 1).Value = "Total Entries Affected"
    ws.Cells(3, 2).Value = dictEntryAgg.Count

    ws.Cells(5, 1).Value = "Total Wrong EV"
    ws.Cells(5, 2).Value = totalWrongEV
    ws.Cells(6, 1).Value = "Total Wrong Duty"
    ws.Cells(6, 2).Value = totalWrongDuty

    ws.Cells(8, 1).Value = "Total Correct EV"
    ws.Cells(8, 2).Value = totalCorrectEV
    ws.Cells(9, 1).Value = "Total Correct Duty"
    ws.Cells(9, 2).Value = totalCorrectDuty

    ws.Cells(11, 1).Value = "Total EV Difference"
    ws.Cells(11, 2).Value = totalEVDiff
    ws.Cells(12, 1).Value = "Total Duty Difference"
    ws.Cells(12, 2).Value = totalDutyDiff

    ws.Range("A3,A5,A6,A8,A9,A11,A12").Font.Bold = True
    ws.Range("B3").NumberFormat = "#,##0"
    ws.Range("B5:B6,B8:B9,B11:B12").NumberFormat = "#,##0.00"

    ' --- SECTION 2: Per-entry breakdown ---
    Dim startRow As Long: startRow = 15
    ws.Cells(startRow, 1).Value = "Per-Entry Breakdown"
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow, 1).Font.Size = 12

    Dim hdrRow As Long: hdrRow = startRow + 1
    ws.Cells(hdrRow, 1).Value = "Entry Number"
    ws.Cells(hdrRow, 2).Value = "Total Wrong EV"
    ws.Cells(hdrRow, 3).Value = "Total Wrong Duty"
    ws.Cells(hdrRow, 4).Value = "Total Correct EV"
    ws.Cells(hdrRow, 5).Value = "Total Correct Duty"
    ws.Cells(hdrRow, 6).Value = "Total Diff EV"
    ws.Cells(hdrRow, 7).Value = "Total Diff Duty"
    ws.Rows(hdrRow).Font.Bold = True

    Dim entryCount As Long
    entryCount = dictEntryAgg.Count

    If entryCount > 0 Then
        Dim keys As Variant: keys = dictEntryAgg.keys
        Dim outArr() As Variant
        ReDim outArr(1 To entryCount, 1 To 7)

        Dim i As Long
        For i = 0 To entryCount - 1
            Dim agg As Variant: agg = dictEntryAgg(keys(i))
            outArr(i + 1, 1) = keys(i)
            outArr(i + 1, 2) = agg(0)   ' wrongEV
            outArr(i + 1, 3) = agg(1)   ' wrongDuty
            outArr(i + 1, 4) = agg(2)   ' correctEV
            outArr(i + 1, 5) = agg(3)   ' correctDuty
            outArr(i + 1, 6) = agg(4)   ' diffEV
            outArr(i + 1, 7) = agg(5)   ' diffDuty
        Next i

        ws.Range(ws.Cells(hdrRow + 1, 1), ws.Cells(hdrRow + entryCount, 7)).Value = outArr

        ' totals row
        Dim totRow As Long: totRow = hdrRow + entryCount + 1
        ws.Cells(totRow, 1).Value = "TOTAL"
        ws.Cells(totRow, 2).Value = totalWrongEV
        ws.Cells(totRow, 3).Value = totalWrongDuty
        ws.Cells(totRow, 4).Value = totalCorrectEV
        ws.Cells(totRow, 5).Value = totalCorrectDuty
        ws.Cells(totRow, 6).Value = totalEVDiff
        ws.Cells(totRow, 7).Value = totalDutyDiff
        ws.Rows(totRow).Font.Bold = True

        Erase outArr
    End If

    ' formatting
    ws.Columns("B:G").NumberFormat = "#,##0.00"
    ws.Columns("A:A").ColumnWidth = 28
    ws.Columns("B:G").ColumnWidth = 20
End Sub

' ================================================================
'  WRITE LOG SHEET
' ================================================================
Private Sub WriteLog(wbOut As Workbook)
    Dim ws As Worksheet
    Set ws = wbOut.Sheets.Add(After:=wbOut.Sheets(wbOut.Sheets.Count))
    ws.Name = "Log"

    ws.Cells(1, 1).Value = "Entry Number"
    ws.Cells(1, 2).Value = "HTS"
    ws.Cells(1, 3).Value = "MID"
    ws.Cells(1, 4).Value = "Note"
    ws.Rows(1).Font.Bold = True

    Dim logCount As Long: logCount = colLog.Count
    If logCount = 0 Then Exit Sub

    Dim outArr() As Variant
    ReDim outArr(1 To logCount, 1 To 4)

    Dim r As Long: r = 0
    Dim item As Variant
    For Each item In colLog
        r = r + 1
        outArr(r, 1) = item(0)
        outArr(r, 2) = item(1)
        outArr(r, 3) = item(2)
        outArr(r, 4) = item(3)
    Next item

    ws.Range(ws.Cells(2, 1), ws.Cells(logCount + 1, 4)).Value = outArr
    ws.Columns.AutoFit
    Erase outArr
End Sub

' ================================================================
'  DETAIL SHEET HELPERS
' ================================================================
Private Sub WriteDetailHeaders(ws As Worksheet)
    Dim hdr As Variant
    hdr = Array("Entry Number", "ReceiptDocID", "TxnDate", "Receipt Date", _
                "Material Number", "OrderNumReceipt", "MID", "CountryOfOrigin", _
                "TxnQty", "Reciprocal HTS 1", "Reciprocal HTS 2", "MFN HTS", _
                "Wrong EV", "Wrong Duty", "Correct EV", "Correct Duty", _
                "EV Difference", "Duty Difference")
    Dim c As Long
    For c = 0 To UBound(hdr)
        ws.Cells(1, c + 1).Value = hdr(c)
    Next c
    ws.Rows(1).Font.Bold = True
End Sub

Private Sub FormatDetailSheet(ws As Worksheet)
    ws.Columns("C:D").NumberFormat = "MM/DD/YYYY"
    ws.Columns("I:I").NumberFormat = "#,##0"
    ws.Columns("M:R").NumberFormat = "#,##0.00"
    ws.Columns.AutoFit
End Sub

