Attribute VB_Name = "ModuleSAP"

Option Explicit
Private mEDCLastDue As Collection
Private mEDCLastDueAt As Double
Private mEDCReferenceCache As Object
Private mEDCGroupTargets As Object
Private mEDCGroupDue As Object
Private mEDCGroupSTO As String
Private mEDCPlantSelected As Boolean
Private mEDCSelectedMat As String
Private mEDCSelectedSTO As String
Private mEDCSupplySelected As Boolean
Private mEDCClassicCols As Object
Private mEDCStoCol As String
Private mEDCItemCol As String
Private mEDCMatCol As String
Private mEDCPlantCol As String
Private mEDCUomCol As String
Private mEDCSession As Object
Private mEDCQty As Object
Private mEDCCheck As Object
Private mEDCIdentity As String
Private mEDCGrid As Object
Private mEDCRow As Long
Private mEDCQtyCol As String

Public LastApplyState As String
Private mDirectPosting As Boolean

Public LastConflictEligible As Boolean

Private mSession As Object

Private mOverviewIndex As Object

Private mUnplannedOriginalSnapshot As Object

Private mSavedForReuse As String

Private mWaitSeconds As Double

Private mMeasureWait As Boolean

Private mFieldPaths As Object

Private mControlPaths As Object

Public LastConflictKind As String

Public LastPostingTimingText As String

Public LastPostingTotalSeconds As Double

Public LastPostingFailed As Boolean

Public LastPostingErrorStage As String

Public LastPostingErrorText As String

#If VBA7 Then

Private Declare PtrSafe Sub SAPX_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

#Else

Private Declare Sub SAPX_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

#End If

Public Sub OB_ResetSession()

LastApplyState = ""

mSavedForReuse = "": mMeasureWait = False

Set mFieldPaths = Nothing

Set mControlPaths = Nothing

Set mSession = Nothing

Set mOverviewIndex = Nothing

End Sub

Public Function NormalizeHeader(ByVal s As String) As String

    Dim r As String, i As Long, ch As String, mapped As String

    Dim code As Long

    For i = 1 To Len(s)

        ch = mid$(s, i, 1)

        code = AscW(ch)

        If code >= 768 And code <= 879 Then

        ElseIf ch = " " Or ch = "." Or ch = "_" Or ch = "-" Then

        Else

            mapped = StripVNChar(AscW(LCase$(ch)))

            If mapped <> "" Then

                r = r & mapped

            Else

                r = r & ch

            End If

        End If

    Next i

    NormalizeHeader = UCase$(r)

End Function

Private Function StripVNChar(ByVal code As Long) As String

    Select Case code

        Case &HE0, &HE1, &H1EA3, &HE3, &H1EA1, &H103, &H1EB1, &H1EAF, &H1EB3, &H1EB5, &H1EB7, _
             &HE2, &H1EA7, &H1EA5, &H1EA9, &H1EAB, &H1EAD

            StripVNChar = "a"

        Case &HE8, &HE9, &H1EBB, &H1EBD, &H1EB9, &HEA, &H1EC1, &H1EBF, &H1EC3, &H1EC5, &H1EC7

            StripVNChar = "e"

        Case &HEC, &HED, &H1EC9, &H129, &H1ECB

            StripVNChar = "i"

        Case &HF2, &HF3, &H1ECF, &HF5, &H1ECD, &HF4, &H1ED3, &H1ED1, &H1ED5, &H1ED7, &H1ED9, _
             &H1A1, &H1EDD, &H1EDB, &H1EDF, &H1EE1, &H1EE3

            StripVNChar = "o"

        Case &HF9, &HFA, &H1EE7, &H169, &H1EE5, &H1B0, &H1EEB, &H1EE9, &H1EED, &H1EEF, &H1EF1

            StripVNChar = "u"

        Case &H1EF3, &HFD, &H1EF7, &H1EF9, &H1EF5

            StripVNChar = "y"

        Case &H111

            StripVNChar = "d"

        Case Else

            StripVNChar = ""

    End Select

End Function

Public Function DownloadPostingForOrder(Optional ByVal showMessage As Boolean = True) As Boolean

Dim ses As Object, products As Object, rows As Collection, p As Variant, grid As Object

Dim cProd As String, cQty As String, cBatch As String, cBin As String, cTst As String, colDate As String, cTime As String, cTotal As String

Dim cDestBin As String, cQtyPost As String, cDestBatch As String, cST As String, cProcType As String, cAvlBU As String

Dim cOwner As String, ownerExpected As String, selectedBatches As Object

Dim cUom As String, uomLastProduct As String, uomValue As String, uomCarry As String, expectedPostingUom As String

Dim r As Long, n As Long, rowCount As Long, lastP As String, pv As String, qv As String, bv As String, tv As String, oneRow As Variant

Dim ws As Worksheet, oldLast As Long, outArr() As Variant, i As Long, msg As String, stage As String

Dim fastColIds As Collection, fastNormIds As Object, fastNormCaps As Object

Dim totalStarted As Double, stageStarted As Double

Dim openSeconds As Double, exposeSeconds As Double, readSeconds As Double, commitSeconds As Double

On Error GoTo Fail

LastPostingTimingText = ""
mDirectPosting = False

LastPostingTotalSeconds = 0#

LastPostingFailed = False

LastPostingErrorStage = ""

LastPostingErrorText = ""

totalStarted = Timer

stage = "READ_ORDER"

Set products = ModuleMain.InputProducts()
Set selectedBatches = ModuleMain.PostingBatches(products)

If products.count = 0 Then Err.Raise vbObjectError + 2900, , "ORDER khong co Product hop le."

Set rows = New Collection

stage = "OPEN_POST"

Set ses = SAPX_GetSessionForTcode(ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST"))

If ses Is Nothing Then Err.Raise vbObjectError + 2901, , "Khong ket noi duoc SAP session."

stage = "OPEN_POST_SEARCH"

Application.StatusBar = "POSTING: dang loc " & products.count & " Product..."

stageStarted = Timer

' RC20-DEF-04: Route both 1 and multiple products through Advanced Search to enforce Owner (R8)

If products.count = 1 And selectedBatches.count = 0 Then
 p = products.keys
 If Not POST_OpenProductResult(ses, CStr(p(0)), grid, msg) Then
  ModuleLog.WriteLog "POST_SEARCH_ROUTE", "DIRECT unavailable; use Advanced Owner search: " & msg
  If Not POST_OpenAdvancedProductResult(ses, products, grid, msg) Then Err.Raise vbObjectError + 2902, , msg
 End If
Else
 If Not POST_OpenAdvancedProductResult(ses, products, grid, msg) Then Err.Raise vbObjectError + 2902, , msg
End If

If grid Is Nothing Then Err.Raise vbObjectError + 2903, , "Advanced Search khong tra result grid."

openSeconds = SAPX_Elapsed(stageStarted)

stage = "MAP_COLUMNS"

' FAST: doc ColumnOrder + caption dung 1 lan, cac cot con lai resolve trong RAM.

SAPX_BuildColumnCatalogFast grid, fastColIds, fastNormIds, fastNormCaps

cProd = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_PRODUCT", ""), "MATNR,PRODUCT,MATERIAL", True, fastColIds, fastNormIds, fastNormCaps)

cQty = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_AQTY", ""), "AVL_ALT_QTY,AQUAN,AQTY,A.QTY AUOM,AVAILABLE QTY,AVAILABLE QUANTITY,AVLQTY", True, fastColIds, fastNormIds, fastNormCaps)

cBatch = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_BATCH", ""), "CHARG,BATCH", True, fastColIds, fastNormIds, fastNormCaps)

cBin = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_SOURCE_BIN", ""), "VLPLA,LGPLA,SOURCE STORAGE BIN,SOURCE BIN,STORAGE BIN", True, fastColIds, fastNormIds, fastNormCaps)

cTst = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_TST", ""), "CAT,STOCK TYPE,ST", True, fastColIds, fastNormIds, fastNormCaps)

colDate = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_GR_DATE", ""), "UI_WDATU,WDATU,GR DATE,GOODS RECEIPT DATE", True, fastColIds, fastNormIds, fastNormCaps)

cTime = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_GR_TIME", ""), "UI_WDATT,WTIME,GR TIME,GOODS RECEIPT TIME", True, fastColIds, fastNormIds, fastNormCaps)

cTotal = SAPX_ResolveColumnFast("", "QUAN,TOTAL STK,TOTAL STOCK", False, fastColIds, fastNormIds, fastNormCaps)

cDestBin = SAPX_ResolveColumnFast("", "NLPLA,DEST. BIN,DEST BIN,DESTINATION STORAGE BIN", False, fastColIds, fastNormIds, fastNormCaps)

cQtyPost = SAPX_ResolveColumnFast("", "POST_QUAN,QTYPOSTCHG,QTY POST CHG", False, fastColIds, fastNormIds, fastNormCaps)

cDestBatch = SAPX_ResolveColumnFast("", "TO_CHARG,DEST.BATCH,DEST BATCH", False, fastColIds, fastNormIds, fastNormCaps)

cST = cTst

ModuleLog.WriteLog "POST_SOURCE_TYPE_COLUMN", cTst

cProcType = SAPX_ResolveColumnFast("", "PROCTY,PROC. TYPE,PROCESS TYPE", False, fastColIds, fastNormIds, fastNormCaps)

cAvlBU = SAPX_ResolveColumnFast("", "AVAIL_QUAN,AVLQTYBUOM,AVAILABLE QTY BUOM", False, fastColIds, fastNormIds, fastNormCaps)

stage = "EXPOSE_RESULT_GRID"

stageStarted = Timer

If Not POST_ExposeResultRows(ses, grid, rowCount, msg) Then

    ModuleLog.WriteLog "CANH BAO", "POSTING chua du de copy: " & msg, "POSTING"

    Application.StatusBar = "POSTING: chua tai du; POSTING cu duoc giu nguyen."

    If showMessage Then MsgBox "POSTING chua tai du nen CHUA copy vao Excel." & vbCrLf & vbCrLf & _
        msg & vbCrLf & vbCrLf & "He thong da keo danh sach SAP xuong tung trang de tai them. " & _
        "Hay tai lai sau khi SAP hien du du lieu. POSTING cu duoc giu nguyen.", vbExclamation, "TAI POSTING"

    LastPostingFailed = True

    LastPostingErrorStage = "EXPOSE_RESULT_GRID"

    LastPostingErrorText = msg

    LastPostingTotalSeconds = SAPX_Elapsed(totalStarted)

    LastPostingTimingText = "POSTING FAIL Expose; Total=" & Format$(LastPostingTotalSeconds, "0.00") & "s"

    Exit Function

End If

exposeSeconds = SAPX_Elapsed(stageStarted)

cUom = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_UOM", "ALTME"), "ALTME,AUOM,A.UOM,ALTERNATIVE UNIT OF MEASURE", True, fastColIds, fastNormIds, fastNormCaps)

cOwner = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_OWNER", ""), "OWNER,ENTITLED,ENTITLED TO DISPOSE,PARTNER", False, fastColIds, fastNormIds, fastNormCaps)

ownerExpected = Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R8").Value2))

If ownerExpected = "" Then Err.Raise vbObjectError + 2900, , "Owner (R8) bat buoc."

expectedPostingUom = UCase$(ModuleHelper.ToolConfig("EXPECTED_UOM", "UN"))

stage = "READ_RESULT_GRID"

stageStarted = Timer

If Not POST_ReadResultRows(ses, grid, products, cProd, cQty, cBatch, cBin, cTst, _
    colDate, cTime, cTotal, cDestBin, cQtyPost, cDestBatch, cST, cProcType, _
    cAvlBU, cUom, expectedPostingUom, rowCount, rows, msg, cOwner, ownerExpected) Then

    ModuleLog.WriteLog "CANH BAO", "POSTING chua du de copy: " & msg, "POSTING"

    Application.StatusBar = "POSTING: chua tai du; POSTING cu duoc giu nguyen."

    If showMessage Then MsgBox "POSTING chua tai du nen CHUA copy vao Excel." & vbCrLf & vbCrLf & _
        msg & vbCrLf & vbCrLf & "He thong da keo danh sach SAP xuong tung trang de tai them. " & _
        "Hay tai lai sau khi SAP hien du du lieu. POSTING cu duoc giu nguyen.", vbExclamation, "TAI POSTING"

    LastPostingFailed = True

    LastPostingErrorStage = "READ_RESULT_GRID"

    LastPostingErrorText = msg

    LastPostingTotalSeconds = SAPX_Elapsed(totalStarted)

    LastPostingTimingText = "POSTING FAIL Read; Open=" & Format$(openSeconds, "0.00") & _
        "s; Expose=" & Format$(exposeSeconds, "0.00") & "s; Total=" & _
        Format$(LastPostingTotalSeconds, "0.00") & "s"

    Exit Function

End If

readSeconds = SAPX_Elapsed(stageStarted)

stage = "COMMIT_POSTING"

stageStarted = Timer

Set ws = ThisWorkbook.Worksheets("POSTING_RAW")

oldLast = POST_LastDataRow(ws, 17)

If oldLast < 2 Then oldLast = 2

If rows.count = 0 Then

    ws.Range("A2:Q" & CStr(oldLast)).ClearContents

    commitSeconds = SAPX_Elapsed(stageStarted)

    LastPostingTotalSeconds = SAPX_Elapsed(totalStarted)

    LastPostingTimingText = "POSTING Time Open=" & Format$(openSeconds, "0.00") & _
        "s; Expose=" & Format$(exposeSeconds, "0.00") & "s; Read=" & _
        Format$(readSeconds, "0.00") & "s; Commit=" & Format$(commitSeconds, "0.00") & _
        "s; Total=" & Format$(LastPostingTotalSeconds, "0.00") & "s; Rows=0"

    DownloadPostingForOrder = True

    Application.StatusBar = "POSTING: SAP tra 0 dong ton cho ORDER."

    If showMessage Then MsgBox "SAP da tim xong nhung khong co dong ton hop le cho ORDER. POSTING cu da duoc xoa de tranh dung ton cu.", vbInformation, "TAI POSTING"

    Exit Function

End If

ReDim outArr(1 To rows.count, 1 To 17)

For i = 1 To rows.count

    oneRow = rows(i)

    For n = 1 To 17: outArr(i, n) = oneRow(n - 1): Next n

Next i

ws.Range("A2:Q" & CStr(oldLast)).ClearContents

ws.Range("C2:C" & CStr(rows.count + 1)).NumberFormat = "@"

ws.Range("I2:I" & CStr(rows.count + 1)).NumberFormat = "@"

ws.Range("A2").Resize(rows.count, 17).Value2 = outArr

ws.Range("B2:B" & CStr(rows.count + 1)).NumberFormat = "General"

ws.Range("F2:F" & CStr(rows.count + 1)).NumberFormat = "General"

ws.Range("H2:H" & CStr(rows.count + 1)).NumberFormat = "General"

ws.Range("N2:N" & CStr(rows.count + 1)).NumberFormat = "General"

ws.Range("Q2:Q" & CStr(rows.count + 1)).NumberFormat = "General"

commitSeconds = SAPX_Elapsed(stageStarted)

LastPostingTotalSeconds = SAPX_Elapsed(totalStarted)

LastPostingTimingText = "POSTING Time Open=" & Format$(openSeconds, "0.00") & _
    "s; Expose=" & Format$(exposeSeconds, "0.00") & "s; Read=" & _
    Format$(readSeconds, "0.00") & "s; Commit=" & Format$(commitSeconds, "0.00") & _
    "s; Total=" & Format$(LastPostingTotalSeconds, "0.00") & "s; Rows=" & CStr(rows.count)

DownloadPostingForOrder = True

Application.StatusBar = "POSTING: da tai " & rows.count & " dong. " & LastPostingTimingText

If showMessage Then MsgBox "Da tai " & rows.count & " dong POSTING.", vbInformation, "TAI POSTING"

Exit Function

Fail:

LastPostingFailed = True

LastPostingErrorStage = stage

LastPostingErrorText = Err.Description

LastPostingTotalSeconds = SAPX_Elapsed(totalStarted)

LastPostingTimingText = "POSTING FAIL stage=" & stage & "; Total=" & _
    Format$(LastPostingTotalSeconds, "0.00") & "s"

Application.StatusBar = False

ModuleLog.WriteLog "LOI", "TAI POSTING [" & stage & "]: " & Err.Description

If showMessage Then MsgBox "TAI POSTING loi tai " & stage & ":" & vbCrLf & Err.Description & vbCrLf & vbCrLf & "POSTING cu duoc giu nguyen.", vbCritical, "TAI POSTING"

End Function

Private Function POST_ExposeResultRows(ByVal ses As Object, ByVal grid As Object, _
    ByRef expectedRows As Long, ByRef detail As String) As Boolean

    Dim maxRows As Long, observedRows As Long, nextRow As Long, pageRows As Long

    Dim stableAtEnd As Long, timeoutSec As Long, maxRowsAllowed As Long, started As Double

    If Not POST_GetGridRowCount(grid, maxRows, detail) Then Exit Function

    maxRowsAllowed = ModuleHelper.ToolConfigLong("POST_MAX_EXPECTED_ROWS", 20000, 1, 50000)

    timeoutSec = ModuleHelper.ToolConfigLong("POST_CAPTURE_TIMEOUT_SEC", 45, 10, 300)

    If maxRows > maxRowsAllowed Then

        detail = "SAP tra " & CStr(maxRows) & " dong, vuot gioi han an toan " & _
            CStr(maxRowsAllowed) & ". SAP: " & POST_SAPReplyText(ses)

        Exit Function

    End If

    ' A zero-row grid can be a valid complete result.  Wait once for SAP to

    ' settle before accepting it; no scrolling is possible in that case.

    If maxRows = 0 Then

        If Not SAPX_WaitIdle(ses, 10) Then

            detail = "SAP van dang xu ly ket qua 0 dong. SAP: " & POST_SAPReplyText(ses)

            Exit Function

        End If

        SAPX_Sleep 150

        If Not POST_GetGridRowCount(grid, observedRows, detail) Then Exit Function

        If observedRows = 0 Then

            expectedRows = 0

            POST_ExposeResultRows = True

            Exit Function

        End If

        maxRows = observedRows

    End If

    pageRows = POST_VisibleGridRows(grid)

    nextRow = 0

    started = Timer

    Do

        If SAPX_Elapsed(started) > timeoutSec Then

            detail = "SAP chua hien du " & CStr(maxRows) & " dong trong " & _
                CStr(timeoutSec) & " giay. SAP: " & POST_SAPReplyText(ses)

            Exit Function

        End If

        If Not POST_PullGridToRow(ses, grid, nextRow, detail) Then Exit Function

        If Not POST_GetGridRowCount(grid, observedRows, detail) Then Exit Function

        If observedRows < maxRows Then

            detail = "So dong SAP thay doi trong khi tai: truoc=" & CStr(maxRows) & _
                "; sau=" & CStr(observedRows) & ". SAP: " & POST_SAPReplyText(ses)

            Exit Function

        End If

        If observedRows > maxRows Then

            If observedRows > maxRowsAllowed Then

                detail = "SAP tra " & CStr(observedRows) & " dong, vuot gioi han an toan " & _
                    CStr(maxRowsAllowed) & ". SAP: " & POST_SAPReplyText(ses)

                Exit Function

            End If

            maxRows = observedRows

            stableAtEnd = 0

        End If

        If nextRow + pageRows >= maxRows Then

            stableAtEnd = stableAtEnd + 1

            If stableAtEnd >= 2 Then Exit Do

            nextRow = maxRows - 1

        Else

            nextRow = nextRow + pageRows

        End If

        pageRows = POST_VisibleGridRows(grid)

    Loop

    expectedRows = maxRows

    POST_ExposeResultRows = True

End Function

Private Function POST_ReadResultRows(ByVal ses As Object, ByVal grid As Object, _
    ByVal products As Object, ByVal cProd As String, ByVal cQty As String, _
    ByVal cBatch As String, ByVal cBin As String, ByVal cTst As String, _
    ByVal colDate As String, ByVal cTime As String, ByVal cTotal As String, _
    ByVal cDestBin As String, ByVal cQtyPost As String, ByVal cDestBatch As String, _
    ByVal cST As String, ByVal cProcType As String, ByVal cAvlBU As String, _
    ByVal cUom As String, ByVal expectedUom As String, _
    ByVal expectedRows As Long, ByRef rows As Collection, ByRef detail As String, _
    Optional ByVal cOwner As String = "", Optional ByVal expectedOwner As String = "") As Boolean

    Dim r As Long, finalRows As Long, nextPageAt As Long, pageRows As Long

    Dim lastP As String, pv As String, qv As String, bv As String, tv As String

    Dim binv As String, datev As String, timev As String, totalv As String

    Dim destBinv As String, qtyPostv As String, destBatchv As String, stv As String

    Dim procTypev As String, avlBUv As String, oneRow As Variant

    Dim uomValue As String, uomCarry As String

    Set rows = New Collection

    pageRows = POST_VisibleGridRows(grid)

    nextPageAt = 0

    Application.StatusBar = "POSTING: dang doc du " & CStr(expectedRows) & " dong ton..."

    For r = 0 To expectedRows - 1

        If r >= nextPageAt Then

            If Not POST_PullGridToRow(ses, grid, r, detail) Then Exit Function

            pageRows = POST_VisibleGridRows(grid)

            nextPageAt = r + pageRows

        End If

        If Not POST_TryGridText(grid, r, cProd, pv) Or _
           Not POST_TryGridText(grid, r, cQty, qv) Or _
           Not POST_TryGridText(grid, r, cBatch, bv) Or _
           Not POST_TryGridText(grid, r, cBin, binv) Or _
           Not POST_TryGridText(grid, r, cTst, tv) Or _
           Not POST_TryGridText(grid, r, colDate, datev) Or _
           Not POST_TryGridText(grid, r, cTime, timev) Or _
           Not POST_TryGridText(grid, r, cTotal, totalv) Or _
           Not POST_TryGridText(grid, r, cDestBin, destBinv) Or _
           Not POST_TryGridText(grid, r, cQtyPost, qtyPostv) Or _
           Not POST_TryGridText(grid, r, cDestBatch, destBatchv) Or _
           Not POST_TryGridText(grid, r, cST, stv) Or _
           Not POST_TryGridText(grid, r, cProcType, procTypev) Or _
           Not POST_TryGridText(grid, r, cAvlBU, avlBUv) Or _
           Not POST_TryGridText(grid, r, cUom, uomValue) Then

            detail = "Khong doc duoc dong SAP thu " & CStr(r + 1) & " sau khi da keo tai. " & _
                "SAP: " & POST_SAPReplyText(ses)

            Exit Function

        End If

        If pv = "" And qv = "" And bv = "" And binv = "" And tv = "" And _
           datev = "" And timev = "" And stv = "" Then

            detail = "SAP chua hien noi dung dong " & CStr(r + 1) & " sau khi da keo tai. " & _
                "SAP: " & POST_SAPReplyText(ses)

            Exit Function

        End If

        If pv <> "" Then

            If pv <> lastP Then uomCarry = ""

            lastP = pv

        Else

            pv = lastP

        End If

        uomValue = UCase$(Trim$(uomValue))

        If uomValue <> "" Then uomCarry = uomValue Else uomValue = uomCarry

        If products.exists(Trim$(pv)) Then
            If Not ModuleMain.PostingRowAllowed(Trim$(pv), Trim$(bv)) Then GoTo SkipWrongOwnerRow

            If Not ModuleHelper.IsCompatibleUoM(uomValue, expectedUom) Then

                detail = "UoM POSTING: Material=" & pv & "; row=" & CStr(r + 1) & _
                    "; actual=[" & uomValue & "]; expected=" & expectedUom

                Exit Function

            End If

            Dim rowOwner As String

            If cOwner <> "" Then

                If POST_TryGridText(grid, r, cOwner, rowOwner) Then

                    rowOwner = Trim$(rowOwner)
                    If mDirectPosting And rowOwner = "" Then detail = "Direct search has blank Owner; use Advanced layout": Exit Function

                    If rowOwner <> "" And expectedOwner <> "" Then

                        If StrComp(rowOwner, expectedOwner, vbTextCompare) <> 0 Then GoTo SkipWrongOwnerRow

                    End If

                End If

            End If

            If POST_ParseIntegerQty(qv) > 0 And Trim$(bv) <> "" Then

                oneRow = Array(Trim$(pv), POST_ParseIntegerQty(qv), bv, binv, tv, _
                    IIf(cTotal <> "", POST_ParseIntegerQty(totalv), POST_ParseIntegerQty(qv)), _
                    destBinv, POST_ParseIntegerQty(qtyPostv), destBatchv, stv, _
                    datev, timev, procTypev, _
                    IIf(cAvlBU <> "", POST_ParseIntegerQty(avlBUv), POST_ParseIntegerQty(qv)), "", "", 0)

                rows.Add oneRow

            End If

SkipWrongOwnerRow:

        End If

    Next r

    If Not POST_GetGridRowCount(grid, finalRows, detail) Then Exit Function

    If finalRows <> expectedRows Then

        detail = "So dong SAP thay doi trong khi doc: truoc=" & CStr(expectedRows) & _
            "; sau=" & CStr(finalRows) & ". SAP: " & POST_SAPReplyText(ses)

        Exit Function

    End If

    POST_ReadResultRows = True

End Function

Private Function POST_PullGridToRow(ByVal ses As Object, ByVal grid As Object, _
    ByVal rowNo As Long, ByRef detail As String) As Boolean

    Dim firstRow As Long, waitSec As Long, settleMs As Long

    If grid Is Nothing Then

        detail = "Khong con result grid SAP de tai them. SAP: " & POST_SAPReplyText(ses)

        Exit Function

    End If

    waitSec = ModuleHelper.ToolConfigLong("POST_CAPTURE_PAGE_WAIT_SEC", 10, 1, 60)

    settleMs = ModuleHelper.ToolConfigLong("POST_SCROLL_SETTLE_MS", 100, 10, 1000)

    If Not SAPX_WaitIdle(ses, waitSec) Then

        detail = "SAP van dang xu ly truoc khi keo tai dong " & CStr(rowNo + 1) & _
            ". SAP: " & POST_SAPReplyText(ses)

        Exit Function

    End If

    On Error Resume Next

    Err.Clear

    firstRow = CLng(grid.FirstVisibleRow)

    If Err.Number <> 0 Then Err.Clear: firstRow = -1

    If firstRow <> rowNo Then grid.FirstVisibleRow = rowNo

    If Err.Number <> 0 Then

        detail = "Khong the keo SAP result toi dong " & CStr(rowNo + 1) & _
            ". SAP: " & POST_SAPReplyText(ses)

        Err.Clear

        On Error GoTo 0

        Exit Function

    End If

    On Error GoTo 0

    SAPX_Sleep settleMs

    DoEvents

    If Not SAPX_WaitIdle(ses, waitSec) Then

        detail = "SAP chua tai xong dong " & CStr(rowNo + 1) & _
            " sau khi keo. SAP: " & POST_SAPReplyText(ses)

        Exit Function

    End If

    POST_PullGridToRow = True

End Function

Private Function POST_GetGridRowCount(ByVal grid As Object, ByRef rowCount As Long, _
    ByRef detail As String) As Boolean

    On Error GoTo Failed

    If grid Is Nothing Then Err.Raise vbObjectError + 2964, , "SAP result grid is missing."

    rowCount = CLng(grid.rowCount)

    If rowCount < 0 Then Err.Raise vbObjectError + 2965, , "SAP result row count is negative."

    POST_GetGridRowCount = True

    Exit Function

Failed:

    detail = "Khong doc duoc so dong result SAP: " & Err.Description

End Function

Private Function POST_VisibleGridRows(ByVal grid As Object) As Long

    On Error Resume Next

    POST_VisibleGridRows = CLng(grid.VisibleRowCount)

    On Error GoTo 0

    If POST_VisibleGridRows < 1 Then POST_VisibleGridRows = 20

End Function

Private Function POST_TryGridText(ByVal grid As Object, ByVal rowNo As Long, _
    ByVal colId As String, ByRef valueText As String) As Boolean

    Dim v As Variant

    valueText = ""

    If colId = "" Then

        POST_TryGridText = True

        Exit Function

    End If

    On Error GoTo Failed

    v = grid.GetCellValue(rowNo, colId)

    If Not IsError(v) And Not IsNull(v) And Not IsEmpty(v) Then valueText = Trim$(CStr(v))

    POST_TryGridText = True

    Exit Function

Failed:

    valueText = ""

End Function

Private Function POST_SAPReplyText(ByVal ses As Object) As String

    POST_SAPReplyText = SAPX_StatusText(ses)

    If POST_SAPReplyText = "" Then POST_SAPReplyText = "(SAP khong tra status text)"

End Function

Private Function POST_LastDataRow(ByVal ws As Worksheet, Optional ByVal maxCol As Long = 17) As Long

Dim c As Long, candidate As Long, lastRow As Long

lastRow = 1

For c = 1 To maxCol

    candidate = ws.Cells(ws.rows.count, c).End(xlUp).row

    If candidate > lastRow Then lastRow = candidate

Next c

POST_LastDataRow = lastRow

End Function

Private Function POST_ParseIntegerQty(ByVal rawValue As Variant) As Double

Dim parsedQty As Double

If ModuleHelper.ParseSAPIntegerQty(rawValue, parsedQty) Then POST_ParseIntegerQty = parsedQty

End Function

Private Function POST_OpenProductResult(ByVal ses As Object, ByVal product As String, ByRef grid As Object, ByRef msg As String) As Boolean
Const base As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/"
Dim ctl As Object, ownerCtl As Object, toggle As Object, ownerValue As String
Dim ids As Collection, ni As Object, nc As Object, ownerCol As String, stage As String
On Error GoTo Fail
stage = "Navigate"
SAPX_Navigate ses, ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST")
Set ctl = SAPX_FindById(ses, base & "subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/ctxt/SCWM/S_ASPQ_STOCK-MATNR")
If ctl Is Nothing Then
 Set toggle = SAPX_FindById(ses, base & "btnGV_BUTTON_TEXT")
 If toggle Is Nothing Then Err.Raise 5, , "Thieu o MATNR thuong va nut dong Advanced"
 toggle.Press
 If Not SAPX_WaitIdle(ses, 15) Then Err.Raise 5, , "SAP Busy"
 Set ctl = SAPX_FindById(ses, base & "subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/ctxt/SCWM/S_ASPQ_STOCK-MATNR")
End If
If ctl Is Nothing Then Err.Raise 5, , "Khong tim thay MATNR theo mb12.vbs"
' Direct Product search; Owner must be verifiable in every returned row.
stage = "MATNR readback"
SAPX_SetValue ctl, product
If SAPX_ControlValue(ctl) <> product Then Err.Raise 5, , "MATNR readback khong khop"
stage = "CMD_GO"
ses.findById(base & "subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/btnCMD_GO").Press
If Not SAPX_WaitIdle(ses, 60) Then Err.Raise 5, , "SAP Busy khi tim Material"
If SAPX_StatusType(ses) = "E" Or SAPX_StatusType(ses) = "A" Then Err.Raise 5, , SAPX_StatusText(ses)
Set grid = SAPX_FindById(ses, base & "subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell")
If grid Is Nothing Then Err.Raise 5, , "Thieu stock grid theo mb12.vbs"
SAPX_BuildColumnCatalogFast grid, ids, ni, nc
ownerCol = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("POST_COL_OWNER", ""), "OWNER", False, ids, ni, nc)
If ownerCol = "" Then Err.Raise 5, , "Direct search has no Owner column"
Dim ownerRow As Long, rowOwnerValue As String
For ownerRow = 0 To CLng(grid.rowCount) - 1
 SAPX_EnsureGridRowVisible grid, ownerRow
 rowOwnerValue = RequiredCell(grid, ownerRow, ownerCol)
 If rowOwnerValue = "" Then Err.Raise 5, , "Direct search has blank Owner values"
Next ownerRow
ModuleLog.WriteLog "POST_SEARCH_ROUTE", "DIRECT_PRODUCT; Owner checked from result rows"
mDirectPosting = True
POST_OpenProductResult = True
Exit Function
Fail:
msg = "POSTING thuong / " & stage & ": " & Err.Description
End Function

Private Function POST_FindProductComboRec(ByVal obj As Object, ByVal depth As Long) As Object

Dim typ As String, txt As String, i As Long, ch As Object, f As Object

If obj Is Nothing Or depth > 16 Then Exit Function

On Error Resume Next: typ = UCase$(CStr(obj.Type)): txt = UCase$(POST_ObjectText(obj)): On Error GoTo 0

If InStr(typ, "COMBOBOX") > 0 And InStr(txt, "PRODUCT") > 0 Then Set POST_FindProductComboRec = obj: Exit Function

On Error Resume Next

For i = 0 To obj.children.count - 1

    Set ch = SAPX_CollectionObjectAt(obj.children, i): Set f = POST_FindProductComboRec(ch, depth + 1)

    If Not f Is Nothing Then Set POST_FindProductComboRec = f: Exit Function

Next i

On Error GoTo 0

End Function

Private Sub POST_ScoreProductInputRec(ByVal obj As Object, ByVal combo As Object, ByRef best As Object, ByRef bestScore As Double, ByVal depth As Long)

Dim typ As String, idt As String, l As Double, t As Double, cr As Double, ct As Double, dx As Double, dy As Double, score As Double, i As Long, ch As Object

If obj Is Nothing Or combo Is Nothing Or depth > 16 Then Exit Sub

On Error Resume Next: typ = UCase$(CStr(obj.Type)): idt = UCase$(CStr(obj.id)): On Error GoTo 0

If InStr(typ, "TEXTFIELD") > 0 And InStr(idt, "OKCD") = 0 Then

    On Error Resume Next: l = CDbl(obj.ScreenLeft): t = CDbl(obj.ScreenTop): cr = CDbl(combo.ScreenLeft) + CDbl(combo.width): ct = CDbl(combo.ScreenTop): On Error GoTo 0

    dx = l - cr: dy = Abs(t - ct)

    If dx >= -5 And dx <= 600 And dy <= 25 Then

        score = 100000# - dx - dy * 20#

        If InStr(idt, "PRODUCT") > 0 Or InStr(idt, "MATNR") > 0 Or InStr(idt, "MATID") > 0 Then score = score + 10000#

        If score > bestScore Then bestScore = score: Set best = obj

    End If

End If

On Error Resume Next

For i = 0 To obj.children.count - 1: Set ch = SAPX_CollectionObjectAt(obj.children, i): POST_ScoreProductInputRec ch, combo, best, bestScore, depth + 1: Next i

On Error GoTo 0

End Sub

Private Sub POST_FindProductControls(ByVal ses As Object, ByRef crit As Object, ByRef productCtl As Object)

Dim root As Object, score As Double

Set crit = SAPX_FindById(ses, ModuleHelper.ToolConfig("POST_CRITERION_ID", ""))

Set productCtl = SAPX_FindById(ses, ModuleHelper.ToolConfig("POST_PRODUCT_ID", ""))

On Error Resume Next: Set root = ses.findById("wnd[0]", False): On Error GoTo 0

If crit Is Nothing Then Set crit = POST_FindProductComboRec(root, 0)

If productCtl Is Nothing And Not crit Is Nothing Then POST_ScoreProductInputRec root, crit, productCtl, score, 0

If productCtl Is Nothing Then Set productCtl = SAPX_FindByTokens(ses, "ASPQ_STOCK-MATNR|MATNR|MATID|PRODUCT", True)

End Sub

Private Function POST_EnsureProductCriterion(ByVal ctl As Object) As Boolean

Dim txt As String

If ctl Is Nothing Then POST_EnsureProductCriterion = True: Exit Function

txt = UCase$(POST_ObjectText(ctl))

If InStr(txt, "PRODUCT") > 0 Then POST_EnsureProductCriterion = True: Exit Function

POST_SetCriterion ctl

txt = UCase$(POST_ObjectText(ctl))

POST_EnsureProductCriterion = (InStr(txt, "PRODUCT") > 0 Or UCase$(Trim$(SAPX_ControlValue(ctl))) = "PROD")

End Function

Private Sub POST_SetCriterion(ByVal ctl As Object)

On Error Resume Next

ctl.key = "PROD"

If Err.Number <> 0 Then Err.Clear: ctl.text = "Product"

If Err.Number <> 0 Then Err.Clear: ctl.key = "1"

On Error GoTo 0

End Sub

Private Function POST_OpenAdvancedProductResult(ByVal ses As Object, ByVal products As Object, ByRef grid As Object, ByRef msg As String) As Boolean

Const ID_ADV_TOGGLE As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/btnGV_BUTTON_TEXT"

Const ID_ADV_RESET As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_AS_BUTTONS:/SCMB/SAPLSERVICES:1000/btnCMD_CLEAR_ADVANCED"

Const ID_ADV_START As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_AS_BUTTONS:/SCMB/SAPLSERVICES:1000/btnCMD_START_ADVANCED"

Const ID_PRODUCT_MULTI As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/btn%_P_MATNR_%_APP_%-VALU_PUSH"

Const ID_OWNER_LOW As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/ctxtP_OWNER-LOW"

Const ID_STOCK_GRID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell"

Dim ctl As Object, pop As Object, ws As Worksheet, c As Long, lr As Long

Dim ownerValue As String, statusText As String, statusType As String, batchFilters As Object

On Error GoTo Fail

msg = "": Set grid = Nothing

On Error Resume Next

Set pop = ses.findById("wnd[1]", False)

If Not pop Is Nothing Then

    Set ctl = ses.findById("wnd[1]/tbar[0]/btn[12]", False)

    If Not ctl Is Nothing Then ctl.Press Else pop.Close

End If

On Error GoTo Fail

SAPX_Navigate ses, ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST")

Set ctl = SAPX_FindById(ses, ID_OWNER_LOW)

If ctl Is Nothing Then

    Set ctl = SAPX_FindById(ses, ID_ADV_TOGGLE)

    If ctl Is Nothing Then Err.Raise vbObjectError + 2950, , "Khong tim thay nut Open Advanced Search."

    ctl.Press

    If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 2951, , "SAP Busy khi mo Advanced Search."

End If

Set ctl = SAPX_FindById(ses, ID_ADV_RESET)

If ctl Is Nothing Then Err.Raise vbObjectError + 2952, , "Thieu nut xoa Advanced Search; khong the dam bao Batch trong."

ctl.Press

If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 2952, , "SAP Busy khi reset Advanced Search."

ownerValue = Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R8").Value2))

If ownerValue = "" Then Err.Raise vbObjectError + 2952, , "Owner bat buoc."

Set ctl = SAPX_FindById(ses, ID_OWNER_LOW)

If ctl Is Nothing Then Err.Raise vbObjectError + 2952, , "Khong tim thay Owner trong Advanced Search sau khi mo lai."

SAPX_SetValue ctl, ownerValue

If StrComp(SAPX_ControlValue(ctl), ownerValue, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 2953, , "Owner readback khong khop " & ownerValue

If products.count = 1 Then
 Dim oneProduct As Variant
 oneProduct = products.keys
 Set ctl = SAPX_FindById(ses, Replace(ID_PRODUCT_MULTI, "btn%_P_MATNR_%_APP_%-VALU_PUSH", "ctxtP_MATNR-LOW"))
 If ctl Is Nothing Then Err.Raise vbObjectError + 2954, , "Missing single Product field in Advanced Search"
 SAPX_SetValue ctl, CStr(oneProduct(0))
 If SAPX_ControlValue(ctl) <> CStr(oneProduct(0)) Then Err.Raise vbObjectError + 2954, , "Product readback mismatch"
 ModuleLog.WriteLog "POST_SEARCH_ROUTE", "ADVANCED_SINGLE; Owner verified; no multiple selection popup"
Else
Set ctl = SAPX_FindById(ses, ID_PRODUCT_MULTI)

If ctl Is Nothing Then Err.Raise vbObjectError + 2954, , "Khong tim thay Multiple Selection cua Product."

ctl.Press

If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 2955, , "SAP Busy khi mo Product list."

Set pop = SAPX_FindById(ses, "wnd[1]")

If pop Is Nothing Then Err.Raise vbObjectError + 2956, , "Khong mo duoc popup Product list."

On Error Resume Next

Set ctl = ses.findById("wnd[1]/tbar[0]/btn[16]", False)

If Not ctl Is Nothing Then ctl.Press

On Error GoTo Fail

Set ws = ThisWorkbook.Worksheets("SEARCH_LIST")

lr = ws.Cells(ws.rows.count, 1).End(xlUp).row

If lr < 2 Then Err.Raise vbObjectError + 2957, , "INPUT khong co Material."

ws.Range("A2:A" & CStr(lr)).Copy

On Error Resume Next

Set ctl = ses.findById("wnd[1]/tbar[0]/btn[24]", False)

On Error GoTo Fail

If ctl Is Nothing Then Err.Raise vbObjectError + 2958, , "Popup Product khong co nut Paste from Clipboard."

ctl.Press

SAPX_WaitIdle ses, 10

Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[8]")

If ctl Is Nothing Then Err.Raise vbObjectError + 2959, , "Popup Product khong co nut OK."

ctl.Press

If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 2960, , "SAP Busy khi nhan list Product."

ModuleLog.WriteLog "POST_SEARCH_ROUTE", "ADVANCED_MULTI; materials=" & products.count
End If

Set batchFilters = ModuleMain.PostingBatches(products)
POST_ApplyBatchFilter ses, batchFilters

Set ctl = SAPX_FindById(ses, ID_ADV_START)

If ctl Is Nothing Then Err.Raise vbObjectError + 2961, , "Khong tim thay nut Advanced Search."

ctl.Press

If Not SAPX_WaitIdle(ses, 60) Then Err.Raise vbObjectError + 2962, , "SAP Busy qua lau khi tai " & products.count & " Product."

statusText = SAPX_StatusText(ses): statusType = SAPX_StatusType(ses)

If statusType = "E" Or statusType = "A" Then Err.Raise vbObjectError + 2963, , "SAP: " & statusText

Set grid = SAPX_FindById(ses, ID_STOCK_GRID)

If grid Is Nothing Then Set grid = SAPX_WaitGrid(ses, 25)

If grid Is Nothing Then Err.Raise vbObjectError + 2964, , "Khong tim thay stock grid sau Advanced Search."

POST_OpenAdvancedProductResult = True

Application.CutCopyMode = False

Exit Function

Fail:

msg = "Advanced Search: " & Err.Description

Application.CutCopyMode = False

End Function

Private Function POST_ObjectText(ByVal obj As Object) As String

On Error Resume Next

POST_ObjectText = Trim$(CStr(obj.text))

If POST_ObjectText = "" Then POST_ObjectText = Trim$(CStr(obj.value))

If POST_ObjectText = "" Then POST_ObjectText = Trim$(CStr(obj.key))

On Error GoTo 0

End Function

Private Function SAPX_StatusText(ByVal ses As Object) As String

On Error Resume Next: SAPX_StatusText = Trim$(CStr(ses.findById("wnd[0]/sbar", False).text)): On Error GoTo 0

End Function

Private Function SAPX_StatusType(ByVal ses As Object) As String

On Error Resume Next: SAPX_StatusType = UCase$(Trim$(CStr(ses.findById("wnd[0]/sbar", False).messageType))): On Error GoTo 0

End Function

' RC20-DEF-08: Status bar messageId and messageNumber

Private Function SAPX_MessageId(ByVal ses As Object) As String

On Error Resume Next: SAPX_MessageId = UCase$(Trim$(CStr(ses.findById("wnd[0]/sbar", False).messageId))): On Error GoTo 0

End Function

Private Function SAPX_MessageNumber(ByVal ses As Object) As String

On Error Resume Next: SAPX_MessageNumber = Trim$(CStr(ses.findById("wnd[0]/sbar", False).messageNumber)): On Error GoTo 0

End Function

Private Sub SAPX_DumpControlIds(ByVal obj As Object, ByRef dumpText As String, ByVal depth As Long)

Dim i As Long, ch As Object, idText As String, typeText As String, labelText As String

If obj Is Nothing Or depth > 14 Or Len(dumpText) > 30000 Then Exit Sub

On Error Resume Next

idText = CStr(obj.id): typeText = CStr(obj.Type): labelText = SAPX_ControlText(obj)

On Error GoTo 0

If InStr(1, UCase$(idText & " " & labelText), "DATE", vbTextCompare) > 0 Or _
   InStr(1, UCase$(idText & " " & labelText), "WADAT", vbTextCompare) > 0 Or _
   InStr(1, UCase$(idText & " " & labelText), "GI", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "TEXTFIELD", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "BUTTON", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "RADIO", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "LABEL", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "SHELL", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "GRID", vbTextCompare) > 0 Then

    dumpText = dumpText & idText & " | " & typeText & " | " & labelText & vbLf

End If

On Error Resume Next

For i = 0 To obj.children.count - 1

    Set ch = SAPX_CollectionObjectAt(obj.children, i)

    SAPX_DumpControlIds ch, dumpText, depth + 1

Next i

Err.Clear

On Error GoTo 0

End Sub

Private Sub VL6_OpenSelection(ByVal ses As Object)

Dim ctl As Object, btn As Object, currentUser As String

If Not SAPX_TcodeIs(ses, "VL06O") Then SAPX_Navigate ses, "VL06O"
On Error Resume Next: currentUser = Trim$(UCase$(CStr(ses.info.User))): On Error GoTo 0
If currentUser <> "" Then ModuleHelper.PutConfig "SAP_CURRENT_USER", currentUser: ModuleLog.WriteLog "VL06O_USER", "Current SAP User=" & currentUser

Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_FROM_ID", "wnd[0]/usr/ctxtIT_WADAT-LOW"))

If Not ctl Is Nothing Then Exit Sub

Set btn = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_LIST_OUTBOUND_BUTTON_ID", "wnd[0]/usr/btnBUTTON6"))

If Not btn Is Nothing Then btn.Press: SAPX_WaitIdle ses, 15

Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_FROM_ID", "wnd[0]/usr/ctxtIT_WADAT-LOW"))

If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "WADAT-LOW|WADAT_LOW|WADATLOW", True)

If Not ctl Is Nothing Then Exit Sub

Set btn = SAPX_FindButtonText(ses.findById("wnd[0]/usr", False), "LIST OUTBOUND DELIVERIES", 0)

If Not btn Is Nothing Then btn.Press: SAPX_WaitIdle ses, 15

Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_FROM_ID", "wnd[0]/usr/ctxtIT_WADAT-LOW"))

If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "WADAT-LOW|WADAT_LOW|WADATLOW", True)

End Sub

Private Function VL6_SetSelectionShipTo(ByVal ses As Object, ByVal shipValue As String) As Boolean

Dim ctl As Object, hi As Object, configuredId As String

configuredId = ModuleHelper.ToolConfig("VL06O_SHIP_TO_ID", "wnd[0]/usr/ctxtIT_KUNWE-LOW")

Set ctl = SAPX_FindById(ses, configuredId)

If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "IT_KUNWE-LOW|IT_KUNNR-LOW|S_KUNWE-LOW|S_KUNNR-LOW|KUNWE-LOW|KUNNR-LOW|SHIP-TO PARTY|SHIP TO PARTY", True)

If ctl Is Nothing Then Exit Function

SAPX_SetValue ctl, shipValue

Set hi = SAPX_FindById(ses, Replace(CStr(ctl.id), "-LOW", "-HIGH"))

If Not hi Is Nothing Then hi.text = ""

On Error Resume Next

VL6_SetSelectionShipTo = (StrComp(Trim$(CStr(ctl.text)), Trim$(shipValue), vbTextCompare) = 0)

On Error GoTo 0

If Not VL6_SetSelectionShipTo Then Err.Raise vbObjectError + 2942, , "Khong dien duoc Ship-to=" & shipValue & " tai man hinh chon VL06O."

End Function

Private Sub VL6_PressItemView(ByVal ses As Object)

Dim btn As Object

Set btn = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_ITEM_VIEW_BUTTON_ID", "wnd[0]/tbar[1]/btn[18]"))

If btn Is Nothing Then Err.Raise vbObjectError + 4202, , "Thieu nut Item View theo VBS."

btn.Press

If Not SAPX_WaitIdle(ses, 20) Then Err.Raise vbObjectError + 4202, , "Item View timeout."

End Sub

Private Function VL6_DirectGrid(ByVal ses As Object) As Object

Set VL6_DirectGrid = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_GRID_ID", "wnd[0]/usr/cntlGRID1/shellcont/shell"))

End Function

Private Function VL6_ItemGridReady(ByVal grid As Object) As Boolean

If grid Is Nothing Then Exit Function

On Error GoTo NotReady

If Not SAPX_ColumnExists(grid, "MATNR") Then Exit Function

If Not SAPX_ColumnExists(grid, ModuleHelper.ToolConfig("VL06O_COL_ITEM", "POSNR")) Then Exit Function

VL6_ItemGridReady = True

NotReady:

End Function

Private Function VL6_EnsureItemGrid(ByVal ses As Object) As Object

Dim grid As Object, started As Double, busyNow As Boolean, switchStarted As Double

switchStarted = Timer

Set grid = VL6_DirectGrid(ses)

If Not VL6_ItemGridReady(grid) Then VL6_PressItemView ses

started = Timer

Do

 busyNow = True

 On Error Resume Next: busyNow = CBool(ses.Busy): On Error GoTo 0

 If Not busyNow Then

  Set grid = VL6_DirectGrid(ses)

  If VL6_ItemGridReady(grid) Then

   Set VL6_EnsureItemGrid = grid

   ModuleLog.WriteLog "VL06O_ITEM_VIEW", "Ready=" & Format$(SAPX_Elapsed(switchStarted), "0.00") & "s"

   Exit Function

  End If

 End If

 If SAPX_Elapsed(started) >= 20 Then Exit Do

 DoEvents: SAPX_Sleep 25

Loop

Err.Raise vbObjectError + 4202, , "Khong co Item View grid tai duong dan da ghi."

End Function

Private Function SAPX_TcodeIs(ByVal ses As Object, ByVal tcode As String) As Boolean

Dim a As String, b As String

On Error Resume Next: a = UCase$(Trim$(CStr(ses.info.Transaction))): On Error GoTo 0

b = UCase$(Trim$(tcode))

If Left$(a, 1) = "/" Then a = mid$(a, 2)

If Left$(b, 1) = "/" Then b = mid$(b, 2)

SAPX_TcodeIs = (a = b)

End Function

Private Sub SAPX_Navigate(ByVal ses As Object, ByVal tcode As String)

Dim cmd As String

cmd = tcode: If Left$(cmd, 1) = "/" Then cmd = "/n" & cmd Else cmd = "/n" & cmd

ses.findById("wnd[0]/tbar[0]/okcd").text = cmd: ses.findById("wnd[0]").sendVKey 0

If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 2940, , "SAP Busy khi mo " & tcode

End Sub

Private Function SAPX_WaitIdle(ByVal ses As Object, ByVal timeoutSec As Double) As Boolean

Dim t As Double, busyNow As Boolean, eno As Long

t = Timer

Do

 On Error Resume Next

 Err.Clear

 busyNow = ses.Busy

 eno = Err.Number

 On Error GoTo 0

 If eno <> 0 Then Err.Raise vbObjectError + 2940, , "Mat ket noi SAP khi cho."

 If Not busyNow Then

  If mMeasureWait Then mWaitSeconds = mWaitSeconds + SAPX_Elapsed(t)

  SAPX_WaitIdle = True: Exit Function

 End If

 DoEvents: SAPX_Sleep 25

Loop While SAPX_Elapsed(t) < timeoutSec

If mMeasureWait Then mWaitSeconds = mWaitSeconds + SAPX_Elapsed(t)

End Function

Private Function SAPX_Elapsed(ByVal t As Double) As Double

If Timer >= t Then SAPX_Elapsed = Timer - t Else SAPX_Elapsed = 86400# - t + Timer

End Function

Private Function SAPX_FindById(ByVal ses As Object, ByVal idText As String) As Object

If Trim$(idText) = "" Then Exit Function

On Error Resume Next: Set SAPX_FindById = ses.findById(idText, False): On Error GoTo 0

End Function

Private Function SAPX_CollectionObjectAt(ByVal items As Object, ByVal indexNo As Long) As Object

Dim idx As Variant

If items Is Nothing Then Exit Function

idx = CLng(indexNo)

On Error Resume Next

Err.Clear

Set SAPX_CollectionObjectAt = items.ElementAt(idx)

If SAPX_CollectionObjectAt Is Nothing Or Err.Number <> 0 Then

    Err.Clear

    Set SAPX_CollectionObjectAt = items.item(idx)

End If

If SAPX_CollectionObjectAt Is Nothing Or Err.Number <> 0 Then

    Err.Clear

    Set SAPX_CollectionObjectAt = items(CLng(indexNo))

End If

If SAPX_CollectionObjectAt Is Nothing Or Err.Number <> 0 Then

    Err.Clear

    Set SAPX_CollectionObjectAt = items(CInt(indexNo))

End If

Err.Clear

On Error GoTo 0

End Function

Private Function SAPX_CollectionValueAt(ByVal items As Object, ByVal indexNo As Long) As Variant

Dim idx As Variant

If items Is Nothing Then Exit Function

idx = CLng(indexNo)

On Error Resume Next

Err.Clear

SAPX_CollectionValueAt = items.ElementAt(idx)

If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items.item(idx)

If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items(CLng(indexNo))

If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items(CInt(indexNo))

Err.Clear

On Error GoTo 0

End Function

Private Function SAPX_FindByTokens(ByVal ses As Object, ByVal tokens As String, ByVal textOnly As Boolean) As Object

Dim root As Object, found As Object, key As String

key = CStr(ses.id) & "|" & tokens & "|" & CStr(textOnly)

If mControlPaths Is Nothing Then Set mControlPaths = CreateObject("Scripting.Dictionary")

If mControlPaths.exists(key) Then

 Set found = SAPX_FindById(ses, CStr(mControlPaths(key)))

 If Not found Is Nothing Then Set SAPX_FindByTokens = found: Exit Function

 mControlPaths.Remove key

End If

On Error Resume Next: Set root = ses.findById("wnd[0]/usr", False): On Error GoTo 0

Set found = SAPX_FindTokensRec(root, UCase$(tokens), textOnly, 0)

If Not found Is Nothing Then mControlPaths(key) = CStr(found.id)

Set SAPX_FindByTokens = found

End Function

Private Function SAPX_FindTokensRec(ByVal obj As Object, ByVal tokens As String, ByVal textOnly As Boolean, ByVal depth As Long) As Object

Dim idt As String, typ As String, a As Variant, ch As Object, i As Long

If obj Is Nothing Or depth > 12 Then Exit Function

On Error Resume Next: idt = UCase$(CStr(obj.id)): typ = UCase$(CStr(obj.Type)): On Error GoTo 0

If (Not textOnly Or InStr(typ, "TEXTFIELD") > 0 Or InStr(typ, "COMBOBOX") > 0) Then

    For Each a In Split(tokens, "|"): If Trim$(CStr(a)) <> "" And InStr(idt, UCase$(Trim$(CStr(a)))) > 0 Then Set SAPX_FindTokensRec = obj: Exit Function

    Next a

End If

On Error Resume Next

For i = 0 To obj.children.count - 1

    Set ch = SAPX_CollectionObjectAt(obj.children, i)

    Set SAPX_FindTokensRec = SAPX_FindTokensRec(ch, tokens, textOnly, depth + 1)

    If Not SAPX_FindTokensRec Is Nothing Then Exit Function

Next i

On Error GoTo 0

End Function

Private Sub SAPX_SetValue(ByVal ctl As Object, ByVal valueText As String)

Dim typ As String

On Error Resume Next: typ = UCase$(CStr(ctl.Type)): On Error GoTo 0

On Error Resume Next

If InStr(typ, "COMBOBOX") > 0 Then ctl.key = valueText

If Err.Number <> 0 Or InStr(typ, "COMBOBOX") = 0 Then Err.Clear: ctl.text = valueText

If Err.Number <> 0 Then Err.Clear: ctl.key = valueText

On Error GoTo 0

End Sub

Private Function SAPX_ControlValue(ByVal ctl As Object) As String

On Error Resume Next: SAPX_ControlValue = Trim$(CStr(ctl.text)): If SAPX_ControlValue = "" Then SAPX_ControlValue = Trim$(CStr(ctl.key)): On Error GoTo 0

End Function

Private Function SAPX_WaitGrid(ByVal ses As Object, ByVal timeoutSec As Double) As Object

Dim started As Double, grid As Object, rowCount As Long, lastRowCount As Long

Dim stableHits As Long, busyNow As Boolean

started = Timer: lastRowCount = -1

Do

    Set grid = SAPX_FindLargestGrid(ses)

    busyNow = True

    On Error Resume Next

    busyNow = CBool(ses.Busy)

    If Not grid Is Nothing Then rowCount = CLng(grid.rowCount) Else rowCount = -1

    If Err.Number <> 0 Then Err.Clear: rowCount = -1

    On Error GoTo 0

    If Not grid Is Nothing And Not busyNow And rowCount >= 0 Then

        If rowCount = lastRowCount Then

            stableHits = stableHits + 1

        Else

            stableHits = 0

            lastRowCount = rowCount

        End If

        If stableHits >= 1 Then

            Set SAPX_WaitGrid = grid

            Exit Function

        End If

    Else

        stableHits = 0

    End If

    SAPX_Sleep 25

    DoEvents

Loop While SAPX_Elapsed(started) < timeoutSec

If Not grid Is Nothing And Not busyNow And rowCount >= 0 Then Set SAPX_WaitGrid = grid

End Function

Private Function SAPX_FindLargestGrid(ByVal ses As Object) As Object

Dim root As Object, best As Object, bestRows As Long

On Error Resume Next: Set root = ses.findById("wnd[0]", False): On Error GoTo 0

SAPX_FindGridRec root, best, bestRows, 0

Set SAPX_FindLargestGrid = best

End Function

Private Sub SAPX_FindGridRec(ByVal obj As Object, ByRef best As Object, ByRef bestRows As Long, ByVal depth As Long)

Dim typ As String, subTyp As String, n As Long, i As Long, ch As Object, hasRowCount As Boolean

If obj Is Nothing Or depth > 14 Then Exit Sub

On Error Resume Next

typ = UCase$(CStr(obj.Type))

subTyp = UCase$(CStr(obj.SubType))

Err.Clear

n = CLng(obj.rowCount)

hasRowCount = (Err.Number = 0)

Err.Clear

If InStr(typ, "GRIDVIEW") > 0 Or InStr(subTyp, "GRIDVIEW") > 0 Or _
   (InStr(typ, "GUISHELL") > 0 And hasRowCount) Then

    If best Is Nothing Or n > bestRows Then Set best = obj: bestRows = n

End If

For i = 0 To obj.children.count - 1: Set ch = SAPX_CollectionObjectAt(obj.children, i): SAPX_FindGridRec ch, best, bestRows, depth + 1: Next i

On Error GoTo 0

End Sub

Private Sub SAPX_BuildColumnCatalogFast(ByVal grid As Object, ByRef ids As Collection, ByRef normIds As Object, ByRef normCaps As Object)

Dim id As Variant, key As String

Set ids = SAPX_ColumnIds(grid)

Set normIds = CreateObject("Scripting.Dictionary")

Set normCaps = CreateObject("Scripting.Dictionary")

normIds.CompareMode = vbTextCompare

normCaps.CompareMode = vbTextCompare

For Each id In ids

    key = CStr(id)

    normIds(key) = NormalizeHeader(key)

    normCaps(key) = NormalizeHeader(SAPX_ColumnCaption(grid, key))

Next id

End Sub

Private Function SAPX_ResolveColumnFast(ByVal configured As String, ByVal aliasesCsv As String, ByVal required As Boolean, _
ByVal ids As Collection, ByVal normIds As Object, ByVal normCaps As Object) As String

Dim id As Variant, a As Variant, key As String, ni As String, nc As String, na As String

If Not ids Is Nothing Then

    If Trim$(configured) <> "" Then

        For Each id In ids

            If StrComp(CStr(id), configured, vbTextCompare) = 0 Then SAPX_ResolveColumnFast = CStr(id): Exit Function

        Next id

    End If

    ' Pass 1: exact normalized ID/caption.

    For Each id In ids

        key = CStr(id)

        ni = CStr(normIds(key))

        nc = CStr(normCaps(key))

        For Each a In Split(aliasesCsv, ",")

            na = NormalizeHeader(Trim$(CStr(a)))

            If na <> "" And (ni = na Or nc = na) Then SAPX_ResolveColumnFast = key: Exit Function

        Next a

    Next id

    ' Pass 2: fuzzy only for meaningful aliases.

    For Each id In ids

        key = CStr(id)

        ni = CStr(normIds(key))

        nc = CStr(normCaps(key))

        For Each a In Split(aliasesCsv, ",")

            na = NormalizeHeader(Trim$(CStr(a)))

            If Len(na) >= 4 And (InStr(1, ni, na, vbTextCompare) > 0 Or InStr(1, nc, na, vbTextCompare) > 0) Then

                SAPX_ResolveColumnFast = key

                Exit Function

            End If

        Next a

    Next id

End If

If required Then Err.Raise vbObjectError + 2941, , "Thieu cot SAP: " & aliasesCsv

End Function

Private Function SAPX_ResolveColumn(ByVal grid As Object, ByVal configured As String, ByVal aliasesCsv As String, ByVal required As Boolean) As String

Dim ids As Collection, id As Variant, cap As String, a As Variant, nc As String, na As String, ni As String

If Trim$(configured) <> "" Then If SAPX_ColumnExists(grid, configured) Then SAPX_ResolveColumn = configured: Exit Function

Set ids = SAPX_ColumnIds(grid)

' Pass 1: exact normalized ID or caption. This prevents short aliases such

' as ST from matching SOURCE STORAGE BIN.

For Each id In ids

    ni = NormalizeHeader(CStr(id))

    nc = NormalizeHeader(SAPX_ColumnCaption(grid, CStr(id)))

    For Each a In Split(aliasesCsv, ",")

        na = NormalizeHeader(Trim$(CStr(a)))

        If na <> "" And (ni = na Or nc = na) Then SAPX_ResolveColumn = CStr(id): Exit Function

    Next a

Next id

' Pass 2: fuzzy match only for meaningful aliases (4+ normalized chars).

For Each id In ids

    cap = UCase$(CStr(id) & " " & SAPX_ColumnCaption(grid, CStr(id)))

    nc = NormalizeHeader(cap)

    For Each a In Split(aliasesCsv, ",")

        na = NormalizeHeader(Trim$(CStr(a)))

        If Len(na) >= 4 And InStr(1, nc, na, vbTextCompare) > 0 Then SAPX_ResolveColumn = CStr(id): Exit Function

    Next a

Next id

If required Then Err.Raise vbObjectError + 2941, , "Thieu cot SAP: " & aliasesCsv

End Function

Private Function SAPX_ColumnIds(ByVal grid As Object) As Collection

Dim c As New Collection, o As Object, i As Long, v As Variant

On Error Resume Next: Set o = grid.ColumnOrder

If Not o Is Nothing Then

    For i = 0 To o.count - 1

        v = SAPX_CollectionValueAt(o, i)

        If Trim$(CStr(v & "")) <> "" Then c.Add CStr(v)

    Next i

End If

On Error GoTo 0

Set SAPX_ColumnIds = c

End Function

Private Function SAPX_ColumnCaption(ByVal grid As Object, ByVal colId As String) As String

Dim v As Variant, o As Object, i As Long, s As String

On Error Resume Next

v = grid.GetDisplayedColumnTitle(colId): s = CStr(v & "")

If s = "" Then

    v = grid.GetColumnTitles(colId)

    If IsArray(v) Then

        For i = LBound(v) To UBound(v)

            If Trim$(CStr(v(i) & "")) <> "" Then s = s & " " & CStr(v(i))

        Next i

    End If

End If

If s = "" And IsObject(v) Then

    Set o = v

    For i = 0 To o.count - 1

        s = s & " " & CStr(o(i))

    Next i

End If

On Error GoTo 0

SAPX_ColumnCaption = Trim$(s)

End Function

Private Function SAPX_ColumnExists(ByVal grid As Object, ByVal colId As String) As Boolean

Dim ids As Collection, v As Variant

Set ids = SAPX_ColumnIds(grid)

For Each v In ids: If StrComp(CStr(v), colId, vbTextCompare) = 0 Then SAPX_ColumnExists = True: Exit Function

Next v

End Function

Private Function SAPX_GridText(ByVal grid As Object, ByVal rowNo As Long, ByVal colId As String) As String

If colId = "" Then Exit Function

On Error Resume Next: SAPX_GridText = Trim$(CStr(grid.GetCellValue(rowNo, colId) & "")): On Error GoTo 0

End Function

Private Sub SAPX_EnsureGridRowVisible(ByVal grid As Object, ByVal rowNo As Long)

Dim firstRow As Long, visibleRows As Long

' GetCellValue can read rows outside the viewport. Only touch the scrollbar

' periodically so large VL06O lists do not make two COM calls for every row.

If rowNo > 0 And rowNo Mod 20 <> 0 Then Exit Sub

On Error Resume Next

firstRow = CLng(grid.FirstVisibleRow)

visibleRows = CLng(grid.VisibleRowCount)

If visibleRows < 1 Then visibleRows = 1

If rowNo < firstRow Or rowNo >= firstRow + visibleRows Then

    grid.FirstVisibleRow = rowNo

    DoEvents

End If

On Error GoTo 0

End Sub

Private Function SAPX_FindButtonText(ByVal obj As Object, ByVal wanted As String, ByVal depth As Long) As Object

Dim typ As String, txt As String, i As Long, ch As Object

If obj Is Nothing Or depth > 12 Then Exit Function

On Error Resume Next: typ = UCase$(CStr(obj.Type)): txt = UCase$(CStr(obj.text)): On Error GoTo 0

If InStr(typ, "BUTTON") > 0 And InStr(txt, UCase$(wanted)) > 0 Then Set SAPX_FindButtonText = obj: Exit Function

On Error Resume Next

For i = 0 To obj.children.count - 1: Set ch = SAPX_CollectionObjectAt(obj.children, i): Set SAPX_FindButtonText = SAPX_FindButtonText(ch, wanted, depth + 1): If Not SAPX_FindButtonText Is Nothing Then Exit Function

Next i

On Error GoTo 0

End Function

Private Sub SAPX_ClearGridSelection(ByVal grid As Object)

' A saved ALV variant can leave more than one column selected. Clear the

' selection before SelectColumn so the filter popup belongs only to colId.

If grid Is Nothing Then Exit Sub

On Error Resume Next

Err.Clear

grid.clearSelection

Err.Clear

grid.selectedRows = ""

Err.Clear

grid.SelectedColumns = ""

Err.Clear

grid.SelectedCells = ""

Err.Clear

On Error GoTo 0

End Sub

Private Function SAPX_ControlText(ByVal obj As Object) As String

On Error Resume Next: SAPX_ControlText = CStr(obj.text & ""): On Error GoTo 0

End Function

Private Function SAPX_GetSessionForTcode(ByVal tcode As String) As Object

Dim app As Object, con As Object, ses As Object, ci As Long, si As Long, wanted As String

Dim sapUser As String, sessionId As String, eno As Long, detail As String

wanted = Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R7").Value2))

If Not mSession Is Nothing Then

 ' Keep the same session for the complete run; never silently switch after disconnect.

 If Not SAPX_WaitIdle(mSession, 30) Then Err.Raise vbObjectError + 4200, , "SAP session dang Busy qua 30 giay."

 On Error Resume Next

 Err.Clear

 sessionId = CStr(mSession.id)

 sapUser = Trim$(CStr(mSession.info.User))

 eno = Err.Number

 On Error GoTo 0

 If eno <> 0 Then Err.Raise vbObjectError + 4200, , "Phien SAP da mat ket noi. Dung luot chay, khong tu chuyen session."

 If wanted <> "" And UCase$(wanted) <> "AUTO" And wanted <> sessionId Then Err.Raise vbObjectError + 4200, , "Session ID da thay doi trong luot chay."

Else

 On Error Resume Next

 Err.Clear

 Set app = GetObject("SAPGUI").GetScriptingEngine

 eno = Err.Number: detail = Err.Description

 On Error GoTo 0

 If eno <> 0 Or app Is Nothing Then Err.Raise vbObjectError + 4200, , "Khong truy cap duoc SAP GUI Scripting: " & detail

 If app.children.count = 0 Then Err.Raise vbObjectError + 4200, , "SAP GUI chua co connection. Hay mo SAP va dang nhap."

 If wanted = "" Or UCase$(wanted) = "AUTO" Then

  ' BBBG baseline: first connection, first session. Busy is waited, not filtered out.

  Set con = app.children(0)

  If con.children.count = 0 Then Err.Raise vbObjectError + 4200, , "Connection dau chua co session. Hay dang nhap SAP."

  Set mSession = con.children(0)

 Else

  For ci = 0 To app.children.count - 1

   Set con = app.children(CLng(ci))

   For si = 0 To con.children.count - 1

    Set ses = con.children(CLng(si))

    If CStr(ses.id) = wanted Then

     Set mSession = ses

     Exit For

    End If

   Next si

   If Not mSession Is Nothing Then Exit For

  Next ci

  If mSession Is Nothing Then Err.Raise vbObjectError + 4200, , "Khong tim thay Session ID trong INPUT R7: " & wanted

 End If

 If Not SAPX_WaitIdle(mSession, 30) Then Err.Raise vbObjectError + 4200, , "SAP session dang Busy qua 30 giay."

 sapUser = Trim$(CStr(mSession.info.User))

 sessionId = CStr(mSession.id)

 ModuleLog.WriteLog "INFO", "SAP session=" & sessionId & "; System=" & CStr(mSession.info.SystemName) & "; Client=" & CStr(mSession.info.Client), "CONNECT"

End If

If sapUser = "" Then Err.Raise vbObjectError + 4200, , "Phien SAP chua dang nhap. Hay hoan tat dang nhap SAP roi chay lai."

Set SAPX_GetSessionForTcode = mSession

End Function

Private Sub VL6_FindFilterLabels(ByVal obj As Object, ByVal aliases As String, ByRef labels As Collection, ByVal depth As Long)

Dim i As Long, ch As Object, typ As String, txt As String, ident As String

If obj Is Nothing Or depth > 12 Then Exit Sub

On Error Resume Next

typ = CStr(obj.Type): txt = NormalizeHeader(CStr(obj.text)): ident = UCase$(CStr(obj.id))

If (typ = "GuiLabel" Or InStr(ident, "_%-TEXT") > 0) And txt <> "" Then

 If InStr(1, "|" & aliases & "|", "|" & txt & "|", vbTextCompare) > 0 Then labels.Add obj

End If

For i = 0 To obj.children.count - 1

 Set ch = SAPX_CollectionObjectAt(obj.children, i)

 VL6_FindFilterLabels ch, aliases, labels, depth + 1

Next i

On Error GoTo 0

End Sub

Private Function VL6_FilterLabel(ByVal ses As Object, ByVal token As String) As String

Const base As String = "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/"

Dim ctl As Object, prefix As Variant

For Each prefix In Array("txt", "lbl")

 Set ctl = SAPX_FindById(ses, base & CStr(prefix) & "%_" & token & "_%_APP_%-TEXT")

 If Not ctl Is Nothing Then

  VL6_FilterLabel = NormalizeHeader(SAPX_ControlText(ctl))

  If VL6_FilterLabel <> "" Then Exit Function

 End If

Next prefix

End Function

Private Function VL6_FilterField(ByVal ses As Object, ByVal grid As Object, ByVal colId As String) As String

Const base As String = "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/"

Dim labels As Collection, lab As Object, ctl As Object, i As Long, found As String, token As String, aliases As String, matches As Long, labelText As String, recorded As String

aliases = NormalizeHeader(colId) & "|" & NormalizeHeader(SAPX_ColumnCaption(grid, colId))

Select Case UCase$(colId)

 Case "MATNR": aliases = aliases & "|MATERIAL|PRODUCT"

 Case "KUNNR", "KUNWE": aliases = aliases & "|SHIPTO|SHIPTOPARTY|SHIPTOPTY"

 Case "OWNER": aliases = aliases & "|OWNER|STOCKOWNER"

End Select

' First use labels attached to the exact DYN row, regardless of GUI type.

For i = 1 To 99

 token = "%%DYN" & Format$(i, "000")

 Set ctl = SAPX_FindById(ses, base & "ctxt" & token & "-LOW")

 If ctl Is Nothing Then Exit For

 labelText = VL6_FilterLabel(ses, token)

 If labelText <> "" Then

  If InStr(1, "|" & aliases & "|", "|" & labelText & "|", vbTextCompare) > 0 Then matches = matches + 1: found = token

 End If

Next i

If matches = 1 Then VL6_FilterField = found: Exit Function

If matches > 1 Then Err.Raise 5, , "Trung nhan o loc " & colId

' Older GUI label controls: match the left label on the same screen row.

Set labels = New Collection

VL6_FindFilterLabels ses.findById("wnd[1]/usr"), aliases, labels, 0

For i = 1 To 99

 token = "%%DYN" & Format$(i, "000")

 Set ctl = SAPX_FindById(ses, base & "ctxt" & token & "-LOW")

 If ctl Is Nothing Then Exit For

 For Each lab In labels

  If Abs(CDbl(lab.ScreenTop) - CDbl(ctl.ScreenTop)) <= 10 And CDbl(lab.ScreenLeft) < CDbl(ctl.ScreenLeft) Then

   If found <> token Then matches = matches + 1: found = token

  End If

 Next lab

Next i

If matches = 1 Then VL6_FilterField = found: Exit Function

If matches > 1 Then Err.Raise 5, , "Nhieu o loc khop " & colId

' Exact paths recorded by the user for KUNNR and MATNR, after selecting that column.

If UCase$(colId) = "KUNNR" Or UCase$(colId) = "KUNWE" Or UCase$(colId) = "MATNR" Then

 recorded = ModuleHelper.ToolConfig("VL06O_FILTER_" & UCase$(colId) & "_FIELD", "%%DYN001")

 Set ctl = SAPX_FindById(ses, base & "ctxt" & recorded & "-LOW")

 If ctl Is Nothing Then Err.Raise 5, , "Thieu o loc da ghi " & recorded & " / " & colId

 labelText = VL6_FilterLabel(ses, recorded)

 If labelText <> "" Then

  If InStr(1, "|" & aliases & "|", "|" & labelText & "|", vbTextCompare) = 0 Then Err.Raise 5, , "Nhan o loc " & recorded & " khong khop " & colId & ": " & labelText

 End If

 ModuleLog.WriteLog "VL06O_FILTER_FIELD", colId & " recorded path=" & recorded

 VL6_FilterField = recorded: Exit Function

End If

If SAPX_FindById(ses, base & "ctxt%%DYN002-LOW") Is Nothing Then

 If Not SAPX_FindById(ses, base & "ctxt%%DYN001-LOW") Is Nothing Then VL6_FilterField = "%%DYN001": Exit Function

End If

Err.Raise 5, , "Khong xac dinh o loc " & colId & "."

End Function

Private Sub VL6_ClearSelectedColumns(ByVal grid As Object)

Dim columns As Object, value As Variant, names As Collection, i As Long, columnId As Variant

On Error GoTo FullClear

Set columns = grid.SelectedColumns

Set names = New Collection

For i = 0 To columns.count - 1

 value = columns.item(CLng(i)): names.Add CStr(value)

Next i

For Each columnId In names: grid.DeselectColumn CStr(columnId): Next columnId

Exit Sub

FullClear:

On Error GoTo 0

Set names = SAPX_ColumnIds(grid)

For Each columnId In names: grid.DeselectColumn CStr(columnId): Next columnId

End Sub

Private Sub VL6_ApplyALVFilter(ByVal ses As Object, ByVal grid As Object, ByVal colId As String, ByVal scalar As String, Optional ByVal products As Object, Optional ByVal postingGrid As Boolean = False, Optional ByVal clearOnly As Boolean = False)

Const base As String = "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/"

Dim ctl As Object, ws As Worksheet, p As Variant, r As Long, stage As String, detail As String

Dim columnIds As Collection, columnId As Variant, useList As Boolean, keys As Variant, fieldToken As String

On Error GoTo Fail

useList = False

If Not products Is Nothing Then

 If products.count = 0 Then Err.Raise 5, , "Material list rong"

 If products.count = 1 Then

  keys = products.keys: scalar = CStr(keys(0))

 Else

  useList = True

 End If

End If

stage = "SelectColumn " & colId

SAPX_ClearGridSelection grid

VL6_ClearSelectedColumns grid

grid.setCurrentCell -1, colId

grid.SelectColumn colId

stage = "Open filter"

If postingGrid Then

 grid.pressToolbarButton "&MB_FILTER"

Else

 ses.findById("wnd[0]/tbar[1]/btn[29]").Press

End If

If Not SAPX_WaitIdle(ses, 15) Then Err.Raise 5, , "Timeout"

' Clear saved single values and multiple selections before replacing this column.

stage = "Check filter field"

fieldToken = VL6_FilterField(ses, grid, colId)

ModuleLog.WriteLog "VL06O_FILTER_FIELD", colId & "=" & fieldToken

stage = "Clear old values"

Set ctl = ses.findById(base & "ctxt" & fieldToken & "-LOW")

ctl.text = ""

Set ctl = ses.findById(base & "ctxt" & fieldToken & "-HIGH", False)

If Not ctl Is Nothing Then ctl.text = ""

If clearOnly Then
 ses.findById(base & "btn%_" & fieldToken & "_%_APP_%-VALU_PUSH").Press
 If Not SAPX_WaitIdle(ses, 15) Then Err.Raise 5, , "Timeout clearing column filter"
 ses.findById("wnd[2]/tbar[0]/btn[16]").Press
 ses.findById("wnd[2]/tbar[0]/btn[8]").Press
 If Not SAPX_WaitIdle(ses, 15) Then Err.Raise 5, , "Timeout clearing multiple selections"
ElseIf Not useList Then

 stage = "Enter Material directly"

 Set ctl = ses.findById(base & "ctxt" & fieldToken & "-LOW")

 SAPX_SetValue ctl, scalar

 If StrComp(SAPX_ControlValue(ctl), scalar, vbTextCompare) <> 0 Then Err.Raise 5, , "Material readback khong khop " & colId

Else

 stage = "Open multiple selection"

 ses.findById(base & "btn%_" & fieldToken & "_%_APP_%-VALU_PUSH").Press

 If Not SAPX_WaitIdle(ses, 15) Then Err.Raise 5, , "Timeout"

 ses.findById("wnd[2]/tbar[0]/btn[16]").Press

 If Not SAPX_WaitIdle(ses, 15) Then Err.Raise 5, , "Timeout khi xoa gia tri loc cu"

 stage = "Paste Material list"

 Set ws = ThisWorkbook.Worksheets("SEARCH_LIST")

 ws.columns(1).ClearContents

 ws.Cells(1, 1).Value2 = "MATERIAL"

 r = 2

 For Each p In products.keys

  ws.Cells(r, 1).NumberFormat = "@"

  ws.Cells(r, 1).Value2 = CStr(p): r = r + 1

 Next p

 If r = 2 Then Err.Raise 5, , "Material list rong"

 ws.Range("A2:A" & CStr(r - 1)).Copy

 ses.findById("wnd[2]/tbar[0]/btn[24]").Press

 If Not SAPX_WaitIdle(ses, 15) Then Err.Raise 5, , "Timeout"

 ses.findById("wnd[2]/tbar[0]/btn[8]").Press

 Application.CutCopyMode = False

 If Not SAPX_WaitIdle(ses, 15) Then Err.Raise 5, , "Timeout"

End If

stage = "Apply filter"

ses.findById("wnd[1]/tbar[0]/btn[0]").Press

If Not SAPX_WaitIdle(ses, 30) Then Err.Raise 5, , "Timeout"

If Not ses.findById("wnd[1]", False) Is Nothing Then Err.Raise 5, , "Popup filter chua dong"

SAPX_ClearGridSelection grid

Exit Sub

Fail:

detail = Err.Description

Application.CutCopyMode = False

Err.Raise vbObjectError + 4204, , "VL06O filter " & colId & " / " & stage & ": " & detail & ". Dung, khong quet toan bo grid."

End Sub

Public Sub OB_ReadVL06O(ByVal products As Object, Optional ByVal recoverDates As Boolean = False)

Dim ses As Object, grid As Object, ctl As Object, ws As Worksheet, n As Long, r As Long, count As Long

Dim ids As Collection, ni As Object, nc As Object, cDel As String, cItem As String, cMat As String, cQty As String

Dim cBatch As String, cSloc As String, cShip As String, cGM As String, cSTO As String, cSTOI As String, cUom As String, cParent As String
Dim colCreator As String, colDate As String, rowCreator As String, rowDate As String, currentUser As String

Dim arr() As Variant, final() As Variant, i As Long, j As Long, mat As String, ob As String, item As String, ship As String, sloc As String, gm As String

Dim qtyText As String, qty As Double, reason As String, dfrom As Date, seen As Object, key As String, expectedUom As String

Dim rowSloc As String, rowBatch As String, rowParent As String, rowUom As String, eligible As Long, validQty As Boolean, cachedRef As Variant

Set ses = SAPX_GetSessionForTcode("VL06O")

If Not SAPX_TcodeIs(ses, "VL06O") Then SAPX_Navigate ses, "VL06O"

VL6_OpenSelection ses

If SAPX_FindById(ses, "wnd[0]/usr/ctxtIT_WADAT-LOW") Is Nothing Then

 SAPX_Navigate ses, "VL06O"

 VL6_OpenSelection ses

End If

ship = CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2)

sloc = CStr(ThisWorkbook.Worksheets("INPUT").Range("R6").Value2)

If Not VL6_SetSelectionShipTo(ses, ship) Then Err.Raise vbObjectError + 4201, , "Thieu Ship-to tren man hinh chon VL06O. Kiem tra CONFIG VL06O_SHIP_TO_ID."

ModuleLog.WriteLog "VL06O_SELECTION", "Ship-to=" & ship & "; loc o selection, khong loc lai Item View"

Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT_ID", "wnd[0]/usr/ctxtIF_VSTEL-LOW"))

If ctl Is Nothing Then Err.Raise vbObjectError + 4201, , "Thieu Shipping Point/Receiving Pt."

SAPX_SetValue ctl, ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")

If SAPX_ControlValue(ctl) <> ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093") Then Err.Raise vbObjectError + 4201, , "Shipping Point readback khong khop."

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/ctxtIF_VSTEL-HIGH")

If Not ctl Is Nothing Then ctl.text = ""

ModuleLog.WriteLog "VL06O_SELECTION", "Shipping Point=" & ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/ctxtIT_WADAT-LOW")

If ctl Is Nothing Then Err.Raise vbObjectError + 4201, , "Thieu Planned GI From."

dfrom = DateAdd("d", -60, Date)

' RC35: exact rolling 60 days for initial and targeted readback; no year-wide recovery.

ModuleLog.WriteLog "VL06O_DATE_SCOPE", Format$(dfrom, "dd.mm.yyyy") & " to " & Format$(Date, "dd.mm.yyyy")

ctl.text = Format$(dfrom, "dd.mm.yyyy")

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/ctxtIT_WADAT-HIGH")

If ctl Is Nothing Then Err.Raise vbObjectError + 4201, , "Thieu Planned GI To."

ctl.text = Format$(Date, "dd.mm.yyyy")

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/ctxtIT_WBSTK-LOW")

If ctl Is Nothing Then Err.Raise vbObjectError + 4201, , "Thieu GM status."

ctl.text = "A"

ses.findById("wnd[0]/tbar[1]/btn[8]").Press

If Not SAPX_WaitIdle(ses, 60) Then Err.Raise vbObjectError + 4202, , "VL06O Execute timeout."

If InStr(1, SAPX_StatusText(ses), "no deliveries", vbTextCompare) > 0 Or InStr(1, SAPX_StatusText(ses), "no data", vbTextCompare) > 0 Or InStr(1, SAPX_StatusText(ses), "keine lieferungen", vbTextCompare) > 0 Then
 Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
 ws.Range("A6:P" & Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)).ClearContents
 ModuleLog.WriteLog "VL06O_EMPTY", "Ship-to=" & ship & "; " & SAPX_StatusText(ses)
 Exit Sub
End If
Set grid = VL6_EnsureItemGrid(ses)

SAPX_BuildColumnCatalogFast grid, ids, ni, nc

cDel = SAPX_ResolveColumnFast("VBELN", "VBELN,DELIVERY", True, ids, ni, nc)

cItem = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_ITEM", "POSNR"), "POSNR,DELIVERY ITEM,ITEM", True, ids, ni, nc)

cMat = SAPX_ResolveColumnFast("MATNR", "MATNR,MATERIAL", True, ids, ni, nc)

cQty = SAPX_ResolveColumnFast("LFIMG", "LFIMG,DELIVERY QTY,DLV.QTY", True, ids, ni, nc)

cBatch = SAPX_ResolveColumnFast("S_CHARG", "S_CHARG,CHARG,BATCH", True, ids, ni, nc)

cSloc = SAPX_ResolveColumnFast("LGORT", "LGORT,STORAGE LOCATION", True, ids, ni, nc)

cShip = SAPX_ResolveColumnFast("KUNNR", "KUNNR,KUNWE,SHIP-TO PARTY", True, ids, ni, nc)

cGM = SAPX_ResolveColumnFast("WBSTA", "WBSTA,WBSTK,GOODS MOVEMENT STATUS", True, ids, ni, nc)

cSTO = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_STO", "VGBEL"), "VGBEL,REFERENCE DOCUMENT", False, ids, ni, nc)

cSTOI = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_STO_ITEM", "VGPOS"), "VGPOS,POSNV,EBELP,REFERENCE ITEM,ORIGINAL ITEM,ORIG.ITEM,STO ITEM", False, ids, ni, nc)

cUom = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_UOM", "VRKME"), "VRKME,SALES UNIT,DELIVERY UNIT", True, ids, ni, nc)

cParent = SAPX_ResolveColumnFast("UECHA", "UECHA,HIGHER LEVEL BATCH ITEM", False, ids, ni, nc)
colCreator = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_CREATOR", "ERNAM"), "ERNAM,CREATED BY,ERSTELLER,PERSON RESPONSIBLE", False, ids, ni, nc)
colDate = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_DATE", "ERDAT"), "ERDAT,CREATED ON,ENTRY DATE,CREATION DATE", False, ids, ni, nc)
If colCreator <> "" Then ModuleLog.WriteLog "VL06O_CREATOR_COL", "Technical field=" & colCreator

ModuleLog.WriteLog "VL06O_FILTER", "Before=" & CStr(grid.rowCount) & "; Materials=" & products.count

VL6_ApplyALVFilter ses, grid, cMat, "", products

Set grid = VL6_DirectGrid(ses)

If grid Is Nothing Then Err.Raise vbObjectError + 4204, , "Mat grid sau loc Material."

ModuleLog.WriteLog "VL06O_FILTER", "After=" & CStr(grid.rowCount) & "; Ship-to=" & ship & "; SLoc=" & sloc

n = CLng(grid.rowCount)

If n > 20000 Then Err.Raise vbObjectError + 4202, , "VL06O >20000 rows. Giam khoang ngay."

Set seen = CreateObject("Scripting.Dictionary")

expectedUom = UCase$(ModuleHelper.ToolConfig("EXPECTED_UOM", "UN"))

If n > 0 Then ReDim arr(1 To n, 1 To 16)

For r = 0 To n - 1

 SAPX_EnsureGridRowVisible grid, r

 mat = RequiredCell(grid, r, cMat)

 If mat <> "" And Not products.exists(mat) Then Err.Raise vbObjectError + 4204, , "Material filter khong dung: " & mat

 If products.exists(mat) Then

  If RequiredCell(grid, r, cShip) <> ship Then Err.Raise vbObjectError + 4204, , "Ship-to selection khong dung: " & RequiredCell(grid, r, cShip) & "; expected=" & ship

  ob = RequiredCell(grid, r, cDel): item = RequiredCell(grid, r, cItem)

  If ob = "" Or item = "" Then Err.Raise vbObjectError + 4203, , "VL06O an Delivery/Item trung lap. Bo suppression/grouping trong layout."

  gm = RequiredCell(grid, r, cGM)

  qtyText = RequiredCell(grid, r, cQty)

  rowSloc = RequiredCell(grid, r, cSloc): rowBatch = RequiredCell(grid, r, cBatch)

  rowParent = RequiredCell(grid, r, cParent): rowUom = UCase$(RequiredCell(grid, r, cUom))

  reason = "": qty = 0

  validQty = ModuleHelper.ParseSAPIntegerQty(qtyText, qty, reason)

  If validQty Then

   reason = ""

   If qty <= 0 Then reason = "QTY <= 0"

  Else

   reason = "QTY: " & reason

  End If

  If gm <> "A" Then reason = reason & "; GM=" & gm

  If rowBatch <> "" Then reason = reason & "; DA CO BATCH"

  If rowParent <> "" And val(rowParent) <> 0 Then reason = reason & "; BATCH SPLIT CHILD"

  If Not ModuleHelper.IsCompatibleUoM(rowUom, expectedUom) Then reason = reason & "; UOM=" & rowUom & "; CAN=" & expectedUom

  If rowSloc <> "" Then reason = reason & "; DA CO SLOC=" & rowSloc
  If rowSloc = "" And Not ModuleHelper.ToolConfigBool("VL06O_ALLOW_BLANK_SLOC", True) Then reason = reason & "; BLANK SLOC DISABLED"

  key = ob & "|" & item

  If seen.exists(key) Then Err.Raise vbObjectError + 4203, , "Trung OB/Item trong VL06O: " & key

  seen(key) = True: count = count + 1

  arr(count, 1) = ob: arr(count, 2) = item: arr(count, 3) = mat

  If validQty Then arr(count, 4) = qty Else arr(count, 4) = qtyText

  arr(count, 5) = rowBatch: arr(count, 6) = ship: arr(count, 7) = rowSloc

  arr(count, 8) = RequiredCell(grid, r, cSTO): arr(count, 9) = RequiredCell(grid, r, cSTOI)
  If val(CStr(arr(count, 9))) = 0 Then arr(count, 9) = EDCProbeReferenceItem(grid, r)
  If val(CStr(arr(count, 9))) = 0 Then
   If Not mEDCReferenceCache Is Nothing Then
    key = ob & "|" & item & "|" & mat
    If mEDCReferenceCache.exists(key) Then
     cachedRef = mEDCReferenceCache(key): arr(count, 8) = cachedRef(0): arr(count, 9) = cachedRef(1)
    End If
   End If
  End If

  If reason = "" Then

   arr(count, 10) = "CANDIDATE": arr(count, 11) = "READY": eligible = eligible + 1

  Else

   arr(count, 10) = "BLOCK": arr(count, 11) = reason

   ModuleLog.WriteLog "VL06O_BLOCK", key & "; " & reason

  End If

  rowCreator = "": rowDate = ""
  If colCreator <> "" Then rowCreator = Trim$(UCase$(RequiredCell(grid, r, colCreator)))
  If colDate <> "" Then rowDate = Trim$(RequiredCell(grid, r, colDate))
  arr(count, 12) = gm: arr(count, 13) = rowUom: arr(count, 14) = rowParent
  arr(count, 15) = rowCreator: arr(count, 16) = rowDate

 End If

Next r

Set ws = ThisWorkbook.Worksheets("VL06O_DATA")

ws.Range("A6:P" & CStr(Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row))).ClearContents
ws.Range("O5").Value2 = "CREATED_BY": ws.Range("P5").Value2 = "CREATED_ON"

If count > 0 Then

 ReDim final(1 To count, 1 To 16)

 For i = 1 To count: For j = 1 To 16: final(i, j) = arr(i, j): Next j: Next i

 ws.Range("A6").Resize(count, 16).Value2 = final

 ws.Range("A5:P" & count + 5).Sort Key1:=ws.Range("A6"), Order1:=xlAscending, Key2:=ws.Range("B6"), Order2:=xlAscending, header:=xlYes

End If

OB_ClassifyCandidates ws
eligible = Application.CountIf(ws.Range("J6:J" & CStr(Application.Max(6, count + 5))), "CANDIDATE")
ModuleLog.WriteLog "VL06O", count & " dong da luu; " & eligible & " CANDIDATE; Ship-to=" & ship & "; SLoc=" & sloc

End Sub

Private Function RequiredCell(ByVal grid As Object, ByVal row As Long, ByVal col As String) As String

If col = "" Then Exit Function

On Error GoTo Fail

RequiredCell = Trim$(CStr(grid.GetCellValue(row, col)))

Exit Function

Fail:

Err.Raise vbObjectError + 4204, , "Khong doc duoc grid row=" & row & " col=" & col

End Function

Private Function TableField(ByVal tbl As Object, ByVal token As String, ByVal visibleRow As Long) As Object

Dim ch As Object, i As Long, id As String, at As Long, coords As Variant, cacheKey As String, relativeId As String

cacheKey = CStr(tbl.id) & "|" & token

If mFieldPaths Is Nothing Then Set mFieldPaths = CreateObject("Scripting.Dictionary")

If mFieldPaths.exists(cacheKey) Then

 relativeId = CStr(mFieldPaths(cacheKey)) & CStr(visibleRow) & "]"

 On Error Resume Next

 Set ch = tbl.findById(relativeId)

 On Error GoTo 0

 If Not ch Is Nothing Then

  Set TableField = ch: Exit Function

 End If

 mFieldPaths.Remove cacheKey

End If

For i = 0 To tbl.children.count - 1

 Set ch = tbl.children(CLng(i)): id = CStr(ch.id)

 If InStr(1, id, token & "[", vbTextCompare) > 0 Then

  at = InStrRev(id, "["): coords = Split(mid$(id, at + 1, Len(id) - at - 1), ",")

  If UBound(coords) = 1 Then

   If CLng(coords(1)) = visibleRow Then

    relativeId = mid$(id, Len(CStr(tbl.id)) + 2)

    mFieldPaths(cacheKey) = Left$(relativeId, InStrRev(relativeId, ","))

    Set TableField = ch: Exit Function

   End If

  End If

 End If

Next i

Err.Raise vbObjectError + 4205, , "Thieu field " & token & " row " & visibleRow & ". Can MAP layout VL02N."

End Function

Private Function VisibleItemRow(ByVal tbl As Object, ByVal item As String) As Long

Dim r As Long, actual As Long, fld As Object, value As String

VisibleItemRow = -1: actual = CLng(tbl.VerticalScrollbar.Position)

For r = 0 To CLng(tbl.VisibleRowCount) - 1

 If actual + r >= CLng(tbl.rowCount) Then Exit For

 Set fld = TableField(tbl, "LIPS-POSNR", r): value = Trim$(CStr(fld.text))

 If value <> "" And val(value) = val(item) Then VisibleItemRow = r: Exit Function

Next r

End Function

Private Function LocateItem(ByVal ses As Object, ByVal item As String, ByRef tbl As Object) As Long

Dim r As Long, page As Long, cnt As Long, pos As Long, actual As Long, v As String, fld As Object
Dim key As String, data As Variant, expectedAbsRow As Long, row As Long, fldMat As Object

key = CStr(CLng(val(item)))

Set tbl = DeliveryOverview(ses)
If tbl Is Nothing Then Err.Raise vbObjectError + 4205, , "Khong co VL02N overview table."
page = CLng(tbl.VisibleRowCount)
If page < 1 Then Err.Raise vbObjectError + 4205, , "VL02N table khong co dong nhin thay."
cnt = CLng(tbl.rowCount)

' Step 1: Check expected row from mOverviewIndex if available
If Not mOverviewIndex Is Nothing Then
 If mOverviewIndex.exists(key) Then
  data = mOverviewIndex(key)
  expectedAbsRow = CLng(data(3))
  actual = CLng(tbl.VerticalScrollbar.Position)
  row = expectedAbsRow - actual
  If row >= 0 And row < page Then
   Set fld = TableField(tbl, "LIPS-POSNR", row)
   v = Trim$(CStr(fld.text))
   If val(v) = val(item) And v <> "" Then
    LocateItem = row
    Exit Function
   End If
  End If
 End If
End If

' Step 2: Fast scan visible rows on CURRENT page (0 To page - 1) before scrolling
actual = CLng(tbl.VerticalScrollbar.Position)
For r = 0 To page - 1
 If actual + r >= cnt Then Exit For
 Set fld = TableField(tbl, "LIPS-POSNR", r)
 v = Trim$(CStr(fld.text))
 If val(v) = val(item) And v <> "" Then
  If Not mOverviewIndex Is Nothing Then
   If mOverviewIndex.exists(key) Then
    data = mOverviewIndex(key): data(3) = actual + r: mOverviewIndex(key) = data
   End If
  End If
  LocateItem = r
  Exit Function
 End If
Next r

' Step 3: If expected row is on another page, jump directly to that position
If Not mOverviewIndex Is Nothing Then
 If mOverviewIndex.exists(key) Then
  data = mOverviewIndex(key)
  expectedAbsRow = CLng(data(3))
  If expectedAbsRow <> actual Then
   tbl.VerticalScrollbar.Position = expectedAbsRow
   Set tbl = DeliveryOverview(ses)
   actual = CLng(tbl.VerticalScrollbar.Position)
   For r = 0 To page - 1
    If actual + r >= cnt Then Exit For
    Set fld = TableField(tbl, "LIPS-POSNR", r)
    v = Trim$(CStr(fld.text))
    If val(v) = val(item) And v <> "" Then
     data(3) = actual + r: mOverviewIndex(key) = data
     LocateItem = r
     Exit Function
    End If
   Next r
  End If
 End If
End If

' Step 4: Fallback page-by-page scroll with Early Exit (never set mOverviewIndex = Nothing)
For pos = 0 To cnt - 1 Step page
 If CLng(tbl.VerticalScrollbar.Position) <> pos Then tbl.VerticalScrollbar.Position = pos
 Set tbl = DeliveryOverview(ses)
 actual = CLng(tbl.VerticalScrollbar.Position)
 For r = 0 To page - 1
  If actual + r >= cnt Then Exit For
  Set fld = TableField(tbl, "LIPS-POSNR", r)
  v = Trim$(CStr(fld.text))
  Set fldMat = TableField(tbl, "LIPS-MATNR", r)
  ' Early stop if row is empty (no more items)
  If v = "" And Trim$(CStr(fldMat.text)) = "" Then Exit For
  If val(v) = val(item) And v <> "" Then
   If Not mOverviewIndex Is Nothing Then
    If mOverviewIndex.exists(key) Then
     data = mOverviewIndex(key): data(3) = actual + r: mOverviewIndex(key) = data
    End If
   End If
   LocateItem = r
   Exit Function
  End If
 Next r
 If v = "" And Trim$(CStr(fldMat.text)) = "" Then Exit For
Next pos

Err.Raise vbObjectError + 4206, , "Khong tim thay delivery item " & item

End Function

Private Function DeliveryNumberField(ByVal ses As Object) As Object

Set DeliveryNumberField = SAPX_FindById(ses, "wnd[0]/usr/ctxtLIKP-VBELN")

If DeliveryNumberField Is Nothing Then Set DeliveryNumberField = SAPX_FindById(ses, "wnd[0]/usr/ctxtRV50A-VBELN")

If DeliveryNumberField Is Nothing Then Set DeliveryNumberField = SAPX_FindByTokens(ses, "LIKP-VBELN", True)

If DeliveryNumberField Is Nothing Then Set DeliveryNumberField = SAPX_FindByTokens(ses, "RV50A-VBELN", True)

End Function

Private Function IsDeliverySyncError(ByVal ses As Object) As Boolean

Dim txt As String, typ As String

txt = LCase$(SAPX_StatusText(ses)): typ = SAPX_StatusType(ses)

IsDeliverySyncError = ((typ = "E" Or typ = "A") And (InStr(txt, "not yet synchronized with remote system") > 0 Or InStr(txt, "not synchronized with remote system") > 0))

End Function

Private Sub AssertDeliveryPopup(ByVal ses As Object)

Dim popup As Object

Set popup = SAPX_FindById(ses, "wnd[1]")

If Not popup Is Nothing Then Err.Raise vbObjectError + 4208, , "SAP popup can xu ly: " & CStr(popup.text)

End Sub

Private Sub AssertDeliveryEntry(ByVal ses As Object)

Dim popup As Object

Set popup = SAPX_FindById(ses, "wnd[1]")

If Not popup Is Nothing Then Err.Raise vbObjectError + 4208, , "SAP popup can xu ly: " & CStr(popup.text)

' An old synchronization status may survive F12; the new OB submit is checked strictly.

If Not IsDeliverySyncError(ses) Then AssertNoPopup ses

End Sub

Private Sub OpenDelivery(ByVal ses As Object, ByVal ob As String, Optional ByVal displayOnly As Boolean = False, Optional ByVal reopenAfterSave As Boolean = False)

Dim fld As Object, overview As Object, overviewTab As Object, attempt As Long, maxAttempts As Long, tx As String

Dim started As Double, failure As String

Set mOverviewIndex = Nothing

If displayOnly Then tx = "VL03N" Else tx = "VL02N"

If Not displayOnly And ModuleHelper.ToolConfigBool("VL02N_SESSION_REUSE", True) Then

 If SAPX_TcodeIs(ses, tx) Then

  AssertDeliveryPopup ses

  Set overview = DeliveryOverview(ses)

  If overview Is Nothing Then

   Set fld = DeliveryNumberField(ses)

   If Not fld Is Nothing Then reopenAfterSave = True

  ElseIf mSavedForReuse <> "" Then

   Set fld = DeliveryNumberField(ses)

   If Not fld Is Nothing Then

    If val(CStr(fld.text)) = val(mSavedForReuse) And SAPX_StatusType(ses) = "S" And IsSaveReceipt(SAPX_StatusText(ses), mSavedForReuse, ses) Then

     ses.findById("wnd[0]").sendVKey 3

     If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4207, , "Delivery reuse Back timeout."

     AssertDeliveryPopup ses

     Set overview = DeliveryOverview(ses)

     If overview Is Nothing Then

      Set fld = DeliveryNumberField(ses)

      If Not fld Is Nothing Then reopenAfterSave = True

     End If

    End If

   End If

  End If

 End If

End If

If Not reopenAfterSave Or Not SAPX_TcodeIs(ses, tx) Then

 SAPX_Navigate ses, tx

 ModuleLog.WriteLog "DELIVERY_NAV", "fresh transaction", ob

Else

 ModuleLog.WriteLog "DELIVERY_NAV", "reuse entry screen", ob

End If

mSavedForReuse = ""

If Not SAPX_TcodeIs(ses, tx) Then Err.Raise vbObjectError + 4207, , "Sai transaction khi mo Delivery."

' A saved delivery can leave its number on the selection screen.

AssertDeliveryPopup ses

maxAttempts = ModuleHelper.ToolConfigLong("VL02N_OPEN_ATTEMPTS", 2, 1, 3)

For attempt = 1 To maxAttempts

 Set fld = Nothing: started = Timer

 Do

  AssertDeliveryPopup ses

  Set overview = DeliveryOverview(ses)

  If Not overview Is Nothing Then

   ' Save has finished. Leave the persisted overview once, then use the entry field.

   If reopenAfterSave Then ses.findById("wnd[0]").sendVKey 3 Else ses.findById("wnd[0]").sendVKey 12

   If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4207, , tx & " Return timeout."

  Else

   Set fld = DeliveryNumberField(ses)

   If Not fld Is Nothing Then Exit Do

  End If

  DoEvents: SAPX_Sleep 25

 Loop While SAPX_Elapsed(started) < 10

 If fld Is Nothing Then Err.Raise vbObjectError + 4207, , "Thieu o Delivery " & tx

 ' Keep the existing number when it already matches. Always press Enter to open.

 If Trim$(CStr(fld.text)) <> ob Then fld.text = ob

 If val(CStr(fld.text)) <> val(ob) Then Err.Raise vbObjectError + 4207, , "Delivery selection readback khong khop."

 ses.findById("wnd[0]").sendVKey 0

 If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4207, , "Open Delivery timeout."

 ModuleLog.WriteLog "DELIVERY_OPEN", tx & "; attempt=" & attempt & "; status=" & SAPX_StatusText(ses), ob

 If Not IsDeliverySyncError(ses) Then

  AssertNoPopup ses

  Set overviewTab = SAPX_FindById(ses, "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01")

  If Not overviewTab Is Nothing Then

   overviewTab.Select

   If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4207, , "Overview tab timeout."

   AssertNoPopup ses

  End If

  started = Timer

  Do

   Set overview = DeliveryOverview(ses)

   If Not overview Is Nothing Then Exit Do

   DoEvents: SAPX_Sleep 25

  Loop While SAPX_Elapsed(started) < 10

  If overview Is Nothing Then Err.Raise vbObjectError + 4207, , "Khong co overview sau khi mo " & tx

  Set fld = DeliveryNumberField(ses)

  If fld Is Nothing Then Err.Raise vbObjectError + 4207, , "Khong doc lai duoc Delivery hien tai."

  If val(CStr(fld.text)) <> val(ob) Then Err.Raise vbObjectError + 4207, , "Delivery readback khong khop."

  Exit Sub

 End If

 failure = SAPX_StatusText(ses)

 If attempt < maxAttempts Then

  ModuleLog.WriteLog "DELIVERY_SYNC_RETRY", "Retry truoc sua du lieu: " & failure, ob

  started = Timer

  Do While SAPX_Elapsed(started) < 1: DoEvents: SAPX_Sleep 25: Loop

  ' No field edit or Save has occurred; return to the entry screen only.

  ses.findById("wnd[0]").sendVKey 12

  If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4207, , "Sync retry Cancel timeout."

  AssertDeliveryEntry ses

 End If

Next attempt

Err.Raise vbObjectError + 4207, , "WAIT_REMOTE_SYNC OB " & ob & ": " & failure

End Sub

Public Function OB_ResolveUnknown(ByVal ob As String, ByVal allocations As Collection, Optional ByVal destination As String = "") As String

Dim ses As Object, itemMap As Object, batches As Object, a As Variant, key As String, planned As Variant

Dim item As Variant, row As Long, tbl As Object, fld As Object, unchanged As Boolean, qty As Double

Set ses = SAPX_GetSessionForTcode("VL03N")

ModuleMain.AssertRunIdentity

OpenDelivery ses, ob, True

Set itemMap = CreateObject("Scripting.Dictionary")

For Each a In allocations

 key = CStr(a(2))

 If Not itemMap.exists(key) Then

  Set batches = CreateObject("Scripting.Dictionary")

  itemMap.Add key, Array(CStr(a(3)), 0#, batches, CDbl(a(9)), CStr(a(10)))

 End If

 planned = itemMap(key): planned(1) = CDbl(planned(1)) + CDbl(a(6)): Set batches = planned(2)

 If Not batches.exists(CStr(a(4))) Then batches(CStr(a(4))) = 0#

 batches(CStr(a(4))) = CDbl(batches(CStr(a(4)))) + CDbl(a(6)): itemMap(key) = planned

Next a

unchanged = True

For Each item In itemMap.keys

 planned = itemMap(item): row = LocateItem(ses, CStr(item), tbl)

 Set fld = TableField(tbl, "LIPS-MATNR", row)

 If Trim$(CStr(fld.text)) <> CStr(planned(0)) Then Err.Raise vbObjectError + 4216, , "Recovery Material khong khop."

 Set fld = TableField(tbl, "LIPSD-G_LFIMG", row)

 If ModuleHelper.CleanDouble(fld.text) <> CDbl(planned(3)) Then unchanged = False

 Set fld = TableField(tbl, "LIPS-CHARG", row)

 If Trim$(CStr(fld.text)) <> "" Then unchanged = False

 Set fld = TableField(tbl, "LIPS-LGORT", row)

 If Trim$(CStr(fld.text)) <> CStr(planned(4)) Then unchanged = False

Next item

ModuleMain.AssertRunIdentity

If unchanged Then

 OB_ResolveUnknown = "UNSAVED"

Else

 If destination = "" Then destination = CStr(ThisWorkbook.Worksheets("INPUT").Range("R6").Value2)

 VerifyDeliveryPlan ses, itemMap, destination

 ModuleMain.AssertRunIdentity

 OB_ResolveUnknown = "DONE"

End If

End Function

Private Sub AssertNoPopup(ByVal ses As Object)

Dim p As Object

Set p = SAPX_FindById(ses, "wnd[1]")

If Not p Is Nothing Then Err.Raise vbObjectError + 4208, , "SAP popup can xu ly: " & CStr(p.text)

If SAPX_StatusType(ses) = "E" Or SAPX_StatusType(ses) = "A" Then Err.Raise vbObjectError + 4208, , "SAP: " & SAPX_StatusText(ses)

End Sub

Private Sub SetTableText(ByVal tbl As Object, ByVal field As String, ByVal row As Long, ByVal text As String)

Dim f As Object, started As Double

started = Timer

Set f = TableField(tbl, field, row): f.text = text

If Trim$(CStr(f.text)) <> text Then Err.Raise vbObjectError + 4209, , "Readback field " & field & " khong khop."

If SAPX_Elapsed(started) > 0.5 Then ModuleLog.WriteLog "SLOW_FIELD_WRITE", field & "; row=" & row & "; sec=" & Format$(SAPX_Elapsed(started), "0.00")

End Sub

Private Function DeliveryOverview(ByVal ses As Object) As Object

Set DeliveryOverview = SAPX_FindById(ses, "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER")

If DeliveryOverview Is Nothing Then Set DeliveryOverview = SAPX_FindByTokens(ses, "tblSAPMV50ATC_LIPS_OVER", False)

End Function

Public Function IsSaveReceipt(ByVal message As String, ByVal ob As String, Optional ByVal ses As Object = Nothing) As Boolean

Dim t As String, msgId As String, msgNo As String, obTrim As String

t = LCase$(Trim$(message))

obTrim = LCase$(Trim$(ob))

' Check multi-language save confirmations

If obTrim <> "" And InStr(1, t, obTrim, vbBinaryCompare) > 0 Then

 If InStr(t, "was saved") > 0 Or InStr(t, "has been saved") > 0 Or InStr(t, "gesichert") > 0 Or _
    InStr(t, "da duoc luu") > 0 Or InStr(t, "da luu") > 0 Or InStr(t, "saved") > 0 Then

  IsSaveReceipt = True

  Exit Function

 End If

End If

' Priority: SAP status bar MessageId and MessageNumber

If Not ses Is Nothing Then

 msgId = SAPX_MessageId(ses)

 msgNo = SAPX_MessageNumber(ses)

 If (msgId = "VL" Or msgId = "V1" Or msgId = "/SCWM/DELIVERY") And _
    (msgNo = "311" Or msgNo = "009" Or msgNo = "302" Or msgNo = "101" Or msgNo = "150") Then

  If obTrim = "" Or InStr(1, t, obTrim, vbBinaryCompare) > 0 Then

   IsSaveReceipt = True

   Exit Function

  End If

 End If

End If

End Function

Private Function OverviewSnapshot(ByVal ses As Object, Optional ByVal indexOnly As Boolean = False) As Object

Dim tbl As Object, fld As Object, rows As Object, pos As Long, cnt As Long, page As Long, actual As Long, rr As Long

Dim itemNo As String, material As String, batch As String, qty As Double, sloc As String, savedPosition As Long, lastPosition As Long, existing As Variant

Set rows = CreateObject("Scripting.Dictionary")

Set tbl = DeliveryOverview(ses)

If tbl Is Nothing Then Err.Raise vbObjectError + 4217, , "Thieu Overview khi kiem tra danh sach item."

savedPosition = CLng(tbl.VerticalScrollbar.Position): lastPosition = -1

cnt = CLng(tbl.rowCount): page = CLng(tbl.VisibleRowCount)

If page < 1 Then Err.Raise vbObjectError + 4217, , "Overview khong co dong hien thi."

For pos = 0 To cnt - 1 Step page

 If CLng(tbl.VerticalScrollbar.Position) <> pos Then If CLng(tbl.VerticalScrollbar.Position) <> pos Then tbl.VerticalScrollbar.Position = pos

 Set tbl = DeliveryOverview(ses): actual = CLng(tbl.VerticalScrollbar.Position)

 If actual <= lastPosition Then Exit For

 lastPosition = actual

 For rr = 0 To CLng(tbl.VisibleRowCount) - 1

  If actual + rr >= cnt Then Exit For

  Set fld = TableField(tbl, "LIPS-POSNR", rr): itemNo = Trim$(CStr(fld.text))

  Set fld = TableField(tbl, "LIPS-MATNR", rr): material = Trim$(CStr(fld.text))

  If itemNo = "" And material = "" Then GoTo finished

  If itemNo = "" Or val(itemNo) <= 0 Or material = "" Then Err.Raise vbObjectError + 4217, , "Dong item/material khong day du tai " & CStr(actual + rr)

  itemNo = CStr(CLng(val(itemNo)))

  If rows.exists(itemNo) Then

   existing = rows(itemNo)

   If CLng(existing(3)) <> actual + rr Then Err.Raise vbObjectError + 4217, , "Trung item " & itemNo

  End If

  batch = "": qty = 0: sloc = ""

  If Not indexOnly Then

  Set fld = TableField(tbl, "LIPS-CHARG", rr): batch = Trim$(CStr(fld.text))

  Set fld = TableField(tbl, "LIPSD-G_LFIMG", rr): qty = ModuleHelper.CleanDouble(fld.text)

  Set fld = TableField(tbl, "LIPS-LGORT", rr): sloc = Trim$(CStr(fld.text))

  End If

  rows(itemNo) = Array(material, batch, qty, actual + rr, sloc)

NextOverviewRow:
  Next rr

Next pos

finished:

Set tbl = DeliveryOverview(ses)

If savedPosition >= LiveOverviewRowCount(rows) Then savedPosition = 0

If CLng(tbl.VerticalScrollbar.Position) <> savedPosition Then tbl.VerticalScrollbar.Position = savedPosition

Set mOverviewIndex = rows: Set OverviewSnapshot = rows

End Function

Private Function LiveOverviewRowCount(ByVal rows As Object) As Long

Dim item As Variant, data As Variant

For Each item In rows.keys

 data = rows(item)

 If CLng(data(3)) + 1 > LiveOverviewRowCount Then LiveOverviewRowCount = CLng(data(3)) + 1

Next item

End Function

Private Function JoinDictKeys(ByVal d As Object) As String
Dim k As Variant, s As String
If d Is Nothing Then Exit Function
For Each k In d.keys
 s = s & IIf(s <> "", ",", "") & CStr(k)
Next k
JoinDictKeys = s
End Function


Private Function PlannedItemKeys(ByVal itemMap As Object) As Object

Dim keep As Object, item As Variant

Set keep = CreateObject("Scripting.Dictionary")

For Each item In itemMap.keys: keep(CStr(CLng(val(CStr(item))))) = True: Next item

Set PlannedItemKeys = keep

End Function

Private Sub ClearOverviewSelection(ByVal ses As Object)

Dim tbl As Object, pos As Long, cnt As Long, page As Long, rr As Long, actual As Long, rows As Object

If mOverviewIndex Is Nothing Then Set mOverviewIndex = OverviewSnapshot(ses, True)

Set rows = mOverviewIndex

Set tbl = DeliveryOverview(ses)

cnt = LiveOverviewRowCount(rows): page = CLng(tbl.VisibleRowCount)

If page < 1 Then Err.Raise vbObjectError + 4217, , "Khong the bo chon item."

For pos = 0 To cnt - 1 Step page

 If CLng(tbl.VerticalScrollbar.Position) <> pos Then tbl.VerticalScrollbar.Position = pos

 Set tbl = DeliveryOverview(ses)

 actual = CLng(tbl.VerticalScrollbar.Position)

 For rr = 0 To CLng(tbl.VisibleRowCount) - 1

  If actual + rr >= cnt Then Exit For

  If tbl.GetAbsoluteRow(actual + rr).selected Then tbl.GetAbsoluteRow(actual + rr).selected = False

 Next rr

Next pos

End Sub

Private Sub RemoveUnplannedItems(ByVal ses As Object, ByVal ob As String, ByVal itemMap As Object, ByRef editStarted As Boolean)

Dim before As Object, live As Object, keep As Object, dropItems As Object, item As Variant, rowData As Variant, originalData As Variant
Dim tbl As Object, fld As Object, btn As Object, popup As Object, pos As Long, cnt As Long, page As Long, rr As Long, actual As Long
Dim itemNo As String, selected As Object, dropPositions As Object, selectedRow As Object, wantSelected As Boolean, started As Double, stage As String, failure As String, errorNumber As Long
Dim dropKeysStr As String, keepKeysStr As String, itm As Variant, material As String, batch As String, qty As Double, sloc As String

On Error GoTo Fail

stage = "READ_ITEMS"
started = Timer

Set keep = PlannedItemKeys(itemMap)
If mOverviewIndex Is Nothing Then Set mOverviewIndex = OverviewSnapshot(ses)
Set before = mOverviewIndex

If keep.count = 0 Then
 ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Ke hoach rong; khong the xu ly item", ob
 Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Ke hoach rong; khong the xu ly item."
End If

' RC21 Rule 2: Mandatory deletion of ALL unplanned items
' Unplanned items are strictly identified by Planned Delivery Item ID (keep).
' Delete all items not in keep, even if they have Batch, same Material, or are split children.
Set dropItems = CreateObject("Scripting.Dictionary")
Set selected = CreateObject("Scripting.Dictionary")
Set dropPositions = CreateObject("Scripting.Dictionary")

For Each item In before.keys
 If Not keep.exists(CStr(item)) Then
  rowData = before(item)
  ' SAFETY INVARIANT: If this item belongs to a material in our input request,
  ' NEVER delete it from the Outbound Delivery! Keep it intact!
  If ModuleMain.InputProducts().exists(CStr(rowData(0))) Then
   keep(CStr(item)) = True
  Else
   dropItems(CStr(item)) = rowData
   dropPositions(CStr(CLng(rowData(3)))) = CStr(item)
  End If
 End If
Next item

If dropItems.count = 0 Then
 ' No unplanned items to delete
 Exit Sub
End If

ModuleMain.AssertRunIdentity

Set btn = SAPX_FindByTokens(ses, "btnBT_POLO_T", False)
If btn Is Nothing Then
 ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Thieu nut xoa BT_POLO_T", ob
 Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Thieu nut xoa BT_POLO_T theo VBS."
End If

Set tbl = DeliveryOverview(ses): cnt = LiveOverviewRowCount(before): page = CLng(tbl.VisibleRowCount)
If page < 1 Then
 ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Overview khong co dong hien thi", ob
 Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Overview khong co dong hien thi."
End If

stage = "SELECT_ITEMS"

' One pass selects all unwanted items, including rows on other pages.
For pos = 0 To cnt - 1 Step page
 If CLng(tbl.VerticalScrollbar.Position) <> pos Then tbl.VerticalScrollbar.Position = pos
 Set tbl = DeliveryOverview(ses): actual = CLng(tbl.VerticalScrollbar.Position)
 For rr = 0 To CLng(tbl.VisibleRowCount) - 1
  If actual + rr >= cnt Then Exit For
  wantSelected = dropPositions.exists(CStr(actual + rr))
  Set selectedRow = tbl.GetAbsoluteRow(actual + rr)
  If wantSelected Then
   item = CStr(dropPositions(CStr(actual + rr))): rowData = dropItems(item)
   Set fld = TableField(tbl, "LIPS-POSNR", rr)
   If val(Trim$(CStr(fld.text))) <> val(CStr(item)) Then
    ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Item thay doi truoc xoa " & item, ob
    Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Item thay doi truoc xoa " & item
   End If
   Set fld = TableField(tbl, "LIPS-MATNR", rr)
   If Trim$(CStr(fld.text)) <> CStr(rowData(0)) Then
    ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Material thay doi truoc xoa " & item, ob
    Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Material thay doi truoc xoa " & item
   End If
  End If
  If CBool(selectedRow.selected) <> wantSelected Then selectedRow.selected = wantSelected
  If CBool(selectedRow.selected) <> wantSelected Then
   ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Selection readback khong khop item " & item, ob
   Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Selection readback khong khop."
  End If
  If wantSelected Then selected(CStr(item)) = True
 Next rr
Next pos

If selected.count <> dropItems.count Then
 ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Chua chon du item can xoa: " & selected.count & "/" & dropItems.count, ob
 Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Chua chon du item can xoa."
End If

ModuleMain.AssertRunIdentity

Set btn = SAPX_FindByTokens(ses, "btnBT_POLO_T", False)
If btn Is Nothing Then
 ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Mat nut xoa sau chon item", ob
 Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Mat nut xoa sau chon item."
End If

stage = "DELETE_BUTTON"
editStarted = True
btn.Press

Set mOverviewIndex = Nothing
If Not SAPX_WaitIdle(ses, 30) Then
 ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Bulk delete timeout", ob
 Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Bulk delete timeout."
End If

Set popup = SAPX_FindById(ses, "wnd[1]")
If Not popup Is Nothing Then
 Set btn = SAPX_FindById(ses, "wnd[1]/usr/btnSPOP-OPTION1")
 If btn Is Nothing Then
  ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Popup xoa khac VBS: " & CStr(popup.text), ob
  Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Popup xoa khac VBS: " & CStr(popup.text)
 End If
 stage = "CONFIRM_DELETE"
 btn.Press
 If Not SAPX_WaitIdle(ses, 30) Then
  ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Confirm bulk delete timeout", ob
  Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Confirm bulk delete timeout."
 End If
End If

AssertNoPopup ses

stage = "VERIFY_DELETE"

Set tbl = DeliveryOverview(ses)
If Not tbl Is Nothing Then
 If CLng(tbl.VerticalScrollbar.Position) <> 0 Then tbl.VerticalScrollbar.Position = 0
 Set tbl = DeliveryOverview(ses)
End If

Set live = CreateObject("Scripting.Dictionary")
If Not tbl Is Nothing And CLng(tbl.VisibleRowCount) > 0 Then
 For rr = 0 To CLng(tbl.VisibleRowCount) - 1
  Set fld = TableField(tbl, "LIPS-POSNR", rr): itemNo = Trim$(CStr(fld.text))
  Set fld = TableField(tbl, "LIPS-MATNR", rr): material = Trim$(CStr(fld.text))
  If itemNo = "" And material = "" Then Exit For
  If itemNo <> "" And val(itemNo) > 0 And material <> "" Then
   itemNo = CStr(CLng(val(itemNo)))
   Set fld = TableField(tbl, "LIPS-CHARG", rr): batch = Trim$(CStr(fld.text))
   Set fld = TableField(tbl, "LIPSD-G_LFIMG", rr): qty = ModuleHelper.CleanDouble(fld.text)
   Set fld = TableField(tbl, "LIPS-LGORT", rr): sloc = Trim$(CStr(fld.text))
   live(itemNo) = Array(material, batch, qty, rr, sloc)
  End If
 Next rr
End If

If live.count < keep.count Then
 Set live = OverviewSnapshot(ses)
End If

If live.count <> keep.count Then
 ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "So item sau xoa (" & live.count & ") khong khop ke hoach (" & keep.count & ")", ob
 Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: So item sau xoa khong khop ke hoach."
End If

For Each item In keep.keys
 If Not live.exists(CStr(item)) Then
  ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Mat item can giu " & item, ob
  Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Mat item can giu " & item
 End If
Next item

For Each item In dropItems.keys
 If live.exists(CStr(item)) Then
  ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Item ngoai ke hoach van ton tai sau xoa: " & item, ob
  Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Item ngoai ke hoach van ton tai sau xoa: " & item
 End If
Next item

dropKeysStr = JoinDictKeys(dropItems)
keepKeysStr = JoinDictKeys(keep)

ModuleLog.WriteLog "REMOVE_UNPLANNED_ITEMS", "Removed=" & dropItems.count & "; DeletedIDs=" & dropKeysStr & "; KeptIDs=" & keepKeysStr & "; sec=" & Format$(SAPX_Elapsed(started), "0.00"), ob

Set mOverviewIndex = live
Exit Sub

Fail:
errorNumber = Err.Number: failure = Err.Description
Set mOverviewIndex = Nothing
If InStr(1, failure, "STOP_UNPLANNED_ITEM_DELETE_FAILED", vbTextCompare) = 0 Then
 ModuleLog.WriteLog "STOP_UNPLANNED_ITEM_DELETE_FAILED", "Delete " & stage & "; SAP error=" & CStr(errorNumber) & ": " & failure, ob
 Err.Raise vbObjectError + 4217, , "STOP_UNPLANNED_ITEM_DELETE_FAILED: Delete " & stage & "; " & failure
Else
 Err.Raise errorNumber, , failure
End If

End Sub

Private Function SplitTableAt(ByVal ses As Object, ByVal absoluteRow As Long, ByRef visibleRow As Long, Optional ByVal currentTable As Object = Nothing) As Object

Dim tbl As Object, actual As Long

Set tbl = currentTable

If tbl Is Nothing Then Set tbl = SAPX_FindByTokens(ses, "tblSAPMV50ATC_LIPS_CHND", False)

If tbl Is Nothing Then Err.Raise vbObjectError + 4212, , "Thieu Split table."

If absoluteRow < 0 Or absoluteRow >= CLng(tbl.rowCount) Then Err.Raise vbObjectError + 4212, , "Split row ngoai pham vi."

actual = CLng(tbl.VerticalScrollbar.Position)

If absoluteRow < actual Or absoluteRow >= actual + CLng(tbl.VisibleRowCount) Then

 tbl.VerticalScrollbar.Position = absoluteRow

 If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4212, , "Split scroll timeout."

 AssertNoPopup ses

 Set tbl = SAPX_FindByTokens(ses, "tblSAPMV50ATC_LIPS_CHND", False)

 If tbl Is Nothing Then Err.Raise vbObjectError + 4212, , "Mat Split table sau scroll."

 actual = CLng(tbl.VerticalScrollbar.Position)

End If

visibleRow = absoluteRow - actual

If visibleRow < 0 Or visibleRow >= CLng(tbl.VisibleRowCount) Then Err.Raise vbObjectError + 4212, , "Split scrollbar khong den dong dich."

Set SplitTableAt = tbl

End Function

Private Sub AssertSplitEmpty(ByVal ses As Object, ByVal neededRows As Long)

Dim tbl As Object, fld As Object, r As Long, vr As Long, n As Long, maxCheckRows As Long

Set tbl = SplitTableAt(ses, 0, vr): n = CLng(tbl.rowCount)

If n > 10000 Or neededRows > n Then Err.Raise vbObjectError + 4212, , "Split table khong du dong an toan."

' RC22 Optimization: Only verify the target slots needed for allocation (up to neededRows)
maxCheckRows = neededRows
If maxCheckRows > n Then maxCheckRows = n

For r = 0 To maxCheckRows - 1

 Set tbl = SplitTableAt(ses, r, vr, tbl)

 Set fld = TableField(tbl, "LIPS-CHARG", vr)

 If Trim$(CStr(fld.text)) <> "" Then Err.Raise vbObjectError + 4212, , "Item da co split; khong ghi de split cu."

 Set fld = TableField(tbl, "LIPS-LFIMG", vr)

 If ModuleHelper.CleanDouble(fld.text) <> 0 Then Err.Raise vbObjectError + 4212, , "Split co Qty cu; khong ghi de."

Next r

End Sub

Private Sub ValidateSplitRows(ByVal ses As Object, ByVal batches As Object, ByVal sloc As String, ByVal expectedTotal As Double, Optional ByVal ob As String = "")

Dim tbl As Object, fld As Object, r As Long, vr As Long, n As Long, batch As String, qty As Double, total As Double, verified As Object
Dim splitSloc As String

Set verified = CreateObject("Scripting.Dictionary")
Set tbl = SplitTableAt(ses, 0, vr): n = CLng(tbl.rowCount)

If n > 10000 Then
 ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Split table vuot gioi han an toan", ob
 Err.Raise vbObjectError + 4212, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Split table vuot gioi han"
End If

For r = 0 To n - 1
 Set tbl = SplitTableAt(ses, r, vr, tbl)
 Set fld = TableField(tbl, "LIPS-CHARG", vr): batch = Trim$(CStr(fld.text))
 Set fld = TableField(tbl, "LIPS-LFIMG", vr): qty = ModuleHelper.CleanDouble(fld.text)
 Set fld = TableField(tbl, "LIPS-LGORT", vr): splitSloc = Trim$(CStr(fld.text))

 If batch <> "" Then
  If Not batches.exists(batch) Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Batch split ngoai ke hoach " & batch, ob
   Err.Raise vbObjectError + 4212, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Batch split ngoai ke hoach " & batch
  End If
  If verified.exists(batch) Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Batch split trung " & batch, ob
   Err.Raise vbObjectError + 4212, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Batch split trung " & batch
  End If

  If splitSloc = "" Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "MISSING_SLOC: Split row batch " & batch & " thieu SLoc", ob
   Err.Raise vbObjectError + 4212, , "PRESAVE_GATE_FAIL: MISSING_SLOC Split row batch " & batch & " thieu SLoc"
  ElseIf splitSloc <> sloc Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "SLOC_MISMATCH: Split row batch " & batch & " SLoc=" & splitSloc & "; expected=" & sloc, ob
   Err.Raise vbObjectError + 4212, , "PRESAVE_GATE_FAIL: SLOC_MISMATCH Split row batch " & batch & " SLoc=" & splitSloc
  End If

  If qty <> CDbl(batches(batch)) Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "SPLIT_QTY_MISMATCH: Split row batch " & batch & " Qty=" & qty & "; expected=" & CDbl(batches(batch)), ob
   Err.Raise vbObjectError + 4212, , "PRESAVE_GATE_FAIL: SPLIT_QTY_MISMATCH Split row batch " & batch
  End If

  verified(batch) = True: total = total + qty
 ElseIf qty <> 0 Then
  ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "MISSING_BATCH: Split row co Qty=" & qty & " nhung thieu Batch", ob
  Err.Raise vbObjectError + 4212, , "PRESAVE_GATE_FAIL: MISSING_BATCH Split row co Qty nhung thieu Batch"
 End If
Next r

If total <> expectedTotal Or verified.count <> batches.count Then
 ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "SPLIT_QTY_MISMATCH: Tong Batch Split khong khop (" & total & " vs " & expectedTotal & "); count=" & verified.count & "/" & batches.count, ob
 Err.Raise vbObjectError + 4212, , "PRESAVE_GATE_FAIL: SPLIT_QTY_MISMATCH Tong Batch Split khong khop."
End If

End Sub

Public Sub OB_Apply(ByVal ob As String, ByVal allocations As Collection, ByRef result As String)

Dim ses As Object, itemMap As Object, item As Variant, batches As Object, a As Variant, tbl As Object, f As Object, btn As Object
Dim row As Long, key As String, qty As Double, currentQty As Double, b As Variant, total As Double, sloc As String
Dim i As Long, splitTable As Object, splitRow As Long, page As Long, rr As Long, start As Double, oldStatus As String
Dim planned As Variant, verifyBatch As String, verifySloc As String, verifyQty As Double, failText As String
Dim editStarted As Boolean, saveAcknowledged As Boolean, phase As String, validatedSplits As Object, splitStarted As Double
Dim applyStarted As Double, phaseStarted As Double, preflightData As Variant
Dim initialItemCount As Long, initialSlocCount As Long, itmKey As Variant, itmData As Variant
Dim targetMat As String, targetSTO As String, firstAlloc As Variant
Dim msgId As String, msgNo As String, statusText As String, statusType As String, isSaved As Boolean, postSavePass As Boolean, fastSaveEnabled As Boolean, isCleanReceipt As Boolean, requiresAudit As Boolean

On Error GoTo Fail

applyStarted = Timer: mWaitSeconds = 0: mMeasureWait = True
result = "": LastApplyState = "": LastConflictEligible = False: LastConflictKind = "": phase = "OPEN": Set ses = SAPX_GetSessionForTcode("VL02N")

ModuleMain.AssertRunIdentity
phaseStarted = Timer
OpenDelivery ses, ob
ModuleLog.WriteLog "SAP_OPEN_SEC", Format$(SAPX_Elapsed(phaseStarted), "0.00"), ob

ModuleMain.AssertRunIdentity

' Rule 3: Storage Location Config Gate
sloc = Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R6").Value2))
If sloc = "" Then
 ModuleLog.WriteLog "STOP_MISSING_SLOC_CONFIG", "Thieu cau hinh SLoc tai INPUT!R6", ob
 LastApplyState = "UNSAVED"
 Err.Raise vbObjectError + 4210, , "STOP_MISSING_SLOC_CONFIG: Thieu cau hinh SLoc tai INPUT!R6 tren OB " & ob
End If

Set itemMap = CreateObject("Scripting.Dictionary")
For Each a In allocations
 key = CStr(a(2))
 If Not itemMap.exists(key) Then
  Set batches = CreateObject("Scripting.Dictionary")
  itemMap.Add key, Array(CStr(a(3)), 0#, batches, CDbl(a(9)), CStr(a(10)))
 End If
 planned = itemMap(key): planned(1) = CDbl(planned(1)) + CDbl(a(6)): Set batches = planned(2)
 If Not batches.exists(CStr(a(4))) Then batches(CStr(a(4))) = 0#
 batches(CStr(a(4))) = CDbl(batches(CStr(a(4)))) + CDbl(a(6)): itemMap(key) = planned
Next a

If allocations.count > 0 Then
 firstAlloc = allocations(1)
 targetMat = CStr(firstAlloc(3))
 targetSTO = CStr(firstAlloc(7))
End If

' Rule 1: Initial Snapshot & Open SLoc Check Gate
phase = "OPEN_SLOC_CHECK"
Set mOverviewIndex = OverviewSnapshot(ses)
Set mUnplannedOriginalSnapshot = mOverviewIndex

initialItemCount = mOverviewIndex.count
initialSlocCount = 0
For Each itmKey In mOverviewIndex.keys
 itmData = mOverviewIndex(itmKey)
 If Trim$(CStr(itmData(4))) <> "" Or Trim$(CStr(itmData(1))) <> "" Then initialSlocCount = initialSlocCount + 1
Next itmKey

ModuleLog.WriteLog "OPEN_SLOC_CHECK", "Items=" & initialItemCount & "; InitialSloc=" & initialSlocCount & "; OB=" & ob & "; Mat=" & targetMat & "; STO=" & targetSTO, ob

If initialSlocCount > 0 Then
 ModuleLog.WriteLog "SKIP_OB_ALREADY_PROCESSED", "OB " & ob & " da co SLoc ban dau (count=" & initialSlocCount & "); mutations=0; saves=0", ob
 ses.findById("wnd[0]").sendVKey 12
 SAPX_WaitIdle ses, 15
 LastApplyState = "SKIPPED_ALREADY_PROCESSED"
 LastConflictEligible = True
 LastConflictKind = "ALREADY_PROCESSED"
 result = "SKIP_OB_ALREADY_PROCESSED: OB " & ob & " da co SLoc ban dau (count=" & initialSlocCount & "); mutations=0; saves=0"
 Err.Raise vbObjectError + 4218, , result
End If

Set validatedSplits = CreateObject("Scripting.Dictionary")
phase = "PREFLIGHT"
start = Timer

For Each item In itemMap.keys
 planned = itemMap(item)
 key = CStr(CLng(val(CStr(item))))
 If Not mOverviewIndex.exists(key) Then Err.Raise vbObjectError + 4210, , "Khong tim thay delivery item " & item
 preflightData = mOverviewIndex(key)
 If CStr(preflightData(0)) <> CStr(planned(0)) Then Err.Raise vbObjectError + 4210, , "Material live khong khop item " & item
 If CStr(preflightData(1)) <> "" Then Err.Raise vbObjectError + 4210, , "Item " & item & " da co batch. Tai lai ke hoach."
 If CStr(preflightData(4)) <> CStr(planned(4)) Then Err.Raise vbObjectError + 4210, , "SLoc live khac snapshot da kiem tra. Tai lai ke hoach."
 AssertSnapshotQty CDbl(preflightData(2)), CDbl(planned(3)), CDbl(planned(1)), CStr(item)
 Set batches = planned(2)
 If batches.count > 1 Then
  Set btn = SAPX_FindByTokens(ses, "btnBT_CHSP_T", False)
  If btn Is Nothing Then Err.Raise vbObjectError + 4210, , "Thieu Batch Split button."
 End If
Next item

ModuleMain.AssertRunIdentity
ModuleLog.WriteLog "ITEM_PREFLIGHT", "Items=" & mOverviewIndex.count & "; sec=" & Format$(SAPX_Elapsed(start), "0.00"), ob

' Rule 2: Mandatory Deletion of Unplanned Items
phase = "DELETE_OTHER_ITEMS"
RemoveUnplannedItems ses, ob, itemMap, editStarted

ModuleMain.AssertRunIdentity

' Rule 3: Field Editing and SLoc Filling
For Each item In itemMap.keys
 planned = itemMap(item): Set batches = planned(2): qty = CDbl(planned(1))
 row = LocateItem(ses, CStr(item), tbl)
 Set f = TableField(tbl, "LIPS-CHARG", row)
 If Trim$(CStr(f.text)) <> "" Then Err.Raise vbObjectError + 4210, , "Item da co batch truoc edit."
 Set f = TableField(tbl, "LIPS-LGORT", row)
 If Trim$(CStr(f.text)) <> CStr(planned(4)) Then Err.Raise vbObjectError + 4210, , "SLoc live khac snapshot truoc edit."
 Set f = TableField(tbl, "LIPSD-G_LFIMG", row)
 AssertSnapshotQty ModuleHelper.CleanDouble(f.text), CDbl(planned(3)), qty, CStr(item)

 editStarted = True: phase = "EDIT"
 SetTableText tbl, "LIPSD-G_LFIMG", row, ModuleHelper.FormatSAPQty(qty)
 SetTableText tbl, "LIPS-LGORT", row, sloc

 If batches.count = 1 Then
  For Each b In batches.keys
   SetTableText tbl, "LIPS-CHARG", row, CStr(b)
   ModuleLog.WriteLog "SLOC_FILL", "Item=" & item & "; ExpectedSloc=" & sloc & "; ActualSloc=" & sloc & "; Batch=" & b & "; Qty=" & qty, ob
  Next b
  ses.findById("wnd[0]").sendVKey 0
  If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4211, , "Validate Direct timeout."
  AssertNoPopup ses
 Else
  ClearOverviewSelection ses
  row = LocateItem(ses, CStr(item), tbl)
  tbl.GetAbsoluteRow(CLng(tbl.VerticalScrollbar.Position) + row).selected = True
  Set f = TableField(tbl, "LIPS-POSNR", row): f.SetFocus
  splitStarted = Timer
  Set btn = SAPX_FindByTokens(ses, "btnBT_CHSP_T", False): btn.Press
  If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4212, , "Batch Split timeout."
  AssertNoPopup ses
  Set splitTable = SAPX_FindByTokens(ses, "tblSAPMV50ATC_LIPS_CHND", False)
  If splitTable Is Nothing Then Err.Raise vbObjectError + 4212, , "Khong co Batch Split table."
  AssertSplitEmpty ses, batches.count
  Set splitTable = Nothing: splitRow = 0
  For Each b In batches.keys
   Set splitTable = SplitTableAt(ses, splitRow, rr, splitTable)
   SetTableText splitTable, "LIPS-CHARG", rr, CStr(b)
   SetTableText splitTable, "LIPS-LGORT", rr, sloc
   SetTableText splitTable, "LIPS-LFIMG", rr, ModuleHelper.FormatSAPQty(CDbl(batches(b)))
   ModuleLog.WriteLog "SLOC_FILL", "SplitItem=" & item & "; ChildRow=" & splitRow & "; ExpectedSloc=" & sloc & "; ActualSloc=" & sloc & "; Batch=" & b & "; Qty=" & CDbl(batches(b)), ob
   splitRow = splitRow + 1
  Next b
  ses.findById("wnd[0]").sendVKey 0
  If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4212, , "Validate Batch Split timeout."
  AssertNoPopup ses
  Set splitTable = SAPX_FindByTokens(ses, "tblSAPMV50ATC_LIPS_CHND", False)
  If splitTable Is Nothing Then Err.Raise vbObjectError + 4212, , "Mat Batch Split table sau Enter."
  ValidateSplitRows ses, batches, sloc, qty, ob
  ses.findById("wnd[0]").sendVKey 3
  If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4212, , "Back Overview timeout."
  Set mOverviewIndex = Nothing
  validatedSplits(CStr(item)) = True
  ModuleLog.WriteLog "SPLIT_VALIDATE", "sec=" & Format$(SAPX_Elapsed(splitStarted), "0.00"), ob
  AssertNoPopup ses
 End If
Next item

AssertNoPopup ses

' Rule 4: Pre-SAVE Gate
phase = "PRESAVE": phaseStarted = Timer
ses.findById("wnd[0]").sendVKey 0
If Not SAPX_WaitIdle(ses, 30) Then
 ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Pre-save timeout", ob
 Err.Raise vbObjectError + 4213, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Pre-save timeout."
End If
AssertNoPopup ses: oldStatus = SAPX_StatusText(ses)
ModuleMain.AssertRunIdentity

VerifyDeliveryPlan ses, itemMap, sloc, validatedSplits, ob
ModuleLog.WriteLog "SAP_FINAL_VERIFY_SEC", Format$(SAPX_Elapsed(phaseStarted), "0.00"), ob

ModuleMain.AssertRunIdentity

' Rule 4: Single Save Execution
phase = "SAVE": phaseStarted = Timer
ses.findById("wnd[0]").sendVKey 11
If Not SAPX_WaitIdle(ses, 60) Then
 ModuleLog.WriteLog "SAVE_RESULT", "State=UNKNOWN; MsgId=TIMEOUT; MsgNo=000; Text=UNKNOWN Save timeout OB " & ob, ob
 Err.Raise vbObjectError + 4213, , "UNKNOWN Save timeout OB " & ob
End If
AssertNoPopup ses

msgId = SAPX_MessageId(ses)
msgNo = SAPX_MessageNumber(ses)
statusText = SAPX_StatusText(ses)
statusType = SAPX_StatusType(ses)

isSaved = (statusType = "S" And IsSaveReceipt(statusText, ob, ses))

If isSaved Then
  ModuleLog.WriteLog "SAVE_RESULT", "State=ACKNOWLEDGED_UNVERIFIED; MsgId=" & msgId & "; MsgNo=" & msgNo & "; Text=" & statusText & "; OB=" & ob & "; sec=" & Format$(SAPX_Elapsed(phaseStarted), "0.00"), ob
 oldStatus = statusText
 saveAcknowledged = True: mSavedForReuse = ob
 ModuleLog.WriteLog "SAP_SAVE_SEC", Format$(SAPX_Elapsed(phaseStarted), "0.00"), ob
 ModuleLog.WriteLog "SAVE_ACK", oldStatus, ob
Else
 ModuleLog.WriteLog "SAVE_RESULT", "State=UNKNOWN; MsgId=" & msgId & "; MsgNo=" & msgNo & "; Text=" & statusText & "; OB=" & ob & "; sec=" & Format$(SAPX_Elapsed(phaseStarted), "0.00"), ob
 LastApplyState = "UNKNOWN"
 Err.Raise vbObjectError + 4213, , "UNKNOWN Save: " & statusText
End If

' Rule 4: Post-SAVE Verify (Fast Save Mode with Ambiguity Fallback Gate)
phase = "POSTSAVE_VERIFY": phaseStarted = Timer
fastSaveEnabled = (UCase$(Trim$(CStr(ModuleHelper.ToolConfig("FAST_SAVE_MODE", "TRUE")))) <> "FALSE")
requiresAudit = (UCase$(Trim$(CStr(ModuleHelper.ToolConfig("STRICT_AUDIT", "FALSE")))) = "TRUE")
isCleanReceipt = (statusType = "S" And IsSaveReceipt(statusText, ob, ses))

If fastSaveEnabled And isCleanReceipt And (Not requiresAudit) Then
 postSavePass = True
 ModuleLog.WriteLog "POSTSAVE_VERIFY", "State=DONE; Mode=FAST_SAVE_MODE; MsgId=" & msgId & "; MsgNo=" & msgNo & "; OB=" & ob & "; VerifiedItems=" & itemMap.count & "; sec=" & Format$(SAPX_Elapsed(phaseStarted), "0.00"), ob
Else
 If requiresAudit Then
  ModuleLog.WriteLog "POSTSAVE_VERIFY", "State=AUDIT_TRIGGERED; Reason=STRICT_AUDIT_ENABLED; OB=" & ob, ob
 Else
  ModuleLog.WriteLog "POSTSAVE_VERIFY", "State=AUDIT_TRIGGERED; Reason=AMBIGUOUS_RECEIPT; MsgId=" & msgId & "; MsgNo=" & msgNo & "; OB=" & ob, ob
 End If
 postSavePass = PostSaveVerify(ses, ob, itemMap, sloc)
End If

If Not postSavePass Then
 LastApplyState = "SAVED_UNVERIFIED"
 Err.Raise vbObjectError + 4213, , "POSTSAVE_VERIFY FAILED/UNKNOWN tren OB " & ob
End If

LastApplyState = "DONE"
result = "SAVED " & ob & ": " & oldStatus & "; SAP SAVE RECEIPT; POSTSAVE_VERIFY DONE"
ModuleLog.WriteLog "SAP_TIMING_V2", "total=" & Format$(SAPX_Elapsed(applyStarted), "0.00") & "; idle_wait=" & Format$(mWaitSeconds, "0.00"), ob
mMeasureWait = False
Exit Sub

Fail:
failText = Err.Description
mSavedForReuse = "": mMeasureWait = False
ModuleLog.WriteLog "SAP_TIMING_FAIL", "total=" & Format$(SAPX_Elapsed(applyStarted), "0.00") & "; idle_wait=" & Format$(mWaitSeconds, "0.00"), ob

If Not editStarted Then
 If phase = "PREFLIGHT" And (InStr(1, failText, "khong khop item", vbTextCompare) > 0 Or InStr(1, failText, "da co batch", vbTextCompare) > 0 Or InStr(1, failText, "SLoc live khac", vbTextCompare) > 0 Or InStr(1, failText, "Qty live thay doi", vbTextCompare) > 0 Or InStr(1, failText, "Khong tim thay delivery item", vbTextCompare) > 0) Then LastConflictEligible = True
 If phase = "OPEN" And (InStr(1, failText, "does not exist", vbTextCompare) > 0 Or InStr(1, failText, "deleted", vbTextCompare) > 0 Or InStr(1, failText, "locked", vbTextCompare) > 0 Or InStr(1, failText, "being processed", vbTextCompare) > 0) Then LastConflictEligible = True
 If phase = "OPEN_SLOC_CHECK" Or InStr(1, failText, "SKIP_OB_ALREADY_PROCESSED", vbTextCompare) > 0 Then LastConflictEligible = True: LastConflictKind = "ALREADY_PROCESSED"
End If

If LastConflictEligible Then
 If LastConflictKind = "" Then LastConflictKind = "CHANGED"
 If phase = "OPEN" And (InStr(1, failText, "does not exist", vbTextCompare) > 0 Or InStr(1, failText, "deleted", vbTextCompare) > 0) Then LastConflictKind = "MISSING"
 If phase = "OPEN" And (InStr(1, failText, "locked", vbTextCompare) > 0 Or InStr(1, failText, "being processed", vbTextCompare) > 0 Or InStr(1, failText, "WAIT_REMOTE_SYNC", vbTextCompare) > 0) Then LastConflictKind = "WAIT"
End If

If Not editStarted And phase = "OPEN" And InStr(1, failText, "WAIT_REMOTE_SYNC", vbTextCompare) > 0 Then LastConflictEligible = True: LastConflictKind = "WAIT"

If saveAcknowledged Then
 If InStr(1, failText, "POSTSAVE_VERIFY", vbTextCompare) > 0 Then
   LastApplyState = "SAVED_UNVERIFIED"
 Else
  LastApplyState = "SAVED_UNVERIFIED"
 End If
ElseIf editStarted Then
 LastApplyState = "UNKNOWN"
ElseIf LastApplyState = "" Then
 LastApplyState = "UNSAVED"
End If

ModuleLog.WriteLog "APPLY_FAIL", "phase=" & phase & "; state=" & LastApplyState & "; " & failText, ob
result = LastApplyState & " OB " & ob & "; phase=" & phase & ": " & failText
Err.Raise vbObjectError + 4214, , result
End Sub

Private Sub AssertSnapshotQty(ByVal actual As Double, ByVal snapshot As Double, ByVal planned As Double, ByVal item As String)

If actual <> snapshot Or actual < planned Then Err.Raise vbObjectError + 4210, , "Qty live thay doi item " & item & "; snapshot=" & snapshot & "; actual=" & actual & ". Tai lai ke hoach."

End Sub

Private Sub VerifyDeliveryPlan(ByVal ses As Object, ByVal itemMap As Object, ByVal sloc As String, Optional ByVal validatedSplits As Object = Nothing, Optional ByVal ob As String = "")

Dim item As Variant, planned As Variant, batches As Object, tbl As Object, splitTable As Object
Dim row As Long, rr As Long, fld As Object, btn As Object, batch As String, qty As Double, verified As Object
Dim liveSnap As Object, liveItm As Variant, curSloc As String, curBatch As String

' 1. Check UI Overview Snapshot
Set liveSnap = OverviewSnapshot(ses)
If liveSnap Is Nothing Then
 ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Khong doc duoc Overview truoc Save", ob
 Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Overview Is Nothing"
End If

' 2. Check EXTRA_ITEM: Absence of unplanned items
For Each liveItm In liveSnap.keys
 If Not itemMap.exists(CStr(liveItm)) Then
  If CLng(val(liveItm)) < 900000 Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "EXTRA_ITEM: Item ngoai ke hoach " & liveItm & " van ton tai trong UI truoc Save", ob
   Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: EXTRA_ITEM " & liveItm
  End If
 End If
Next liveItm

' 3. Verify all planned shipping items
For Each item In itemMap.keys
 planned = itemMap(item): Set batches = planned(2)
 row = LocateItem(ses, CStr(item), tbl)

 Set fld = TableField(tbl, "LIPS-MATNR", row)
 If Trim$(CStr(fld.text)) <> CStr(planned(0)) Then
  ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Final Material readback khong khop item " & item, ob
  Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Material readback khong khop"
 End If

 ' SLoc
 Set fld = TableField(tbl, "LIPS-LGORT", row)
 curSloc = Trim$(CStr(fld.text))
 If curSloc = "" Then
  ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "MISSING_SLOC: Item " & item & " thieu SLoc truoc Save", ob
  Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: MISSING_SLOC Item " & item
 ElseIf curSloc <> sloc Then
  ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "SLOC_MISMATCH: Item " & item & " SLoc=" & curSloc & "; expected=" & sloc, ob
  Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: SLOC_MISMATCH Item " & item & " SLoc=" & curSloc
 End If

 ' Qty
 Set fld = TableField(tbl, "LIPSD-G_LFIMG", row)
 If ModuleHelper.CleanDouble(fld.text) <> CDbl(planned(1)) Then
  ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "SPLIT_QTY_MISMATCH: Final Qty readback khong khop item " & item, ob
  Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: SPLIT_QTY_MISMATCH Final Qty readback khong khop"
 End If

 ' Batch / Split
 Set fld = TableField(tbl, "LIPS-CHARG", row)
 curBatch = Trim$(CStr(fld.text))
 If batches.count = 1 Then
  If curBatch = "" Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "MISSING_BATCH: Item " & item & " thieu Batch truoc Save", ob
   Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: MISSING_BATCH Item " & item
  ElseIf Not batches.exists(curBatch) Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Final Batch readback khong khop item " & item, ob
   Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Final Batch readback khong khop"
  End If
 Else
   ' Re-read the final split after all edits; an earlier validation is not final evidence.
  ClearOverviewSelection ses
  row = LocateItem(ses, CStr(item), tbl)
  tbl.GetAbsoluteRow(CLng(tbl.VerticalScrollbar.Position) + row).selected = True
  Set fld = TableField(tbl, "LIPS-POSNR", row): fld.SetFocus
  Set btn = SAPX_FindByTokens(ses, "btnBT_CHSP_T", False)
  If btn Is Nothing Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Thieu nut final Batch Split", ob
   Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Thieu nut final Batch Split"
  End If
  btn.Press
  If Not SAPX_WaitIdle(ses, 30) Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Final Batch Split timeout", ob
   Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Final Batch Split timeout"
  End If
  AssertNoPopup ses
  Set splitTable = SAPX_FindByTokens(ses, "tblSAPMV50ATC_LIPS_CHND", False)
  If splitTable Is Nothing Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Thieu final Batch Split table", ob
   Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Thieu final Batch Split table"
  End If
  ValidateSplitRows ses, batches, sloc, CDbl(planned(1)), ob
  ses.findById("wnd[0]").sendVKey 3
  If Not SAPX_WaitIdle(ses, 30) Then
   ModuleLog.WriteLog "PRESAVE_GATE_FAIL", "UNKNOWN_UI_STATE: Final Split Back timeout", ob
   Err.Raise vbObjectError + 4215, , "PRESAVE_GATE_FAIL: UNKNOWN_UI_STATE Final Split Back timeout"
  End If
  Set mOverviewIndex = Nothing
  AssertNoPopup ses
 End If

NextVerifiedItem:
Next item

End Sub

Private Function PostSaveVerify(ByVal ses As Object, ByVal ob As String, ByVal itemMap As Object, ByVal sloc As String) As Boolean

Dim liveSnap As Object, children As Object, item As Variant, planned As Variant, batches As Object
Dim phaseStarted As Double

phaseStarted = Timer
On Error GoTo PostSaveFail

' VL03N must reload persisted state; the Save receipt and VL02N edit buffer are not proof.
OpenDelivery ses, ob, True
Set liveSnap = OverviewSnapshot(ses)
If liveSnap Is Nothing Then Err.Raise vbObjectError + 4217, , "Persisted overview unavailable"
Set children = CreateObject("Scripting.Dictionary")
For Each item In itemMap.keys
 planned = itemMap(item): Set batches = planned(2)
 If batches.count > 1 Then
  ' The split screen is scoped to this parent. Read its persisted child IDs so
  ' equal batch names on two parents cannot accidentally satisfy one another.
  Set children(CStr(item)) = ReadSavedSplitChildren(ses, CStr(item))
 End If
Next item

PostSaveVerify = SavedSnapshotMatchesPlan(liveSnap, itemMap, children, sloc, ob)
If Not PostSaveVerify Then Exit Function

On Error Resume Next
ses.findById("wnd[0]").sendVKey 3
SAPX_WaitIdle ses, 10
On Error GoTo PostSaveFail

ModuleLog.WriteLog "POSTSAVE_VERIFY", "State=VERIFIED; Items=" & itemMap.count & "; Children=" & children.count & "; OB=" & ob & "; sec=" & Format$(SAPX_Elapsed(phaseStarted), "0.00"), ob
Exit Function

PostSaveFail:
ModuleLog.WriteLog "POSTSAVE_VERIFY", "State=UNKNOWN; Error=" & Err.Description & "; OB=" & ob, ob
On Error Resume Next
ses.findById("wnd[0]").sendVKey 12
SAPX_WaitIdle ses, 10
PostSaveVerify = False
End Function

Private Function ReadSavedSplitChildren(ByVal ses As Object, ByVal parentItem As String) As Object
Dim tbl As Object, splitTable As Object, fld As Object, btn As Object, children As Object
Dim row As Long, r As Long, vr As Long, n As Long, childId As String, batch As String, qty As Double
Set children = CreateObject("Scripting.Dictionary")
ClearOverviewSelection ses
row = LocateItem(ses, parentItem, tbl)
tbl.GetAbsoluteRow(CLng(tbl.VerticalScrollbar.Position) + row).selected = True
Set fld = TableField(tbl, "LIPS-POSNR", row): fld.SetFocus
Set btn = SAPX_FindByTokens(ses, "btnBT_CHSP_T", False)
If btn Is Nothing Then Err.Raise vbObjectError + 4217, , "Persisted split view unavailable for " & parentItem
btn.Press
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4217, , "Persisted split timeout for " & parentItem
AssertNoPopup ses
Set splitTable = SAPX_FindByTokens(ses, "tblSAPMV50ATC_LIPS_CHND", False)
If splitTable Is Nothing Then Err.Raise vbObjectError + 4217, , "Persisted split table unavailable for " & parentItem
n = CLng(splitTable.rowCount)
If n < 1 Or n > 10000 Then Err.Raise vbObjectError + 4217, , "Ambiguous persisted split row count for " & parentItem
For r = 0 To n - 1
 Set splitTable = SplitTableAt(ses, r, vr, splitTable)
 Set fld = TableField(splitTable, "LIPS-POSNR", vr): childId = Trim$(CStr(fld.text))
 Set fld = TableField(splitTable, "LIPS-CHARG", vr): batch = Trim$(CStr(fld.text))
 Set fld = TableField(splitTable, "LIPS-LFIMG", vr): qty = ModuleHelper.CleanDouble(fld.text)
 If childId <> "" Or batch <> "" Or qty <> 0 Then
  If childId = "" Or batch = "" Then Err.Raise vbObjectError + 4217, , "Incomplete persisted split child for " & parentItem
  childId = CStr(CLng(val(childId)))
  If CLng(childId) < 900000 Or children.exists(childId) Then Err.Raise vbObjectError + 4217, , "Ambiguous persisted split child ID " & childId
  children.Add childId, batch
 End If
Next r
ses.findById("wnd[0]").sendVKey 3
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4217, , "Persisted split return timeout for " & parentItem
AssertNoPopup ses
Set mOverviewIndex = Nothing
Set ReadSavedSplitChildren = children
End Function

' Pure comparison gate: the snapshot is read after VL03N reopen. Split child IDs
' come from each parent's persisted split view, never from row order or a receipt.
Private Function SavedSnapshotMatchesPlan(ByVal liveSnap As Object, ByVal itemMap As Object, ByVal children As Object, ByVal sloc As String, ByVal ob As String) As Boolean
Dim expected As Object, item As Variant, child As Variant, liveItm As Variant, planned As Variant, rowData As Variant
Dim batches As Object, childRows As Object, observed As Object, expectedBatch As String, actualBatch As String
Dim total As Double, actualQty As Double, actualSloc As String, actualMat As String, reason As String
On Error GoTo CompareFail
If liveSnap Is Nothing Or itemMap Is Nothing Or children Is Nothing Then Err.Raise vbObjectError + 4217, , "Readback unavailable"
Set expected = CreateObject("Scripting.Dictionary")
For Each item In itemMap.keys
 If Not liveSnap.exists(CStr(item)) Then Err.Raise vbObjectError + 4217, , "MISSING_PLANNED_ITEM " & item
 If expected.exists(CStr(item)) Then Err.Raise vbObjectError + 4217, , "Duplicate planned ID " & item
 expected.Add CStr(item), True
 planned = itemMap(item): Set batches = planned(2): rowData = liveSnap(CStr(item))
 actualMat = Trim$(CStr(rowData(0))): actualQty = CDbl(rowData(2)): actualSloc = Trim$(CStr(rowData(4)))
 If actualMat <> CStr(planned(0)) Then Err.Raise vbObjectError + 4217, , "MATERIAL_MISMATCH item=" & item & " expected=" & planned(0) & " actual=" & actualMat
 If actualQty <> CDbl(planned(1)) Then Err.Raise vbObjectError + 4217, , "QTY_MISMATCH item=" & item & " expected=" & planned(1) & " actual=" & actualQty
 If actualSloc <> sloc Then Err.Raise vbObjectError + 4217, , "SLOC_MISMATCH item=" & item & " expected=" & sloc & " actual=" & actualSloc
 actualBatch = Trim$(CStr(rowData(1)))
 If batches.count = 1 Then
  If Not batches.exists(actualBatch) Then Err.Raise vbObjectError + 4217, , "BATCH_MISMATCH item=" & item & " actual=" & actualBatch
  If children.exists(CStr(item)) Then Err.Raise vbObjectError + 4217, , "Unexpected split for item " & item
 Else
  If actualBatch <> "" Then Err.Raise vbObjectError + 4217, , "Unexpected parent Batch item=" & item & " actual=" & actualBatch
  If Not children.exists(CStr(item)) Then Err.Raise vbObjectError + 4217, , "Missing split view for item " & item
  Set childRows = children(CStr(item))
  Set observed = CreateObject("Scripting.Dictionary"): total = 0
  For Each child In childRows.keys
   If expected.exists(CStr(child)) Or Not liveSnap.exists(CStr(child)) Then Err.Raise vbObjectError + 4217, , "Missing/duplicate split child ID " & child
   expected.Add CStr(child), True
   rowData = liveSnap(CStr(child))
   actualMat = Trim$(CStr(rowData(0))): actualBatch = Trim$(CStr(rowData(1)))
   actualQty = CDbl(rowData(2)): actualSloc = Trim$(CStr(rowData(4)))
   expectedBatch = CStr(childRows(child))
   If Not batches.exists(expectedBatch) Or observed.exists(expectedBatch) Then Err.Raise vbObjectError + 4217, , "Unexpected/duplicate split Batch child=" & child & " actual=" & expectedBatch
   If actualMat <> CStr(planned(0)) Then Err.Raise vbObjectError + 4217, , "SPLIT_MATERIAL_MISMATCH child=" & child & " expected=" & planned(0) & " actual=" & actualMat
   If actualBatch <> expectedBatch Then Err.Raise vbObjectError + 4217, , "SPLIT_BATCH_MISMATCH child=" & child & " expected=" & expectedBatch & " actual=" & actualBatch
   If actualQty <> CDbl(batches(expectedBatch)) Then Err.Raise vbObjectError + 4217, , "SPLIT_QTY_MISMATCH child=" & child & " expected=" & batches(expectedBatch) & " actual=" & actualQty
   If actualSloc <> sloc Then Err.Raise vbObjectError + 4217, , "SPLIT_SLOC_MISMATCH child=" & child & " expected=" & sloc & " actual=" & actualSloc
   observed.Add expectedBatch, True: total = total + actualQty
  Next child
  If observed.count <> batches.count Or total <> CDbl(planned(1)) Then Err.Raise vbObjectError + 4217, , "SPLIT_SUM_MISMATCH item=" & item & " expected=" & planned(1) & " actual=" & total
 End If
Next item
For Each liveItm In liveSnap.keys
 If Not expected.exists(CStr(liveItm)) Then Err.Raise vbObjectError + 4217, , "EXTRA_ITEM " & liveItm
Next liveItm
SavedSnapshotMatchesPlan = True
Exit Function
CompareFail:
reason = Err.Description
ModuleLog.WriteLog "POSTSAVE_VERIFY", "State=FAILED; Reason=" & reason & "; OB=" & ob, ob
SavedSnapshotMatchesPlan = False
End Function





Public Sub OB_ClassifyCandidates(ByVal ws As Worksheet)
Dim rows As Variant, bad As Object, i As Long, n As Long, ob As String, reason As String
n = ws.Cells(ws.rows.count, 1).End(xlUp).row
If n < 6 Then Exit Sub
rows = ws.Range("A6:P" & n).Value2
Set bad = CreateObject("Scripting.Dictionary")
For i = 1 To UBound(rows, 1)
 ob = CStr(rows(i, 1)): reason = ""
 If Trim$(CStr(rows(i, 7))) <> "" Then reason = "EXISTING SLOC " & rows(i, 7)
 If Trim$(CStr(rows(i, 5))) <> "" Then reason = reason & "; EXISTING BATCH"
 If CStr(rows(i, 12)) <> "A" Then reason = reason & "; GM=" & rows(i, 12)
 If Not ModuleHelper.IsCompatibleUoM(CStr(rows(i, 13)), ModuleHelper.ToolConfig("EXPECTED_UOM", "UN")) Then reason = reason & "; INVALID UOM: " & CStr(rows(i, 13))
 If val(CStr(rows(i, 14))) <> 0 Then reason = reason & "; SPLIT CHILD"
 If reason <> "" Then
  If Not bad.exists(ob) Then bad(ob) = reason
 ElseIf Not ModuleMain.EDCCandidateAllowed(CStr(rows(i, 3)), CStr(rows(i, 8))) Then
  ws.Cells(i + 5, 10).Value2 = "BLOCK"
  ws.Cells(i + 5, 11).Value2 = "BLOCK_ITEM: EDC STO/MATERIAL SCOPE BLOCK"
 End If
Next i
For i = 1 To UBound(rows, 1)
 ob = CStr(rows(i, 1))
 If bad.exists(ob) Then
  ws.Cells(i + 5, 10).Value2 = "BLOCK"
  ws.Cells(i + 5, 11).Value2 = "BLOCK_OB: " & bad(ob)
 End If
Next i
End Sub




Private Sub POST_ApplyBatchFilter(ByVal ses As Object, ByVal batches As Object)
Const prefix As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/"
Dim ctl As Object, ws As Worksheet, saved As Variant, values() As Variant, keys As Variant, i As Long, savedRows As Long, eno As Long, detail As String
On Error GoTo Fail
Set ctl = SAPX_FindById(ses, prefix & "ctxtP_CHARG-LOW")
If ctl Is Nothing Then
 If batches.count > 0 Then Err.Raise vbObjectError + 2965, , "Thieu o Batch Posting theo posting.vbs"
 ModuleLog.WriteLog "POST_BATCH_FILTER", "NONE; no explicit-only Batch scope": Exit Sub
End If
SAPX_SetValue ctl, ""
Set ctl = SAPX_FindById(ses, prefix & "ctxtP_CHARG-HIGH")
If Not ctl Is Nothing Then SAPX_SetValue ctl, ""
If batches.count = 0 Then
 ModuleLog.WriteLog "POST_BATCH_FILTER", "NONE; AUTO Batch included; exact Material/Batch pairs checked in RAM"
 Exit Sub
End If
keys = batches.keys
If batches.count = 1 Then
 Set ctl = SAPX_FindById(ses, prefix & "ctxtP_CHARG-LOW")
 SAPX_SetValue ctl, CStr(keys(0))
 If SAPX_ControlValue(ctl) <> CStr(keys(0)) Then Err.Raise vbObjectError + 2965, , "Batch Posting readback khong khop"
 ModuleLog.WriteLog "POST_BATCH_FILTER", "SINGLE; Batch=" & CStr(keys(0)) & "; no popup"
 Exit Sub
End If
Set ctl = SAPX_FindById(ses, prefix & "btn%_P_CHARG_%_APP_%-VALU_PUSH")
If ctl Is Nothing Then Err.Raise vbObjectError + 2965, , "Thieu nut Batch Multiple Selection theo posting.vbs"
ctl.Press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 2965, , "Timeout Batch Multiple Selection"
Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[16]")
If ctl Is Nothing Then Err.Raise vbObjectError + 2965, , "Thieu nut xoa Batch list cu"
ctl.Press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 2965, , "Timeout clear Batch list"
Set ws = ThisWorkbook.Worksheets("SEARCH_LIST")
savedRows = Application.Max(205, ws.Cells(ws.rows.count, 1).End(xlUp).row)
saved = ws.Range("A1:A" & savedRows).Value2
ws.Range("A1:A" & savedRows).ClearContents
ReDim values(1 To batches.count, 1 To 1)
For i = 0 To UBound(keys): values(i + 1, 1) = CStr(keys(i)): Next i
ws.Range("A2").Resize(batches.count, 1).Value2 = values
ws.Range("A2").Resize(batches.count, 1).Copy
Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[24]")
If ctl Is Nothing Then Err.Raise vbObjectError + 2965, , "Thieu nut Paste Batch"
ctl.Press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 2965, , "Timeout paste Batch"
Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[8]")
If ctl Is Nothing Then Err.Raise vbObjectError + 2965, , "Thieu nut OK Batch"
ctl.Press
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 2965, , "Timeout apply Batch list"
Application.CutCopyMode = False
ws.Range("A1:A" & savedRows).Value2 = saved
ModuleLog.WriteLog "POST_BATCH_FILTER", "MULTI; batches=" & batches.count & "; applied once before Posting execute"
Exit Sub
Fail:
eno = Err.Number: detail = Err.Description
On Error Resume Next
Application.CutCopyMode = False
If Not ws Is Nothing Then
 If IsArray(saved) Then ws.Range("A1:A" & savedRows).Value2 = saved
End If
On Error GoTo 0
Err.Raise eno, , "Posting Batch filter: " & detail
End Sub

Public Sub EDC4007_CheckSTO(ByVal sto As String, ByVal item As String, ByVal mat As String, ByVal qty As Double)
Dim ses As Object, grid As Object, ids As Collection, ni As Object, nc As Object
Dim cSTO As String, cItem As String, cMat As String, cQty As String, cDelivered As String, cPlant As String, cDelete As String, cUom As String
Dim r As Long, found As Long, remaining As Double, ctl As Object
Set ses = SAPX_GetSessionForTcode("ME2N"): SAPX_Navigate ses, "ME2N"
EDC_SetField ses, "LISTU", "ALV"
EDC_SetField ses, "SELPA-LOW", "AVIS"
EDC_SetField ses, "S_WERKS-LOW", "4007"
EDC_SetField ses, "S_RESWK-LOW", ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")
EDC_SetField ses, "S_MATNR-LOW", mat
Set ctl = SAPX_FindByTokens(ses, "S_MATNR-HIGH", True): If Not ctl Is Nothing Then ctl.text = ""
Set ctl = SAPX_FindByTokens(ses, "S_WERKS-HIGH", True): If Not ctl Is Nothing Then ctl.text = ""
Set ctl = SAPX_FindByTokens(ses, "S_RESWK-HIGH", True): If Not ctl Is Nothing Then ctl.text = ""
ses.findById("wnd[0]").sendVKey 8
If Not SAPX_WaitIdle(ses, 60) Then Err.Raise vbObjectError + 4251, , "ME2N timeout."
Set grid = SAPX_WaitGrid(ses, 10)
If grid Is Nothing Then Err.Raise vbObjectError + 4251, , "ME2N khong co ALV."
SAPX_BuildColumnCatalogFast grid, ids, ni, nc
cSTO = SAPX_ResolveColumnFast("EBELN", "EBELN", True, ids, ni, nc)
cItem = SAPX_ResolveColumnFast("EBELP", "EBELP", True, ids, ni, nc)
cMat = SAPX_ResolveColumnFast("EMATN", "EMATN,MATNR", True, ids, ni, nc)
cDelivered = SAPX_ResolveColumnFast("MGLIEF", "MGLIEF", True, ids, ni, nc)
cPlant = SAPX_ResolveColumnFast("WERKS", "WERKS", True, ids, ni, nc)
cDelete = SAPX_ResolveColumnFast("LOEKZ", "LOEKZ", False, ids, ni, nc)
cUom = SAPX_ResolveColumnFast("", "MEINS,BPRME,VRKME", False, ids, ni, nc)
For r = 0 To grid.rowCount - 1
 If SAPX_GridText(grid, r, cSTO) = sto And val(SAPX_GridText(grid, r, cItem)) = val(item) And SAPX_GridText(grid, r, cMat) = mat Then
  If SAPX_GridText(grid, r, cPlant) <> "4007" Then Err.Raise vbObjectError + 4251, , "ME2N Plant khong phai 4007."
  If cDelete <> "" Then
   If Trim$(SAPX_GridText(grid, r, cDelete)) <> "" Then Err.Raise vbObjectError + 4251, , "STO Item da bi danh dau xoa."
  End If
  If cUom <> "" Then ModuleLog.WriteLog "ME2N_UOM_INFO", SAPX_GridText(grid, r, cUom) & "; diagnostic only; Create UoM verified in VL10D"
  found = found + 1
  remaining = ModuleHelper.CleanDouble(SAPX_GridText(grid, r, cDelivered))
 End If
Next r
If found <> 1 Then Err.Raise vbObjectError + 4251, , "ME2N can dung 1 dong STO/Item/Material hop le; found=" & found
ModuleLog.WriteLog "ME2N_REMAINING_INFO", "MGLIEF=" & remaining & "; only informative, VL10D determines due capacity"
ModuleLog.WriteLog "ME2N_STO_CHECK", "4007; " & sto & "/" & item & "; " & mat & "; remaining=" & remaining & "; requested=" & qty
End Sub

Private Sub EDC_SetField(ByVal ses As Object, ByVal token As String, ByVal value As String)
Dim ctl As Object, prefix As Variant, actual As String, typ As String, fieldId As String
For Each prefix In Array("ctxt", "txt", "cmb")
 Set ctl = SAPX_FindById(ses, "wnd[0]/usr/" & CStr(prefix) & token)
 If Not ctl Is Nothing Then Exit For
Next prefix
If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, token, True)
If ctl Is Nothing Then Err.Raise vbObjectError + 4251, , "Thieu field " & token
SAPX_SetValue ctl, value
On Error Resume Next
typ = UCase$(CStr(ctl.Type)): fieldId = CStr(ctl.id)
If InStr(typ, "COMBOBOX") > 0 Then
 actual = Trim$(CStr(ctl.key))
Else
 actual = Trim$(CStr(ctl.text))
End If
On Error GoTo 0
ModuleLog.WriteLog "EDC_FIELD_READBACK", token & "; id=" & fieldId & "; type=" & typ & "; expected=" & value & "; actual=" & actual
If StrComp(actual, Trim$(value), vbTextCompare) <> 0 Then Err.Raise vbObjectError + 4251, , "Readback sai " & token & "; id=" & fieldId & "; actual=" & actual & "; expected=" & value
End Sub

Public Function EDC4007_CreateVL10D(ByVal sto As String, ByVal item As String, ByVal mat As String, ByVal qty As Double) As String
Dim btn As Object, tooltip As String, resultText As String, resultType As String, popup As Object, summary As Boolean
If mEDCIdentity <> sto & "|" & item & "|" & mat & "|" & qty Then Err.Raise vbObjectError + 4253, , "VL10D identity changed."
If Not SAPX_TcodeIs(mEDCSession, "VL10D") Then Err.Raise vbObjectError + 4253, , "Khong con o VL10D."
AssertNoPopup mEDCSession
EDCValidateCurrentTarget sto, item, mat, qty
If mEDCGrid Is Nothing Then
 mEDCQty.text = CStr(qty): If ModuleHelper.CleanDouble(mEDCQty.text) <> qty Then Err.Raise vbObjectError + 4253, , "Qty readback sai."
 mEDCCheck.selected = True
Else
 mEDCGrid.modifyCell mEDCRow, mEDCQtyCol, CStr(qty)
 If ModuleHelper.CleanDouble(SAPX_GridText(mEDCGrid, mEDCRow, mEDCQtyCol)) <> qty Then Err.Raise vbObjectError + 4253, , "Qty ALV readback sai."
 mEDCGrid.selectedRows = CStr(mEDCRow)
End If
EDC4007_CreateVL10D = EDCPressCreate()
End Function

Private Sub EDCSetGroupSTOFilter(ByVal ses As Object, ByVal stoScope As String)
Dim tokens As Variant, button As Object, ws As Worksheet, saved As Variant, i As Long, values() As Variant, ctl As Object
tokens = Split(stoScope, ";")
Set button = SAPX_FindByTokens(ses, "%_ST_EBELN_%_APP_%-VALU_PUSH", False)
If stoScope = "" Then
 If Not button Is Nothing Then
  button.Press
  If SAPX_WaitIdle(ses, 10) Then
   ses.findById("wnd[1]/tbar[0]/btn[16]").Press
   ses.findById("wnd[1]/tbar[0]/btn[8]").Press
  End If
 End If
 EDC_SetField ses, "ST_EBELN-LOW", ""
 Set ctl = SAPX_FindByTokens(ses, "ST_EBELN-HIGH", True): If Not ctl Is Nothing Then SAPX_SetValue ctl, ""
 Exit Sub
End If
If UBound(tokens) = 0 Then
 If Not button Is Nothing Then
  button.Press
  If SAPX_WaitIdle(ses, 10) Then
   ses.findById("wnd[1]/tbar[0]/btn[16]").Press
   ses.findById("wnd[1]/tbar[0]/btn[8]").Press
  End If
 End If
 EDC_SetField ses, "ST_EBELN-LOW", CStr(tokens(0))
 Set ctl = SAPX_FindByTokens(ses, "ST_EBELN-HIGH", True): If Not ctl Is Nothing Then SAPX_SetValue ctl, ""
 ModuleLog.WriteLog "VL10D_STO_FILTER", "SINGLE=" & tokens(0)
 Exit Sub
End If

' Multiple STOs (2 to 4 STOs)
If button Is Nothing Then Err.Raise vbObjectError + 4261, , "VL10D STO multiple-selection button missing."
button.Press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 4261, , "VL10D STO popup timeout."
ses.findById("wnd[1]/tbar[0]/btn[16]").Press
Set ws = ThisWorkbook.Worksheets("SEARCH_LIST")
saved = ws.Range("B1:B10").Value2
ReDim values(1 To UBound(tokens) + 1, 1 To 1)
For i = 0 To UBound(tokens): values(i + 1, 1) = CStr(tokens(i)): Next i
ws.Range("B2").Resize(UBound(tokens) + 1, 1).Value2 = values
ws.Range("B2").Resize(UBound(tokens) + 1, 1).Copy
Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[24]")
If ctl Is Nothing Then Err.Raise vbObjectError + 4261, , "Thieu nut Paste STO"
ctl.Press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 4261, , "VL10D STO paste timeout."
Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[8]")
If ctl Is Nothing Then Err.Raise vbObjectError + 4261, , "Thieu nut OK STO"
ctl.Press
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4261, , "VL10D STO selection apply timeout."
ws.Range("B1:B10").Value2 = saved: Application.CutCopyMode = False
Set ctl = SAPX_FindByTokens(ses, "ST_EBELN-HIGH", True): If Not ctl Is Nothing Then SAPX_SetValue ctl, ""
ModuleLog.WriteLog "VL10D_STO_FILTER", "MULTI=" & (UBound(tokens) + 1) & "; " & stoScope
End Sub

Private Sub EDC_OpenDue(ByVal mat As String, ByVal sto As String)
Dim ses As Object, ctl As Object, tabControl As Object
Set mEDCSession = SAPX_GetSessionForTcode("VL10D"): Set ses = mEDCSession
mEDCSupplySelected = False
mEDCPlantSelected = False
mEDCSelectedMat = "": mEDCSelectedSTO = ""
SAPX_Navigate ses, "VL10D"
EDC_SetField ses, "ST_LEDAT-LOW", ""
Set ctl = SAPX_FindByTokens(ses, "ST_LEDAT-HIGH", True): If Not ctl Is Nothing Then ctl.text = ""
EDC_SetField ses, "P_LERUL", ""
Set tabControl = SAPX_FindById(ses, "wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5")
If tabControl Is Nothing Then Err.Raise vbObjectError + 4261, , "Thieu tab STO VL10D."
tabControl.Select
EDCSetGroupSTOFilter ses, sto
Set tabControl = SAPX_FindById(ses, "wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB3")
If tabControl Is Nothing Then Err.Raise vbObjectError + 4261, , "Thieu tab Material/Plant VL10D."
tabControl.Select
If mat = "" Then
 If mEDCGroupTargets Is Nothing Then Err.Raise vbObjectError + 4261, , "VL10D Material scope missing; no Create."
 If mEDCGroupTargets.count = 0 Then Err.Raise vbObjectError + 4261, , "VL10D Material scope empty; no Create."
 EDCSetGroupMaterialFilter ses, mEDCGroupTargets
Else
 EDC_SetField ses, "ST_MATNR-LOW", mat
End If
Set ctl = SAPX_FindByTokens(ses, "ST_MATNR-HIGH", True)
If Not ctl Is Nothing Then
 SAPX_SetValue ctl, ""
 If SAPX_ControlValue(ctl) <> "" Then Err.Raise vbObjectError + 4261, , "VL10D ST_MATNR-HIGH not cleared."
End If
EDC_SetField ses, "ST_WERKS-LOW", "4007"
mEDCPlantSelected = True
mEDCSelectedMat = mat: mEDCSelectedSTO = sto
Set ctl = SAPX_FindByTokens(ses, "ST_WERKS-HIGH", True)
If Not ctl Is Nothing Then
 SAPX_SetValue ctl, ""
 If SAPX_ControlValue(ctl) <> "" Then Err.Raise vbObjectError + 4261, , "VL10D ST_WERKS-HIGH not cleared."
End If
Set ctl = SAPX_FindByTokens(ses, "ST_RESWK-LOW", True)
If Not ctl Is Nothing Then
 SAPX_SetValue ctl, ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")
 If SAPX_ControlValue(ctl) <> ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093") Then Err.Raise vbObjectError + 4261, , "Supplying Plant readback sai."
 mEDCSupplySelected = True
 Set ctl = SAPX_FindByTokens(ses, "ST_RESWK-HIGH", True): If Not ctl Is Nothing Then ctl.text = ""
End If
ModuleLog.WriteLog "VL10D_SELECTION", "Material=" & IIf(mat = "", EDCGroupMaterialNames(), mat) & "; Receiving Plant=4007; STO=" & IIf(sto = "", "AUTO DISCOVERY", sto)
ses.findById("wnd[0]").sendVKey 8
If Not SAPX_WaitIdle(ses, 60) Then Err.Raise vbObjectError + 4261, , "VL10D due index timeout."
AssertNoPopup ses
If SAPX_StatusType(ses) = "E" Or SAPX_StatusType(ses) = "A" Then Err.Raise vbObjectError + 4261, , SAPX_StatusText(ses)
End Sub

Private Function EDC4007_FindDueSingle(ByVal mat As String, ByVal sto As String) As Collection
Dim result As Collection, ses As Object, grid As Object, ids As Collection, ni As Object, nc As Object
Dim cSTO As String, cItem As String, cMat As String, cPlant As String, cQty As String, cUom As String, cSource As String
Dim r As Long, a As Variant, qty As Double, seen As Object, key As String, source As String
EDC_OpenDue mat, sto
Set ses = mEDCSession: Set result = New Collection: Set seen = CreateObject("Scripting.Dictionary")
Set mEDCQty = Nothing: Set mEDCCheck = Nothing: Set mEDCGrid = Nothing: mEDCIdentity = ""
If InStr(1, SAPX_StatusText(ses), "no data", vbTextCompare) > 0 Or InStr(1, SAPX_StatusText(ses), "no deliveries", vbTextCompare) > 0 Or InStr(1, SAPX_StatusText(ses), "no entries", vbTextCompare) > 0 Then
 ModuleLog.WriteLog "VL10D_NO_DUE", mat & "; STO=" & sto & "; " & SAPX_StatusText(ses)
 Set EDC4007_FindDueSingle = result
 Exit Function
End If
Set grid = SAPX_FindLargestGrid(ses)
If grid Is Nothing Then
 Set result = EDC_ClassicDue(ses, mat, sto)
Else
 SAPX_BuildColumnCatalogFast grid, ids, ni, nc
 cSTO = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL10D_STO_COL", "EBELN"), "EBELN,VGBEL", True, ids, ni, nc)
 cItem = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL10D_ITEM_COL", "EBELP"), "EBELP,VGPOS", True, ids, ni, nc)
 cMat = SAPX_ResolveColumnFast("MATNR", "MATNR,EMATN", True, ids, ni, nc)
 cPlant = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL10D_PLANT_COL", "WERKS"), "WERKS,RECEIVING PLANT", True, ids, ni, nc)
 cQty = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL10D_QTY_COL", ""), "LFIMG,DLVQTY,DELIVERY QTY", True, ids, ni, nc)
 cUom = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL10D_UOM_COL", ""), "VRKME,MEINS,SALES UNIT", True, ids, ni, nc)
 cSource = SAPX_ResolveColumnFast("RESWK", "RESWK,SUPPLYING PLANT", False, ids, ni, nc)
 mEDCStoCol = cSTO: mEDCItemCol = cItem: mEDCMatCol = cMat: mEDCPlantCol = cPlant: mEDCUomCol = cUom
 grid.selectedRows = ""
 For r = 0 To grid.rowCount - 1
  SAPX_EnsureGridRowVisible grid, r
  If EDCMaterialInScope(SAPX_GridText(grid, r, cMat), mat) And SAPX_GridText(grid, r, cPlant) = "4007" And (sto = "" Or ModuleMain.EDCScopeIncludes(sto, SAPX_GridText(grid, r, cSTO))) Then
   If Not ModuleHelper.IsCompatibleUoM(SAPX_GridText(grid, r, cUom), ModuleHelper.ToolConfig("EXPECTED_UOM", "UN")) Then Err.Raise vbObjectError + 4261, , "VL10D UoM khong khop: " & SAPX_GridText(grid, r, cUom)
   source = ""
   If cSource <> "" Then source = SAPX_GridText(grid, r, cSource)
   If source = "" And mEDCSupplySelected Then source = ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")
   If source <> "" And source <> ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093") Then Err.Raise vbObjectError + 4261, , "VL10D sai Supplying Plant=" & source
   qty = ModuleHelper.CleanDouble(SAPX_GridText(grid, r, cQty))
   If qty > 0 Then
    key = SAPX_GridText(grid, r, cSTO) & "|" & SAPX_GridText(grid, r, cItem)
    If SAPX_GridText(grid, r, cSTO) = "" Or val(SAPX_GridText(grid, r, cItem)) <= 0 Then Err.Raise vbObjectError + 4261, , "VL10D thieu STO/Item tren dong ket qua."
    If seen.exists(key) Then Err.Raise vbObjectError + 4261, , "VL10D nhieu schedule rows cho cung STO/Item; khong tu gop: " & key
    seen(key) = True
    a = Array(SAPX_GridText(grid, r, cSTO), SAPX_GridText(grid, r, cItem), SAPX_GridText(grid, r, cMat), "4007", qty, SAPX_GridText(grid, r, cUom), "", "", r, cQty, source)
    result.Add a
   End If
  End If
 Next r
 Set mEDCGrid = grid
End If
ModuleLog.WriteLog "VL10D_DUE_CANDIDATES", mat & "; Plant=4007; STO scope=" & sto & "; eligible rows=" & result.count
Set mEDCLastDue = result
mEDCLastDueAt = Timer
Set EDC4007_FindDueSingle = result
End Function

Private Sub EDC_CollectClassic(ByVal control As Object, ByVal rowMap As Object, ByVal depth As Long)
Dim controlId As String, controlText As String, component As String
Dim leftBracket As Long, commaPos As Long, xKey As String, yKey As String
Dim rowCells As Object, childList As Object, child As Object, childCount As Long, i As Long
If control Is Nothing Then Exit Sub
If depth > 14 Then Exit Sub
On Error Resume Next
controlId = CStr(control.id)
controlText = Trim$(CStr(control.text))
Err.Clear
On Error GoTo 0
component = mid$(controlId, InStrRev(controlId, "/") + 1)
leftBracket = InStrRev(controlId, "[")
commaPos = InStrRev(controlId, ",")
If Left$(component, 4) = "txt[" Or Left$(component, 4) = "lbl[" Or Left$(component, 5) = "ctxt[" Then
 If leftBracket > 0 And commaPos > leftBracket And Right$(controlId, 1) = "]" Then
  xKey = mid$(controlId, leftBracket + 1, commaPos - leftBracket - 1)
  yKey = mid$(controlId, commaPos + 1, Len(controlId) - commaPos - 1)
  If IsNumeric(xKey) And IsNumeric(yKey) And controlText <> "" Then
   xKey = CStr(CLng(xKey)): yKey = CStr(CLng(yKey))
   If Not rowMap.exists(yKey) Then
    Set rowCells = CreateObject("Scripting.Dictionary")
    rowMap.Add yKey, rowCells
   End If
   Set rowCells = rowMap.item(yKey)
   If rowCells.exists(xKey) Then rowCells.Remove xKey
   rowCells.Add xKey, Array(controlText, controlId)
  End If
 End If
End If
Set childList = Nothing
On Error Resume Next
Set childList = control.children
If Not childList Is Nothing Then childCount = childList.count
Err.Clear
On Error GoTo 0
If childList Is Nothing Then Exit Sub
For i = 0 To childCount - 1
 Set child = Nothing
 Set child = SAPX_CollectionObjectAt(childList, i)
 If Not child Is Nothing Then EDC_CollectClassic child, rowMap, depth + 1
Next i
End Sub

Private Function EDC_ClassicColumn(ByVal rows As Object, ByVal configured As String, ByVal headers As String) As String
Dim y As Variant, x As Variant, rowCells As Object, header As Variant, cell As Variant, found As Object
If configured <> "" Then EDC_ClassicColumn = configured: Exit Function
Set found = CreateObject("Scripting.Dictionary")
For Each y In rows.keys
 Set rowCells = rows(y)
 For Each x In rowCells.keys
  cell = rowCells(x)
  For Each header In Split(headers, "|")
   If NormalizeHeader(CStr(cell(0))) = NormalizeHeader(CStr(header)) Then found(CStr(x)) = True
  Next header
 Next x
Next y
If found.count = 1 Then
 For Each x In found.keys: EDC_ClassicColumn = CStr(x): Next x
End If
End Function

Private Function EDC_ClassicValue(ByVal rowCells As Object, ByVal x As String) As String
Dim a As Variant
If rowCells.exists(x) Then a = rowCells(x): EDC_ClassicValue = Trim$(CStr(a(0)))
End Function

Private Function EDC_ClassicMaxRow(ByVal rowMap As Object) As Long
Dim rowKey As Variant
For Each rowKey In rowMap.keys
 If CLng(rowKey) > EDC_ClassicMaxRow Then EDC_ClassicMaxRow = CLng(rowKey)
Next rowKey
End Function

Private Function EDC_ClassicSTOAt(ByVal rowMap As Object, ByVal targetY As Long, ByVal xSTO As String, ByVal xItem As String, ByVal xMat As String, Optional ByVal previousSTO As String = "") As String
Dim rowNo As Long, rowCells As Object, value As String
If xSTO <> xItem Then
 If rowMap.exists(CStr(targetY)) Then
  Set rowCells = rowMap.item(CStr(targetY))
  EDC_ClassicSTOAt = EDC_ClassicValue(rowCells, xSTO)
 End If
 Exit Function
End If
EDC_ClassicSTOAt = previousSTO
For rowNo = 0 To targetY
 If rowMap.exists(CStr(rowNo)) Then
  Set rowCells = rowMap.item(CStr(rowNo))
  value = EDC_ClassicValue(rowCells, xSTO)
  If EDC_ClassicValue(rowCells, xMat) = "" And Len(value) = 10 And value Like "##########" Then EDC_ClassicSTOAt = value
 End If
Next rowNo
' An exact STO filter is authoritative only when no visible group header exists.
If EDC_ClassicSTOAt = "" Then EDC_ClassicSTOAt = mEDCSelectedSTO
End Function

Private Function EDC_ClassicPlant(ByVal rowCells As Object, ByVal xPlant As String, ByVal mat As String) As String
If xPlant <> "" Then
 EDC_ClassicPlant = EDC_ClassicValue(rowCells, xPlant)
ElseIf mEDCPlantSelected And (mEDCSelectedMat = mat Or (mEDCSelectedMat = "" And Not mEDCGroupTargets Is Nothing)) Then
 EDC_ClassicPlant = "4007"
End If
End Function

Private Function EDC_ClassicDue(ByVal ses As Object, ByVal mat As String, ByVal sto As String) As Collection
Dim rowMap As Object, result As Collection, seen As Object, pageRows As Object, rowCells As Object, a As Variant
Dim xSTO As String, xItem As String, xMat As String, xPlant As String, xQty As String, xUom As String, xOpen As String
Dim ctl As Object, checkbox As Object, usr As Object, dump As String, key As String, qty As Double, openQty As Double
Dim page As Long, pos As Long, nextPos As Long, stepSize As Long, rowNo As Long, maxRow As Long
Dim actualSTO As String, lastSTO As String, itemText As String
Set rowMap = CreateObject("Scripting.Dictionary")
Set result = New Collection: Set seen = CreateObject("Scripting.Dictionary")
Set usr = ses.findById("wnd[0]/usr")
On Error Resume Next
usr.VerticalScrollbar.Position = 0
On Error GoTo 0
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 4261, , "VL10D first page timeout."
EDC_CollectClassic usr, rowMap, 0
xSTO = EDC_ClassicColumn(rowMap, ModuleHelper.ToolConfig("VL10D_CLASSIC_STO_X", ""), "EBELN|Purchasing Document|Purch.Doc.|Purchase Order|Purch. Order|STO|OriginDoc.|Origin Doc.")
xItem = EDC_ClassicColumn(rowMap, ModuleHelper.ToolConfig("VL10D_CLASSIC_ITEM_X", ""), "EBELP|Item|PO Item|Purch. Doc. Item|Orig .Item|Orig. Item|Origin Item")
xMat = EDC_ClassicColumn(rowMap, ModuleHelper.ToolConfig("VL10D_CLASSIC_MAT_X", ""), "MATNR|Material")
xPlant = EDC_ClassicColumn(rowMap, ModuleHelper.ToolConfig("VL10D_CLASSIC_PLANT_X", ""), "WERKS|Plant|Receiving Plant|Rec. Plant")
xUom = EDC_ClassicColumn(rowMap, ModuleHelper.ToolConfig("VL10D_CLASSIC_UOM_X", ""), "VRKME|MEINS|SU|Un|UoM|Unit")
xQty = EDC_ClassicColumn(rowMap, ModuleHelper.ToolConfig("VL10D_CLASSIC_QTY_X", ""), "CumQty|Cum Qty|Delivery Qty")
xOpen = EDC_ClassicColumn(rowMap, ModuleHelper.ToolConfig("VL10D_CLASSIC_OPEN_X", ""), "Open qty|Open Quantity")
Set mEDCClassicCols = CreateObject("Scripting.Dictionary")
mEDCClassicCols("STO") = xSTO: mEDCClassicCols("ITEM") = xItem
mEDCClassicCols("MAT") = xMat: mEDCClassicCols("PLANT") = xPlant
mEDCClassicCols("UOM") = xUom: mEDCClassicCols("OPEN") = xOpen
If xSTO = "" Or xItem = "" Or xMat = "" Or xUom = "" Or xQty = "" Or (xPlant = "" And Not mEDCPlantSelected) Then
 SAPX_DumpControlIds usr, dump, 0
 ModuleLog.WriteLog "VL10D_LAYOUT_REQUIRED", Left$(dump, 29000)
 Err.Raise vbObjectError + 4261, , "VL10D classic headers unresolved; xem LOG."
End If
ModuleLog.WriteLog "VL10D_CLASSIC_MAP", "STO=" & xSTO & "; Item=" & xItem & "; Mat=" & xMat & "; Qty=" & xQty & "; Open=" & xOpen & "; Uom=" & xUom & "; Plant=" & IIf(xPlant = "", "SELECTION_READBACK_4007", xPlant)
For page = 1 To 50
 Set pageRows = CreateObject("Scripting.Dictionary")
 EDC_CollectClassic usr, pageRows, 0
 maxRow = EDC_ClassicMaxRow(pageRows)
 For rowNo = 0 To maxRow
  If pageRows.exists(CStr(rowNo)) Then
   Set rowCells = pageRows.item(CStr(rowNo))
   Set checkbox = SAPX_FindById(ses, "wnd[0]/usr/chk[1," & rowNo & "]")
   If Not checkbox Is Nothing Then checkbox.selected = False
   Set ctl = SAPX_FindById(ses, "wnd[0]/usr/txt[" & xQty & "," & rowNo & "]")
   If ctl Is Nothing Then
    Dim deltaCol As Long
    For deltaCol = -6 To 6
     If deltaCol <> 0 Then
      Set ctl = SAPX_FindById(ses, "wnd[0]/usr/txt[" & (CLng(xQty) + deltaCol) & "," & rowNo & "]")
      If Not ctl Is Nothing Then Exit For
     End If
    Next deltaCol
   End If
   actualSTO = EDC_ClassicSTOAt(pageRows, rowNo, xSTO, xItem, xMat, lastSTO)
   If actualSTO <> "" Then lastSTO = actualSTO
   If Not ctl Is Nothing And Not checkbox Is Nothing Then
    If Trim$(EDC_ClassicValue(rowCells, xMat)) = "" And ModuleHelper.CleanDouble(ctl.text) > 0 Then Err.Raise vbObjectError + 4261, , "VL10D Material blank in classic Due; no Create; verify layout."
     If EDCMaterialInScope(EDC_ClassicValue(rowCells, xMat), mat) And EDC_ClassicPlant(rowCells, xPlant, EDC_ClassicValue(rowCells, xMat)) = "4007" Then
     If Len(actualSTO) <> 10 Or Not actualSTO Like "##########" Then Err.Raise vbObjectError + 4261, , "VL10D classic item missing parent STO."
     If sto = "" Or ModuleMain.EDCScopeIncludes(sto, actualSTO) Then
      itemText = EDC_ClassicValue(rowCells, xItem)
      If val(itemText) <= 0 Or val(itemText) >= 1000000 Then Err.Raise vbObjectError + 4261, , "VL10D classic invalid STO Item."
      
      Dim rawUom As String, normUom As String, expUom As String
      Dim uomCtl As Object, uomCtlId As String
      expUom = ModuleHelper.NormalizeUoM(ModuleHelper.ToolConfig("EXPECTED_UOM", "UN"))
      
      Set uomCtl = SAPX_FindById(ses, "wnd[0]/usr/txt[" & xUom & "," & rowNo & "]")
      If uomCtl Is Nothing Then uomCtlId = "UOM_CELL_MISSING" Else uomCtlId = CStr(uomCtl.id)
      
      rawUom = EDC_ClassicValue(rowCells, xUom)
      normUom = ModuleHelper.NormalizeUoM(rawUom)
      
      ModuleLog.WriteLog "VL10D_UOM_READ", "Material=" & EDC_ClassicValue(rowCells, xMat) & "; STO=" & actualSTO & "/" & itemText & "; raw='" & rawUom & "'; normalized='" & normUom & "'; expected='" & expUom & "'; page=" & page & "; row=" & rowNo & "; x=" & xUom & "; ctl=" & uomCtlId
      
      If rawUom = "" Then
          ModuleLog.WriteLog "VL10D_UOM_EMPTY", "Material=" & EDC_ClassicValue(rowCells, xMat) & "; STO=" & actualSTO & "/" & itemText & "; row=" & rowNo & "; x=" & xUom
          Err.Raise vbObjectError + 4261, , "UOM_EMPTY: O UoM trong tren dong " & rowNo & " (x=" & xUom & "); kiem tra layout VL10D."
      ElseIf Not ModuleHelper.IsCompatibleUoM(normUom, expUom) Then
          ModuleLog.WriteLog "VL10D_UOM_MISMATCH", "Material=" & EDC_ClassicValue(rowCells, xMat) & "; normalized='" & normUom & "'; expected='" & expUom & "'"
          Err.Raise vbObjectError + 4261, , "UOM_MISMATCH: Material " & EDC_ClassicValue(rowCells, xMat) & " co UoM='" & normUom & "' khong khop expected='" & expUom & "'."
      End If
      qty = ModuleHelper.CleanDouble(ctl.text)
      If xOpen <> "" Then
       openQty = ModuleHelper.CleanDouble(EDC_ClassicValue(rowCells, xOpen))
       If openQty > 0 And openQty < qty Then qty = openQty
      End If
      If qty > 0 Then
       key = actualSTO & "|" & CStr(CLng(val(itemText)))
       If Not seen.exists(key) Then
        seen.Add key, qty
        a = Array(actualSTO, itemText, EDC_ClassicValue(rowCells, xMat), "4007", qty, EDC_ClassicValue(rowCells, xUom), CStr(ctl.id), CStr(checkbox.id), page, "", "")
        a(10) = ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")
        result.Add a
        ModuleLog.WriteLog "VL10D_DUE_CANDIDATE", key & "; " & EDC_ClassicValue(rowCells, xMat) & "; Plant=4007; Qty=" & qty & "; page=" & page & "; row=" & rowNo
       ElseIf CDbl(seen(key)) <> qty Then
        Err.Raise vbObjectError + 4261, , "VL10D duplicate STO/Item has conflicting Qty: " & key
       End If
      End If
     End If
    End If
   End If
  End If
 Next rowNo
 pos = 0: nextPos = 0: stepSize = 0
 On Error Resume Next
 pos = CLng(usr.VerticalScrollbar.Position)
 stepSize = CLng(usr.VerticalScrollbar.PageSize)
 If stepSize > 0 Then usr.VerticalScrollbar.Position = pos + stepSize
 nextPos = CLng(usr.VerticalScrollbar.Position)
 On Error GoTo 0
 If nextPos <= pos Then Exit For
 If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 4261, , "VL10D scroll timeout."
Next page
If page > 50 Then Err.Raise vbObjectError + 4261, , "VL10D >50 pages; thu hep STO scope."
Set EDC_ClassicDue = result
End Function

Private Sub EDCValidateCurrentTarget(ByVal sto As String, ByVal item As String, ByVal mat As String, ByVal qty As Double)
Dim rowMap As Object, rowCells As Object, idt As String, commaPos As Long, bracketPos As Long, rowKey As String, actualSTO As String
If mEDCGrid Is Nothing Then
 Set rowMap = CreateObject("Scripting.Dictionary")
 EDC_CollectClassic mEDCSession.findById("wnd[0]/usr"), rowMap, 0
 idt = CStr(mEDCQty.id): commaPos = InStrRev(idt, ","): bracketPos = InStrRev(idt, "]")
 rowKey = mid$(idt, commaPos + 1, bracketPos - commaPos - 1)
 If Not rowMap.exists(rowKey) Then Err.Raise vbObjectError + 4261, , "VL10D selected row changed."
 Set rowCells = rowMap.item(rowKey)
 actualSTO = EDC_ClassicSTOAt(rowMap, CLng(rowKey), CStr(mEDCClassicCols("STO")), CStr(mEDCClassicCols("ITEM")), CStr(mEDCClassicCols("MAT")))
 If actualSTO <> sto Then Err.Raise vbObjectError + 4261, , "VL10D live parent STO changed."
 If val(EDC_ClassicValue(rowCells, CStr(mEDCClassicCols("ITEM")))) <> val(item) Then Err.Raise vbObjectError + 4261, , "VL10D live Item changed."
 If EDC_ClassicValue(rowCells, CStr(mEDCClassicCols("MAT"))) <> mat Then Err.Raise vbObjectError + 4261, , "VL10D live Material changed."
 If EDC_ClassicPlant(rowCells, CStr(mEDCClassicCols("PLANT")), mat) <> "4007" Then Err.Raise vbObjectError + 4261, , "VL10D live Plant changed."
 If Not ModuleHelper.IsCompatibleUoM(EDC_ClassicValue(rowCells, CStr(mEDCClassicCols("UOM"))), ModuleHelper.ToolConfig("EXPECTED_UOM", "UN")) Then Err.Raise vbObjectError + 4261, , "VL10D live UoM changed: " & EDC_ClassicValue(rowCells, CStr(mEDCClassicCols("UOM")))
 If EDC4007_CurrentDueQty() < qty Then Err.Raise vbObjectError + 4261, , "VL10D live due Qty reduced."
Else
 If SAPX_GridText(mEDCGrid, mEDCRow, mEDCStoCol) <> sto Or val(SAPX_GridText(mEDCGrid, mEDCRow, mEDCItemCol)) <> val(item) Or SAPX_GridText(mEDCGrid, mEDCRow, mEDCMatCol) <> mat Or SAPX_GridText(mEDCGrid, mEDCRow, mEDCPlantCol) <> "4007" Or Not ModuleHelper.IsCompatibleUoM(SAPX_GridText(mEDCGrid, mEDCRow, mEDCUomCol), ModuleHelper.ToolConfig("EXPECTED_UOM", "UN")) Then Err.Raise vbObjectError + 4261, , "VL10D live ALV identity/Plant/UoM changed."
 If ModuleHelper.CleanDouble(SAPX_GridText(mEDCGrid, mEDCRow, mEDCQtyCol)) < qty Then Err.Raise vbObjectError + 4261, , "VL10D live ALV due Qty reduced; khong tao."
End If
End Sub

Public Function EDC4007_PrepareVL10D(ByVal sto As String, ByVal item As String, ByVal mat As String, ByVal qty As Double) As String
Dim candidates As Collection, a As Variant, chosen As Variant, found As Long, usr As Object, page As Long, stepSize As Long, result As String
Set candidates = EDC4007_FindDue(mat, sto)
For Each a In candidates
 If CStr(a(0)) = sto And val(CStr(a(1))) = val(item) And CStr(a(2)) = mat And CStr(a(3)) = "4007" And CDbl(a(4)) >= qty Then chosen = a: found = found + 1
Next a
If found <> 1 Then Err.Raise vbObjectError + 4261, , "VL10D due row changed; can dung 1 STO/Item/Material/Plant, found=" & found
' 4007 creates directly from VL10D; ME2N is only for a post-Create duplicate investigation.
If mEDCGrid Is Nothing Then
 Set usr = mEDCSession.findById("wnd[0]/usr")
 On Error Resume Next: usr.VerticalScrollbar.Position = 0: On Error GoTo 0
 For page = 2 To CLng(chosen(8))
  stepSize = CLng(usr.VerticalScrollbar.PageSize): usr.VerticalScrollbar.Position = CLng(usr.VerticalScrollbar.Position) + stepSize
  If Not SAPX_WaitIdle(mEDCSession, 10) Then Err.Raise vbObjectError + 4261, , "VL10D selected-page timeout."
 Next page
 Set mEDCQty = SAPX_FindById(mEDCSession, CStr(chosen(6))): Set mEDCCheck = SAPX_FindById(mEDCSession, CStr(chosen(7)))
 If mEDCQty Is Nothing Or mEDCCheck Is Nothing Then Err.Raise vbObjectError + 4261, , "VL10D selected row missing."
 If ModuleHelper.CleanDouble(mEDCQty.text) < qty Then Err.Raise vbObjectError + 4261, , "VL10D due Qty changed."
Else
 mEDCRow = CLng(chosen(8)): mEDCQtyCol = CStr(chosen(9))
End If
mEDCIdentity = sto & "|" & item & "|" & mat & "|" & qty
EDC4007_PrepareVL10D = "Material " & mat & "; Plant 4007" & vbLf & "STO " & sto & "/" & item & "; Due=" & chosen(4) & "; Create Qty=" & qty
ModuleLog.WriteLog "VL10D_CHECK_READY", EDC4007_PrepareVL10D
End Function

Public Sub EDC4007_DiagnoseME2N(ByVal mat As String, ByVal sto As String)
Dim ses As Object, grid As Object, ids As Collection, ni As Object, nc As Object, r As Long, cSTO As String, cItem As String, cRemain As String, info As String
On Error GoTo Unavailable
Set ses = SAPX_GetSessionForTcode("ME2N"): SAPX_Navigate ses, "ME2N"
EDC_SetField ses, "LISTU", "ALV": EDC_SetField ses, "SELPA-LOW", "AVIS"
EDC_SetField ses, "S_WERKS-LOW", "4007": EDC_SetField ses, "S_RESWK-LOW", ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")
EDC_SetField ses, "S_MATNR-LOW", mat
ses.findById("wnd[0]").sendVKey 8
If Not SAPX_WaitIdle(ses, 60) Then Err.Raise vbObjectError + 4261, , "ME2N diagnostic timeout."
Set grid = SAPX_WaitGrid(ses, 10): If grid Is Nothing Then Err.Raise vbObjectError + 4261, , "ME2N no ALV."
SAPX_BuildColumnCatalogFast grid, ids, ni, nc
cSTO = SAPX_ResolveColumnFast("EBELN", "EBELN", True, ids, ni, nc)
cItem = SAPX_ResolveColumnFast("EBELP", "EBELP", True, ids, ni, nc)
cRemain = SAPX_ResolveColumnFast("MGLIEF", "MGLIEF", True, ids, ni, nc)
For r = 0 To grid.rowCount - 1
 If sto = "" Or SAPX_GridText(grid, r, cSTO) = sto Then info = info & SAPX_GridText(grid, r, cSTO) & "/" & SAPX_GridText(grid, r, cItem) & "; Still-to-deliver=" & SAPX_GridText(grid, r, cRemain) & vbLf
 If Len(info) > 25000 Then Exit For
Next r
ModuleLog.WriteLog "ME2N_DIAGNOSTIC_ONLY", mat & "; " & info & "; not VL10D due capacity"
Exit Sub
Unavailable:
ModuleLog.WriteLog "ME2N_DIAGNOSTIC_UNAVAILABLE", Err.Description
End Sub

Public Function EDC4007_FindDue(ByVal mat As String, ByVal sto As String) As Collection
Dim scope As String
scope = ModuleMain.EDCNormalizeSTOs(sto)
Set EDC4007_FindDue = EDC4007_FindDueSingle(mat, scope)
End Function

Public Function EDC4007_CurrentDueQty() As Double
Dim idt As String, commaPos As Long, bracketPos As Long, rowKey As String, rowMap As Object, rowCells As Object, xOpen As String, openQty As Double
If mEDCGrid Is Nothing Then
 If mEDCQty Is Nothing Then Err.Raise vbObjectError + 4269, , "VL10D due control missing."
 EDC4007_CurrentDueQty = ModuleHelper.CleanDouble(mEDCQty.text)
 If Not mEDCClassicCols Is Nothing Then
  xOpen = CStr(mEDCClassicCols("OPEN"))
  If xOpen <> "" Then
   idt = CStr(mEDCQty.id): commaPos = InStrRev(idt, ","): bracketPos = InStrRev(idt, "]")
   rowKey = mid$(idt, commaPos + 1, bracketPos - commaPos - 1)
   Set rowMap = CreateObject("Scripting.Dictionary")
   EDC_CollectClassic mEDCSession.findById("wnd[0]/usr"), rowMap, 0
   If Not rowMap.exists(rowKey) Then Err.Raise vbObjectError + 4269, , "VL10D Open Qty row missing."
   Set rowCells = rowMap.item(rowKey)
   openQty = ModuleHelper.CleanDouble(EDC_ClassicValue(rowCells, xOpen))
   If openQty < EDC4007_CurrentDueQty Then EDC4007_CurrentDueQty = openQty
  End If
 End If
Else
 EDC4007_CurrentDueQty = ModuleHelper.CleanDouble(SAPX_GridText(mEDCGrid, mEDCRow, mEDCQtyCol))
End If
If EDC4007_CurrentDueQty <= 0 Then Err.Raise vbObjectError + 4269, , "VL10D live due Qty invalid."
End Function

Public Function EDC4007_RemainingDue(ByVal sto As String, ByVal item As String, ByVal mat As String) As Double
Dim due As Collection, candidate As Variant, count As Long
Set due = EDC4007_FindDue(mat, sto)
For Each candidate In due
 If CStr(candidate(0)) = sto And val(CStr(candidate(1))) = val(item) Then
  count = count + 1
  EDC4007_RemainingDue = CDbl(candidate(4))
 End If
Next candidate
If count > 1 Then Err.Raise vbObjectError + 4269, , "VL10D due readback ambiguous."
End Function


Private Function EDCMaterialInScope(ByVal actualMat As String, ByVal selectedMat As String) As Boolean
If Trim$(actualMat) = "" Then Exit Function
If selectedMat <> "" Then
 EDCMaterialInScope = (actualMat = selectedMat)
ElseIf Not mEDCGroupTargets Is Nothing Then
 EDCMaterialInScope = mEDCGroupTargets.exists(actualMat)
End If
End Function

Public Sub EDC4007_PrepareGroup(ByVal sto As String, ByVal plans As Object)
Dim due As Collection, mat As Variant, a As Variant, planned As Variant, targets As Object, found As Object
If plans.count = 0 Or sto = "" Then Err.Raise vbObjectError + 4261, , "Empty Create group."
Set mEDCGroupTargets = plans: Set targets = CreateObject("Scripting.Dictionary"): Set found = CreateObject("Scripting.Dictionary")
' Reopen with the exact group filter: cached Due is not proof that live selection
' still has the same Material/Plant/STO after SAP navigation.
Dim canReuseScreen As Boolean, pMat As Variant, allPresent As Boolean, itemDue As Variant
canReuseScreen = False
If Not mEDCLastDue Is Nothing And Not mEDCSession Is Nothing Then
 If Timer - mEDCLastDueAt < 30 And mEDCSelectedSTO = sto Then
  If SAPX_TcodeIs(mEDCSession, "VL10D") Then
   allPresent = True
   For Each pMat In plans.keys
    Dim matFound As Boolean: matFound = False
    For Each itemDue In mEDCLastDue
     If CStr(itemDue(2)) = CStr(pMat) Then matFound = True: Exit For
    Next itemDue
    If Not matFound Then allPresent = False: Exit For
   Next pMat
   If allPresent Then canReuseScreen = True
  End If
 End If
End If
If canReuseScreen Then
 Set due = mEDCLastDue
 ModuleLog.WriteLog "VL10D_SCREEN_REUSED", "STO=" & sto & "; Materials=" & plans.count & "; reuse active Due screen without reload"
Else
 Set due = EDC4007_FindDueSingle("", sto)
End If
For Each a In due
  If Trim$(CStr(a(2))) = "" Then Err.Raise vbObjectError + 4261, , "VL10D group Due Material blank; no Create."
  If plans.exists(CStr(a(2))) Then
  planned = plans(CStr(a(2)))
  If CStr(a(0)) = sto And val(CStr(a(1))) = val(CStr(planned(2))) And CStr(a(3)) = "4007" And (CStr(a(10)) = "" Or CStr(a(10)) = ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")) Then
   If found.exists(CStr(a(2))) Then Err.Raise vbObjectError + 4261, , "Duplicate due item in group."
   If CDbl(a(4)) < CDbl(planned(3)) Then Err.Raise vbObjectError + 4261, , "Group due insufficient: " & a(2)
   targets.Add CStr(a(2)), a: found(CStr(a(2))) = True
  End If
 End If
Next a
If targets.count <> plans.count Then Err.Raise vbObjectError + 4261, , "Group due missing Material/STO/Item; no Create."
Set mEDCGroupDue = targets: mEDCGroupSTO = sto
' No due row is selected until every requested item passes the group preflight.
EDCStageGroup True
ModuleLog.WriteLog "VL10D_GROUP_PREPARED", "STO=" & sto & "; Materials=" & plans.count & "; all selected Qty read back"
End Sub

Private Sub EDCStageGroup(ByVal writeFields As Boolean)
Dim mat As Variant, a As Variant, planned As Variant, usr As Object, page As Long, stepSize As Long, selectedRows As String
Dim t0 As Double: t0 = Timer

If mEDCGrid Is Nothing Then
 Set usr = mEDCSession.findById("wnd[0]/usr")
 ' Group items by page to avoid redundant scrolling to top for every material
 Dim pageDict As Object, pList As Object, pNum As Variant
 Set pageDict = CreateObject("Scripting.Dictionary")
 For Each mat In mEDCGroupTargets.keys
  a = mEDCGroupDue(mat)
  page = CLng(a(8))
  If Not pageDict.exists(page) Then Set pageDict(page) = New Collection
  pageDict(page).Add mat
 Next mat
 
 ' Process page by page
 For Each pNum In pageDict.keys
  page = CLng(pNum)
  On Error Resume Next
  usr.VerticalScrollbar.Position = 0
  On Error GoTo 0
  For page = 2 To CLng(pNum)
   stepSize = CLng(usr.VerticalScrollbar.PageSize)
   If stepSize <= 0 Then Err.Raise vbObjectError + 4261, , "Cannot position group due page."
   usr.VerticalScrollbar.Position = CLng(usr.VerticalScrollbar.Position) + stepSize
  Next page
  If Not SAPX_WaitIdle(mEDCSession, 10) Then Err.Raise vbObjectError + 4261, , "Group due page timeout."
  
  Set pList = pageDict(pNum)
  For Each mat In pList
   planned = mEDCGroupTargets(mat): a = mEDCGroupDue(mat)
   Set mEDCQty = SAPX_FindById(mEDCSession, CStr(a(6)))
   Set mEDCCheck = SAPX_FindById(mEDCSession, CStr(a(7)))
   If mEDCQty Is Nothing Or mEDCCheck Is Nothing Then Err.Raise vbObjectError + 4261, , "Group row disappeared."
   EDCValidateCurrentTarget mEDCGroupSTO, CStr(planned(2)), CStr(mat), CDbl(planned(3))
   If writeFields Then
    mEDCQty.text = CStr(planned(3))
    mEDCCheck.selected = True
   End If
   If ModuleHelper.CleanDouble(mEDCQty.text) <> CDbl(planned(3)) Or Not CBool(mEDCCheck.selected) Then Err.Raise vbObjectError + 4261, , "Group Qty/checkbox readback mismatch: " & mat
  Next mat
 Next pNum
Else
 For Each mat In mEDCGroupTargets.keys
  planned = mEDCGroupTargets(mat): a = mEDCGroupDue(mat)
  mEDCRow = CLng(a(8)): mEDCQtyCol = CStr(a(9))
  SAPX_EnsureGridRowVisible mEDCGrid, mEDCRow
  EDCValidateCurrentTarget mEDCGroupSTO, CStr(planned(2)), CStr(mat), CDbl(planned(3))
  If writeFields Then mEDCGrid.modifyCell mEDCRow, mEDCQtyCol, CStr(planned(3))
  If ModuleHelper.CleanDouble(SAPX_GridText(mEDCGrid, mEDCRow, mEDCQtyCol)) <> CDbl(planned(3)) Then Err.Raise vbObjectError + 4261, , "Group ALV Qty readback mismatch: " & mat
  If selectedRows <> "" Then selectedRows = selectedRows & ","
  selectedRows = selectedRows & CStr(mEDCRow)
 Next mat
 If writeFields Then mEDCGrid.selectedRows = selectedRows
 If Not EDCGroupSelectionMatches(CStr(mEDCGrid.selectedRows), selectedRows) Then Err.Raise vbObjectError + 4261, , "Group ALV selection readback mismatch."
End If

ModuleLog.WriteLog "VL10D_TIMING", "StageGroup write=" & writeFields & "; materials=" & mEDCGroupTargets.count & "; duration=" & Format$(Timer - t0, "0.00") & "s"
End Sub

Public Function EDC4007_CreateGroup(ByVal sto As String) As String
If mEDCGroupTargets Is Nothing Then Err.Raise vbObjectError + 4261, , "No prepared group."
If sto <> mEDCGroupSTO Then Err.Raise vbObjectError + 4261, , "Group STO changed."
If Not SAPX_TcodeIs(mEDCSession, "VL10D") Then Err.Raise vbObjectError + 4261, , "Not in VL10D before group Create."
AssertNoPopup mEDCSession
ModuleMain.AssertRunIdentity
If UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW")) <> "APPLY" Then Err.Raise vbObjectError + 4261, , "PREVIEW: no group Create."
mEDCGroupSTO = ""
EDC4007_CreateGroup = EDCPressCreate()
End Function

Private Function EDCPressCreate() As String
Dim btn As Object, tooltip As String, resultText As String, resultType As String, popup As Object, summary As Boolean
Set btn = SAPX_FindById(mEDCSession, "wnd[0]/tbar[1]/btn[19]")
If btn Is Nothing Then Err.Raise vbObjectError + 4253, , "Thieu lenh Create Delivery theo VBS."
On Error Resume Next: tooltip = LCase$(CStr(btn.tooltip) & " " & CStr(btn.text)): On Error GoTo 0
If InStr(tooltip, "deliver") = 0 And InStr(tooltip, "liefer") = 0 And InStr(tooltip, "background") = 0 And InStr(tooltip, "hintergrund") = 0 Then Err.Raise vbObjectError + 4253, , "Chua xac minh duoc nut tao delivery: " & tooltip
mEDCIdentity = ""
btn.Press
If Not SAPX_WaitIdle(mEDCSession, 60) Then Err.Raise vbObjectError + 4253, , "Create timeout; khong retry."
resultText = SAPX_StatusText(mEDCSession)
resultType = SAPX_StatusType(mEDCSession)
Set popup = SAPX_FindById(mEDCSession, "wnd[1]/usr/txtMESSTXT1")
If Not popup Is Nothing Then
 If InStr(1, CStr(popup.text), "See log for information about creating deliveries", vbTextCompare) > 0 Then
  resultText = CStr(popup.text)
  mEDCSession.findById("wnd[1]/tbar[0]/btn[0]").Press
  If Not SAPX_WaitIdle(mEDCSession, 10) Then Err.Raise vbObjectError + 4253, , "VL10D summary popup timeout."
 End If
End If
AssertNoPopup mEDCSession
summary = (InStr(1, resultText, "See log for information about creating deliveries", vbTextCompare) > 0)
ModuleLog.WriteLog "VL10D_CREATE_RESULT", "Type=" & resultType & "; Summary=" & summary & "; " & resultText
If (resultType = "E" Or resultType = "A") And Not summary Then Err.Raise vbObjectError + 4253, , resultText
If summary Then ModuleLog.WriteLog "VL10D_CREATE_LOG_SUMMARY", "Create submitted; not a success receipt; verify due/OB: " & resultText
EDCPressCreate = resultText & "; submitted, verify actual OB"
End Function


Public Function EDC4007_FindDueGroup(ByVal sto As String, ByVal products As Object) As Collection
Set mEDCGroupTargets = products
Set mEDCLastDue = EDC4007_FindDueSingle("", sto)
mEDCLastDueAt = Timer
Set EDC4007_FindDueGroup = mEDCLastDue
End Function

Public Function EDC4007_FindDueMultiSTO(ByVal stoScope As String, ByVal products As Object) As Collection
Set mEDCGroupTargets = products
Set mEDCLastDue = EDC4007_FindDueSingle("", stoScope)
mEDCLastDueAt = Timer
Set EDC4007_FindDueMultiSTO = mEDCLastDue
End Function


Private Function EDCGroupSelectionMatches(ByVal actual As String, ByVal expected As String) As Boolean
Dim selected As Object, wanted As Object, token As Variant, bounds As Variant, index As Long, key As Variant
Set selected = CreateObject("Scripting.Dictionary"): Set wanted = CreateObject("Scripting.Dictionary")
On Error GoTo Invalid
For Each token In Split(Replace(actual, " ", ""), ",")
 If InStr(CStr(token), "-") > 0 Then
  bounds = Split(CStr(token), "-")
  If UBound(bounds) <> 1 Or Not IsNumeric(bounds(0)) Or Not IsNumeric(bounds(1)) Then Exit Function
  If CLng(bounds(1)) < CLng(bounds(0)) Or CLng(bounds(1)) - CLng(bounds(0)) > 500 Then Exit Function
  For index = CLng(bounds(0)) To CLng(bounds(1)): selected(CStr(index)) = True: Next index
 Else
  If Not IsNumeric(token) Then Exit Function
  selected(CStr(CLng(token))) = True
 End If
Next token
For Each token In Split(expected, ","): wanted(CStr(CLng(token))) = True: Next token
If selected.count <> wanted.count Then Exit Function
For Each key In wanted.keys
 If Not selected.exists(CStr(key)) Then Exit Function
Next key
EDCGroupSelectionMatches = True
Invalid:
End Function


Private Function EDCProbeReferenceItem(ByVal grid As Object, ByVal rowNo As Long) As String
Dim colId As Variant, value As String
For Each colId In Array("VGPOS", "POSNV", "EBELP")
 value = SAPX_GridText(grid, rowNo, CStr(colId))
 If value <> "" And IsNumeric(value) Then
  If val(value) > 0 And val(value) < 1000000 Then
   EDCProbeReferenceItem = CStr(CLng(val(value)))
   ModuleLog.WriteLog "VL06O_SOURCE_ITEM_COLUMN", "Technical field=" & colId & "; row=" & rowNo & "; value=" & value
   Exit Function
  End If
 End If
Next colId
End Function

Private Function EDCGroupMaterialNames() As String
If Not mEDCGroupTargets Is Nothing Then EDCGroupMaterialNames = Join(mEDCGroupTargets.keys, ";")
End Function

Private Sub EDCSetGroupMaterialFilter(ByVal ses As Object, ByVal products As Object)
Dim keys As Variant, button As Object, ws As Worksheet, saved As Variant, i As Long, values() As Variant, table As Object, field As Object, pasted As Object, val As String, actualPosition As Long, ctl As Object
keys = products.keys
Set button = SAPX_FindByTokens(ses, "%_ST_MATNR_%_APP_%-VALU_PUSH", False)
If products.count = 1 Then
  If Trim$(CStr(keys(0))) = "" Then Err.Raise vbObjectError + 4261, , "VL10D group Material blank."
  If Not button Is Nothing Then
  button.Press
  If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 4261, , "VL10D Material selection timeout."
  ses.findById("wnd[1]/tbar[0]/btn[16]").Press
  ses.findById("wnd[1]/tbar[0]/btn[8]").Press
 End If
 EDC_SetField ses, "ST_MATNR-LOW", CStr(keys(0))
 ModuleLog.WriteLog "VL10D_MATERIAL_FILTER", "SINGLE=" & keys(0)
 Exit Sub
End If
If button Is Nothing Then Err.Raise vbObjectError + 4261, , "VL10D Material multiple-selection button missing; no unfiltered Create."
button.Press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 4261, , "VL10D Material popup timeout."
ses.findById("wnd[1]/tbar[0]/btn[16]").Press
Set ws = ThisWorkbook.Worksheets("SEARCH_LIST")
saved = ws.Range("A1:A205").Value2
ReDim values(1 To products.count, 1 To 1)
For i = 0 To UBound(keys): values(i + 1, 1) = CStr(keys(i)): Next i
ws.Range("A2").Resize(products.count, 1).Value2 = values
ws.Range("A2").Resize(products.count, 1).Copy
Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[24]")
If ctl Is Nothing Then Err.Raise vbObjectError + 4261, , "Thieu nut Paste Material"
ctl.Press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 4261, , "VL10D Material paste timeout."
Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[8]")
If ctl Is Nothing Then Err.Raise vbObjectError + 4261, , "Thieu nut OK Material"
ctl.Press
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4261, , "VL10D Material selection apply timeout."
ws.Range("A1:A205").Value2 = saved: Application.CutCopyMode = False
ModuleLog.WriteLog "VL10D_MATERIAL_FILTER", "MULTI=" & products.count & "; " & Join(keys, ";")
End Sub

Public Sub EDCSetReferenceCache(ByVal cacheKey As String, ByVal refArray As Variant)
If mEDCReferenceCache Is Nothing Then Set mEDCReferenceCache = CreateObject("Scripting.Dictionary")
mEDCReferenceCache(cacheKey) = refArray
End Sub

Public Function EDC4007_ReadOBReference(ByVal ob As String, ByVal item As String, ByVal mat As String) As Variant
Dim ses As Object, table As Object, field As Object, rowNo As Long, sto As String, sourceItem As String, receivingPlant As String, qty As Double, cacheKey As String, tabs As Collection, tabRef As Object
cacheKey = ob & "|" & item & "|" & mat
If mEDCReferenceCache Is Nothing Then Set mEDCReferenceCache = CreateObject("Scripting.Dictionary")
' Pending reconciliation always reads VL03N live; an earlier in-memory reference is not proof.
Set ses = SAPX_GetSessionForTcode("VL03N")
On Error GoTo NotVisible
OpenDelivery ses, ob, True
rowNo = LocateItem(ses, item, table)
On Error GoTo 0
Set field = EDCOptionalTableField(table, "LIPS-MATNR", rowNo)
If Not field Is Nothing Then
 If Trim$(CStr(field.text)) <> mat Then Err.Raise vbObjectError + 4269, , "OB_REFERENCE_MISMATCH: Material live changed."
End If
Set field = EDCOptionalTableField(table, "LIPS-LFIMG", rowNo)
If Not field Is Nothing Then
 qty = ModuleHelper.CleanDouble(field.text)
Else
 qty = 0
End If
If qty <= 0 Then Err.Raise vbObjectError + 4269, , "OB_REFERENCE_QTY_INVALID: " & ob
Set field = EDCOptionalTableField(table, "LIPS-UMWRK", rowNo)
If Not field Is Nothing Then receivingPlant = Trim$(CStr(field.text))
If receivingPlant = "" Then Err.Raise vbObjectError + 4269, , "OB_REFERENCE_LAYOUT_REQUIRED: receiving Plant LIPS-UMWRK unavailable; do not infer from VL06O."
Set field = EDCOptionalTableField(table, "LIPS-VGBEL", rowNo): If Not field Is Nothing Then sto = Trim$(CStr(field.text))
Set field = EDCOptionalTableField(table, "LIPS-VGPOS", rowNo): If Not field Is Nothing Then sourceItem = Trim$(CStr(field.text))
If sto = "" Or val(sourceItem) = 0 Then
 Set field = TableField(table, "LIPS-POSNR", rowNo)
 If field Is Nothing Then Err.Raise vbObjectError + 4269, , "OB_REFERENCE_LAYOUT_REQUIRED: Item field missing."
 field.SetFocus: ses.findById("wnd[0]").sendVKey 2
 If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 4269, , "OB reference detail timeout."
 Set field = SAPX_FindByTokens(ses, "LIPS-VGBEL|LIPSD-VGBEL", True): If Not field Is Nothing Then sto = Trim$(CStr(field.text))
 Set field = SAPX_FindByTokens(ses, "LIPS-VGPOS|LIPSD-VGPOS", True): If Not field Is Nothing Then sourceItem = Trim$(CStr(field.text))
 If sto = "" Or val(sourceItem) = 0 Then
  Set tabs = New Collection: EDCReferenceTabs ses.findById("wnd[0]/usr"), tabs, 0
  For Each tabRef In tabs
   tabRef.Select
   If Not SAPX_WaitIdle(ses, 5) Then Err.Raise vbObjectError + 4269, , "Reference tab timeout."
   Set field = SAPX_FindByTokens(ses, "LIPS-VGBEL|LIPSD-VGBEL", True): If Not field Is Nothing Then sto = Trim$(CStr(field.text))
   Set field = SAPX_FindByTokens(ses, "LIPS-VGPOS|LIPSD-VGPOS", True): If Not field Is Nothing Then sourceItem = Trim$(CStr(field.text))
   If sto <> "" And val(sourceItem) > 0 Then Exit For
  Next tabRef
 End If
End If
If Len(sto) <> 10 Or Not sto Like "##########" Or val(sourceItem) <= 0 Then Err.Raise vbObjectError + 4269, , "OB_REFERENCE_LAYOUT_REQUIRED: " & ob & "/" & item & "; VL06O/VL03N source Item unavailable; expose VGPOS in layout. No repeated VL06O read."
mEDCReferenceCache(cacheKey) = Array(sto, CStr(CLng(val(sourceItem))), qty, receivingPlant)
EDC4007_ReadOBReference = mEDCReferenceCache(cacheKey)
Exit Function
NotVisible:
Err.Raise vbObjectError + 4269, , "OB_REFERENCE_NOT_VISIBLE_YET: " & ob & "/" & item & "; " & Err.Description
End Function

Private Sub EDCReferenceTabs(ByVal ctl As Object, ByVal tabs As Collection, ByVal depth As Long)
Dim children As Object, child As Object, text As String, typ As String
If ctl Is Nothing Or depth > 12 Then Exit Sub
On Error Resume Next
typ = CStr(ctl.Type): text = LCase$(CStr(ctl.text))
On Error GoTo 0
If typ = "GuiTab" Then
 If InStr(text, "reference") > 0 Or InStr(text, "predecessor") > 0 Or InStr(text, "origin") > 0 Or InStr(text, "order data") > 0 Then tabs.Add ctl
End If
On Error Resume Next: Set children = ctl.children: On Error GoTo 0
If Not children Is Nothing Then
 For Each child In children: EDCReferenceTabs child, tabs, depth + 1: Next child
End If
End Sub


Private Function EDCOptionalTableField(ByVal table As Object, ByVal token As String, ByVal rowNo As Long) As Object
On Error Resume Next
Set EDCOptionalTableField = TableField(table, token, rowNo)
On Error GoTo 0
End Function













Public Function SAPX_DeleteDraftOB(ByVal ses As Object, ByVal ob As String) As String
Dim fld As Object, btnDel As Object, btnOption1 As Object
Dim sText As String, sType As String
On Error GoTo Failed

SAPX_Navigate ses, "VL02N"
If Not SAPX_WaitIdle(ses, 15) Then SAPX_DeleteDraftOB = "DELETE_FAILED": Exit Function

Set fld = DeliveryNumberField(ses)
If fld Is Nothing Then Set fld = SAPX_FindById(ses, "wnd[0]/usr/ctxtLIKP-VBELN")
If fld Is Nothing Then SAPX_DeleteDraftOB = "DELETE_FAILED": Exit Function

fld.text = ob
ses.findById("wnd[0]").sendVKey 0
If Not SAPX_WaitIdle(ses, 15) Then SAPX_DeleteDraftOB = "DELETE_FAILED": Exit Function

sText = SAPX_StatusText(ses)
sType = SAPX_StatusType(ses)

' Check if delivery does not exist (already deleted)
If InStr(1, sText, "does not exist", vbTextCompare) > 0 Or _
   InStr(1, sText, "nicht vorhanden", vbTextCompare) > 0 Or _
   InStr(1, sText, "khong ton tai", vbTextCompare) > 0 Then
 ModuleLog.WriteLog "DELETE_OB_ALREADY_GONE", "OB=" & ob & " does not exist in SAP: " & sText
 SAPX_DeleteDraftOB = "DELETE_VERIFIED"
 Exit Function
End If

' Check if STO or delivery is locked by another user
If InStr(1, sText, "is currently being processed by", vbTextCompare) > 0 Or _
   InStr(1, sText, "wird bereits von", vbTextCompare) > 0 Then
 If InStr(1, sText, "Purchase order", vbTextCompare) > 0 Or InStr(1, sText, "Bestellung", vbTextCompare) > 0 Then
  ModuleLog.WriteLog "WAIT_STO_LOCK", "STO locked by user for OB=" & ob & ": " & sText
  SAPX_DeleteDraftOB = "WAIT_STO_LOCK"
 Else
  ModuleLog.WriteLog "WAIT_OB_LOCK", "OB=" & ob & " locked by user: " & sText
  SAPX_DeleteDraftOB = "WAIT_OB_LOCK"
 End If
 Exit Function
End If

If sType = "E" Then
 ModuleLog.WriteLog "DELETE_OB_OPEN_ERR", "Cannot open " & ob & ": " & sText
 SAPX_DeleteDraftOB = "DELETE_FAILED"
 Exit Function
End If

' Dismiss any informational non-delete popup if present
If Not SAPX_FindById(ses, "wnd[1]") Is Nothing Then
 Dim popText As String
 popText = CStr(SAPX_FindById(ses, "wnd[1]").text)
 If InStr(1, popText, "Delete", vbTextCompare) = 0 And InStr(1, popText, "Lösch", vbTextCompare) = 0 Then
  ses.findById("wnd[1]").sendVKey 0
  SAPX_WaitIdle ses, 10
 End If
End If

ModuleLog.WriteLog "DELETE_OB_SUBMIT", "Pressing toolbar delete btn[14] per deleteob.vbs for draft OB=" & ob

' Step 1 per deleteob.vbs: Press Delete button on Toolbar 1 (btn[14])
Set btnDel = SAPX_FindById(ses, "wnd[0]/tbar[1]/btn[14]")
If btnDel Is Nothing Then Set btnDel = SAPX_FindById(ses, "wnd[0]/tbar[0]/btn[14]")
If btnDel Is Nothing Then
 ModuleLog.WriteLog "DELETE_OB_NO_BTN14", "Toolbar btn[14] missing in VL02N for OB=" & ob
 SAPX_DeleteDraftOB = "DELETE_FAILED"
 Exit Function
End If

btnDel.Press
If Not SAPX_WaitIdle(ses, 15) Then SAPX_DeleteDraftOB = "DELETE_FAILED": Exit Function

' Step 2 per deleteob.vbs: Confirm deletion in popup (btnSPOP-OPTION1)
If Not SAPX_FindById(ses, "wnd[1]") Is Nothing Then
 Set btnOption1 = SAPX_FindById(ses, "wnd[1]/usr/btnSPOP-OPTION1")
 If btnOption1 Is Nothing Then Set btnOption1 = SAPX_FindById(ses, "wnd[1]/usr/btnSPOP-VAROPTION1")
 If btnOption1 Is Nothing Then Set btnOption1 = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[11]")
 If Not btnOption1 Is Nothing Then
  btnOption1.Press
  ModuleLog.WriteLog "DELETE_OB_CONFIRMED", "Confirmed deletion popup for OB=" & ob
 Else
  ses.findById("wnd[1]").sendVKey 0
 End If
 If Not SAPX_WaitIdle(ses, 30) Then SAPX_DeleteDraftOB = "DELETE_FAILED": Exit Function
End If

sText = SAPX_StatusText(ses)
sType = SAPX_StatusType(ses)
ModuleLog.WriteLog "DELETE_OB_RESULT", "OB=" & ob & "; type=" & sType & "; text=" & sText

' Step 3: Hau kiem live bang cach mo lai tra so OB trong VL02N
SAPX_Navigate ses, "VL02N"
SAPX_WaitIdle ses, 10
Set fld = DeliveryNumberField(ses)
If fld Is Nothing Then Set fld = SAPX_FindById(ses, "wnd[0]/usr/ctxtLIKP-VBELN")
If Not fld Is Nothing Then
 fld.text = ob
 ses.findById("wnd[0]").sendVKey 0
 SAPX_WaitIdle ses, 10
 sText = SAPX_StatusText(ses)
 sType = SAPX_StatusType(ses)
 If InStr(1, sText, "does not exist", vbTextCompare) > 0 Or _
    InStr(1, sText, "nicht vorhanden", vbTextCompare) > 0 Or _
    InStr(1, sText, "khong ton tai", vbTextCompare) > 0 Or _
    (sType = "E" And InStr(1, sText, ob) > 0) Then
  ModuleLog.WriteLog "DELETE_VERIFIED", "Confirmed OB=" & ob & " no longer exists on SAP"
  SAPX_DeleteDraftOB = "DELETE_VERIFIED"
  Exit Function
 End If
End If

' Fallback based on deletion message text
If InStr(1, sText, "deleted", vbTextCompare) > 0 Or _
   InStr(1, sText, "gelöscht", vbTextCompare) > 0 Or _
   InStr(1, sText, "da xoa", vbTextCompare) > 0 Then
 SAPX_DeleteDraftOB = "DELETE_VERIFIED"
Else
 SAPX_DeleteDraftOB = "DELETE_FAILED"
End If
Exit Function

Failed:
 ModuleLog.WriteLog "DELETE_OB_FAIL", "OB=" & ob & "; " & Err.Description
 SAPX_DeleteDraftOB = "DELETE_FAILED"
End Function

Public Function SAPX_FindAndReleaseDraft3001OB(ByVal targetMat As String, ByVal deficitQty As Double) As String
Dim ses As Object, ctl As Object, grid As Object, dfrom As Date
Dim ids As Collection, ni As Object, nc As Object
Dim cDel As String, cItem As String, cMat As String, cQty As String
Dim cBatch As String, cSloc As String, cGM As String, cParent As String
Dim r As Long, n As Long, mat As String, ob As String, rowBatch As String, rowSloc As String
Dim rowGM As String, rowParent As String, qtyText As String, qty As Double, reason As String
Dim bestOB As String, bestQty As Double, minDiff As Double, diff As Double, candidateCount As Long
Dim ship3001 As String, sploc3093 As String, singleProd As Object, releaseSuccess As Boolean

On Error GoTo Failed
Set ses = SAPX_GetSessionForTcode("VL06O")
If ses Is Nothing Then Exit Function

ship3001 = ModuleHelper.ToolConfig("PLANT_3001_SHIP_TO", "3001")
sploc3093 = ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093")

SAPX_Navigate ses, "VL06O"
VL6_OpenSelection ses

If Not VL6_SetSelectionShipTo(ses, ship3001) Then
 ModuleLog.WriteLog "VL06O_3001_SELECTION_ERR", "Could not set Ship-to=" & ship3001
 Exit Function
End If

Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT_ID", "wnd[0]/usr/ctxtIF_VSTEL-LOW"))
If Not ctl Is Nothing Then SAPX_SetValue ctl, sploc3093
Set ctl = SAPX_FindById(ses, "wnd[0]/usr/ctxtIF_VSTEL-HIGH")
If Not ctl Is Nothing Then ctl.text = ""

dfrom = DateAdd("d", -60, Date)
Set ctl = SAPX_FindById(ses, "wnd[0]/usr/ctxtIT_WADAT-LOW")
If Not ctl Is Nothing Then ctl.text = Format$(dfrom, "dd.mm.yyyy")
Set ctl = SAPX_FindById(ses, "wnd[0]/usr/ctxtIT_WADAT-HIGH")
If Not ctl Is Nothing Then ctl.text = Format$(Date, "dd.mm.yyyy")

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/ctxtIT_WBSTK-LOW")
If Not ctl Is Nothing Then ctl.text = "A"

ses.findById("wnd[0]/tbar[1]/btn[8]").Press
If Not SAPX_WaitIdle(ses, 30) Then Exit Function

If InStr(1, SAPX_StatusText(ses), "no deliveries", vbTextCompare) > 0 Or _
   InStr(1, SAPX_StatusText(ses), "no data", vbTextCompare) > 0 Or _
   InStr(1, SAPX_StatusText(ses), "keine lieferungen", vbTextCompare) > 0 Then
 ModuleLog.WriteLog "VL06O_3001_EMPTY", "No deliveries found for 3001 in 60-day window"
 Exit Function
End If

Set grid = VL6_EnsureItemGrid(ses)
If grid Is Nothing Then Exit Function

SAPX_BuildColumnCatalogFast grid, ids, ni, nc
cDel = SAPX_ResolveColumnFast("VBELN", "VBELN,DELIVERY", True, ids, ni, nc)
cItem = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_ITEM", "POSNR"), "POSNR,DELIVERY ITEM,ITEM", True, ids, ni, nc)
cMat = SAPX_ResolveColumnFast("MATNR", "MATNR,MATERIAL", True, ids, ni, nc)
cQty = SAPX_ResolveColumnFast("LFIMG", "LFIMG,DELIVERY QTY,DLV.QTY", True, ids, ni, nc)
cBatch = SAPX_ResolveColumnFast("S_CHARG", "S_CHARG,CHARG,BATCH", True, ids, ni, nc)
cSloc = SAPX_ResolveColumnFast("LGORT", "LGORT,STORAGE LOCATION", True, ids, ni, nc)
cGM = SAPX_ResolveColumnFast("WBSTA", "WBSTA,WBSTK,GOODS MOVEMENT STATUS", True, ids, ni, nc)
cParent = SAPX_ResolveColumnFast("UECHA", "UECHA,HIGHER LEVEL BATCH ITEM", False, ids, ni, nc)

Set singleProd = CreateObject("Scripting.Dictionary")
singleProd(targetMat) = True
VL6_ApplyALVFilter ses, grid, cMat, "", singleProd
Set grid = VL6_DirectGrid(ses)
If grid Is Nothing Then Exit Function

n = CLng(grid.rowCount)
ModuleLog.WriteLog "VL06O_3001_FILTER", "Rows for Material=" & targetMat & ": " & n

minDiff = 1E+30: bestOB = "": bestQty = 0: candidateCount = 0

For r = 0 To n - 1
 SAPX_EnsureGridRowVisible grid, r
 mat = RequiredCell(grid, r, cMat)
 If StrComp(mat, targetMat, vbTextCompare) = 0 Then
  rowBatch = RequiredCell(grid, r, cBatch)
  rowSloc = RequiredCell(grid, r, cSloc)
  rowGM = RequiredCell(grid, r, cGM)
  rowParent = RequiredCell(grid, r, cParent)
  ob = RequiredCell(grid, r, cDel)
  qtyText = RequiredCell(grid, r, cQty)
  
  If rowBatch = "" And rowSloc = "" And (rowGM = "A" Or rowGM = "") And val(rowParent) = 0 Then
   If ModuleHelper.ParseSAPIntegerQty(qtyText, qty, reason) And qty > 0 Then
    candidateCount = candidateCount + 1
    diff = Abs(qty - deficitQty)
    ModuleLog.WriteLog "DRAFT_3001_CANDIDATE", "OB=" & ob & "; Mat=" & targetMat & "; Qty=" & qty & "; diff=" & diff & "; deficit=" & deficitQty
    If diff < minDiff Then
     minDiff = diff
     bestOB = ob
     bestQty = qty
    End If
   End If
  End If
 End If
Next r

If bestOB <> "" And bestQty > 0 Then
 ModuleLog.WriteLog "DRAFT_3001_SELECTED", "Selected candidate OB=" & bestOB & " (Qty=" & bestQty & "; minDiff=" & minDiff & ") for deficit=" & deficitQty
 Dim delStatus As String
 delStatus = SAPX_DeleteDraftOB(ses, bestOB)
 If delStatus = "DELETE_VERIFIED" Then
  SAPX_FindAndReleaseDraft3001OB = bestOB & "|" & CStr(bestQty)
  ModuleLog.WriteLog "DRAFT_3001_RELEASE_OK", "Released OB=" & bestOB & " (Qty=" & bestQty & ")"
 ElseIf delStatus = "WAIT_STO_LOCK" Then
  SAPX_FindAndReleaseDraft3001OB = "LOCK|" & bestOB
  ModuleLog.WriteLog "DRAFT_3001_STO_LOCKED", "STO locked by another user; cannot release OB=" & bestOB
 ElseIf delStatus = "WAIT_OB_LOCK" Then
  SAPX_FindAndReleaseDraft3001OB = "LOCK|" & bestOB
  ModuleLog.WriteLog "DRAFT_3001_OB_LOCKED", "OB=" & bestOB & " is locked by another user"
 Else
  SAPX_FindAndReleaseDraft3001OB = "FAIL|" & bestOB
  ModuleLog.WriteLog "DRAFT_3001_RELEASE_FAIL", "Failed to delete OB=" & bestOB & "; status=" & delStatus
 End If
Else
 ModuleLog.WriteLog "DRAFT_3001_NONE", "No unallocated/unpicked draft OB found for Material=" & targetMat
End If

SAPX_Navigate ses, "VL06O"
Exit Function
Failed:
 ModuleLog.WriteLog "DRAFT_3001_ERROR", Err.Number & "; " & Err.Description
 SAPX_FindAndReleaseDraft3001OB = ""
End Function

' =====================================================================
' BRANCH 2001 (SHOP XMD) EWM MONITOR (/SCWM/MON) & PRDO FUNCTIONS
' =====================================================================

Public Function MON2001_GetSession() As Object
Set MON2001_GetSession = SAPX_GetSessionForTcode("/SCWM/MON")
End Function

Public Function SafeGridCell(ByVal grid As Object, ByVal row As Long, ByVal col As String) As String
On Error Resume Next
SafeGridCell = Trim$(CStr(grid.GetCellValue(row, col)))
On Error GoTo 0
End Function

Public Sub MON2001_OpenMonitor(ByVal ses As Object, ByVal wh As String)
Dim ctl As Object, pop As Object
If wh = "" Then wh = "3093"
If Not SAPX_TcodeIs(ses, "/SCWM/MON") Then
    SAPX_Navigate ses, "/SCWM/MON"
End If
' Check if Warehouse popup appears (wnd[1])
Set pop = SAPX_FindById(ses, "wnd[1]")
If Not pop Is Nothing Then
    Set ctl = SAPX_FindById(ses, "wnd[1]/usr/ctxt/SCWM/S_MON_LGNUM-LGNUM")
    If ctl Is Nothing Then Set ctl = SAPX_FindById(ses, "wnd[1]/usr/ctxtP_LGNUM")
    If Not ctl Is Nothing Then
        ctl.text = wh
        ses.findById("wnd[1]/tbar[0]/btn[0]").Press
        SAPX_WaitIdle ses, 15
    End If
End If
End Sub

Public Function MON2001_ReadODOAndTasks(ByVal ses As Object, ByVal obList As Object, ByRef outTasks As Collection) As Boolean
Dim tree As Object, odoGrid As Object, whtGrid As Object, ctl As Object, ws As Worksheet
Dim r As Long, count As Long, obKey As Variant, arrValues() As Variant
Dim mat As String, charg As String, vlpla As String, vsolm As Double, nistm As Double, docno As String, docno_h As String
Dim rowCount As Long, i As Long

Set outTasks = New Collection
MON2001_ReadODOAndTasks = False

If Not ses.findById("wnd[1]", False) Is Nothing Then
    ses.findById("wnd[1]").Close
    SAPX_WaitIdle ses, 5
End If

Set tree = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[0]/shell")
If tree Is Nothing Then Err.Raise vbObjectError + 4301, , "Khong tim thay Tree dieu huong /SCWM/MON"

tree.selectedNode = "N000000010"
tree.doubleClickNode "N000000010"
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4301, , "Timeout mo ODO Selection"

If obList.count = 1 Then
    For Each obKey In obList.keys: Exit For: Next obKey
    Set ctl = SAPX_FindById(ses, "wnd[1]/usr/txtS_ERP_H-LOW")
    If Not ctl Is Nothing Then ctl.text = CStr(obKey)
    Set ctl = SAPX_FindById(ses, "wnd[1]/usr/txtS_ERP_H-HIGH")
    If Not ctl Is Nothing Then ctl.text = ""
ElseIf obList.count > 1 Then
    Set ctl = SAPX_FindById(ses, "wnd[1]/usr/btn%_S_ERP_H_%_APP_%-VALU_PUSH")
    If ctl Is Nothing Then Err.Raise vbObjectError + 4301, , "Khong tim thay nut Multiple Selection S_ERP_H"
    ctl.Press
    If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4301, , "Timeout mo Multiple Selection S_ERP_H"
    
    ses.findById("wnd[2]/tbar[0]/btn[16]").Press
    SAPX_WaitIdle ses, 5
    
    Set ws = ThisWorkbook.Worksheets("SEARCH_LIST")
    ws.columns(1).ClearContents
    ws.Cells(1, 1).Value2 = "ERP_OB"
    ReDim arrValues(1 To obList.count, 1 To 1)
    i = 1
    For Each obKey In obList.keys
        arrValues(i, 1) = CStr(obKey)
        i = i + 1
    Next obKey
    ws.Range("A2").Resize(obList.count, 1).Value2 = arrValues
    ws.Range("A2:A" & CStr(obList.count + 1)).Copy
    
    ses.findById("wnd[2]/tbar[0]/btn[24]").Press
    SAPX_WaitIdle ses, 10
    ses.findById("wnd[2]/tbar[0]/btn[8]").Press
    Application.CutCopyMode = False
    SAPX_WaitIdle ses, 10
End If

ses.findById("wnd[1]/tbar[0]/btn[8]").Press
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4301, , "Timeout thuc thi ODO Selection"

Set odoGrid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If odoGrid Is Nothing Then Exit Function
If odoGrid.rowCount = 0 Then
    MON2001_ReadODOAndTasks = True
    Exit Function
End If

odoGrid.SelectAll
odoGrid.pressToolbarButton "N000000096"
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4301, , "Timeout tai Warehouse Tasks N000000096"

Set whtGrid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell")
If whtGrid Is Nothing Then
    MON2001_ReadODOAndTasks = True
    Exit Function
End If

rowCount = whtGrid.rowCount
For r = 0 To rowCount - 1
    SAPX_EnsureGridRowVisible whtGrid, r
    mat = SafeGridCell(whtGrid, r, "MATNR")
    charg = SafeGridCell(whtGrid, r, "CHARG")
    vlpla = SafeGridCell(whtGrid, r, "VLPLA")
    vsolm = val(SafeGridCell(whtGrid, r, "VSOLM"))
    nistm = val(SafeGridCell(whtGrid, r, "NISTM"))
    docno = SafeGridCell(whtGrid, r, "DOCNO")
    docno_h = SafeGridCell(whtGrid, r, "DOCNO_H")
    If docno_h = "" Then docno_h = SafeGridCell(odoGrid, 0, "DOCNO_H")
    
    outTasks.Add Array(r, mat, charg, vlpla, vsolm, nistm, docno, docno_h)
Next r

MON2001_ReadODOAndTasks = True
End Function

Public Function MON2001_ReadStockOverview(ByVal ses As Object, ByVal matList As Object, ByVal owner As String, ByRef outStock As Object) As Boolean
Dim tree As Object, stockGrid As Object, ctl As Object, ws As Worksheet
Dim r As Long, rowCount As Long, matKey As Variant, arrValues() As Variant
Dim mat As String, charg As String, lgtyp As String, quan As Double, absquan As Double, outquan As Double
Dim key As String, stockItem As Variant, i As Long

Set outStock = CreateObject("Scripting.Dictionary")
MON2001_ReadStockOverview = False
If owner = "" Then owner = "3093"

If Not ses.findById("wnd[1]", False) Is Nothing Then
    ses.findById("wnd[1]").Close
    SAPX_WaitIdle ses, 5
End If

Set tree = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[0]/shell")
If tree Is Nothing Then Err.Raise vbObjectError + 4302, , "Khong tim thay Tree dieu huong /SCWM/MON"

tree.selectedNode = "N000000139"
tree.doubleClickNode "N000000139"
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4302, , "Timeout mo Stock Overview Selection"

Set ctl = SAPX_FindById(ses, "wnd[1]/usr/ctxtS_OWNER-LOW")
If Not ctl Is Nothing Then ctl.text = owner

Set ctl = SAPX_FindById(ses, "wnd[1]/usr/ctxtS_BATCH-LOW")
If Not ctl Is Nothing Then ctl.text = ""

If matList.count = 1 Then
    For Each matKey In matList.keys: Exit For: Next matKey
    Set ctl = SAPX_FindById(ses, "wnd[1]/usr/ctxtS_MATNR-LOW")
    If Not ctl Is Nothing Then ctl.text = CStr(matKey)
    Set ctl = SAPX_FindById(ses, "wnd[1]/usr/ctxtS_MATNR-HIGH")
    If Not ctl Is Nothing Then ctl.text = ""
ElseIf matList.count > 1 Then
    Set ctl = SAPX_FindById(ses, "wnd[1]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH")
    If ctl Is Nothing Then Err.Raise vbObjectError + 4302, , "Khong tim thay nut Multiple Selection S_MATNR"
    ctl.Press
    If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4302, , "Timeout mo Multiple Selection S_MATNR"
    
    ses.findById("wnd[2]/tbar[0]/btn[16]").Press
    SAPX_WaitIdle ses, 5
    
    Set ws = ThisWorkbook.Worksheets("SEARCH_LIST")
    ws.columns(1).ClearContents
    ws.Cells(1, 1).Value2 = "MATERIAL"
    ReDim arrValues(1 To matList.count, 1 To 1)
    i = 1
    For Each matKey In matList.keys
        arrValues(i, 1) = CStr(matKey)
        i = i + 1
    Next matKey
    ws.Range("A2").Resize(matList.count, 1).Value2 = arrValues
    ws.Range("A2:A" & CStr(matList.count + 1)).Copy
    
    ses.findById("wnd[2]/tbar[0]/btn[24]").Press
    SAPX_WaitIdle ses, 10
    ses.findById("wnd[2]/tbar[0]/btn[8]").Press
    Application.CutCopyMode = False
    SAPX_WaitIdle ses, 10
End If

ses.findById("wnd[1]/tbar[0]/btn[8]").Press
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 4302, , "Timeout thuc thi Stock Overview Selection"

Set stockGrid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If stockGrid Is Nothing Then Exit Function

rowCount = stockGrid.rowCount
For r = 0 To rowCount - 1
    SAPX_EnsureGridRowVisible stockGrid, r
    mat = SafeGridCell(stockGrid, r, "MATNR")
    charg = SafeGridCell(stockGrid, r, "CHARG")
    lgtyp = SafeGridCell(stockGrid, r, "LGTYP")
    quan = val(SafeGridCell(stockGrid, r, "QUAN"))
    absquan = val(SafeGridCell(stockGrid, r, "ABSQUAN"))
    outquan = val(SafeGridCell(stockGrid, r, "OUT_QUAN"))
    
    key = mat & "|" & charg
    If Not outStock.exists(key) Then
        outStock.Add key, Array(mat, charg, lgtyp, quan, absquan, outquan)
    Else
        stockItem = outStock(key)
        stockItem(3) = CDbl(stockItem(3)) + quan
        stockItem(4) = CDbl(stockItem(4)) + absquan
        stockItem(5) = CDbl(stockItem(5)) + outquan
        outStock(key) = stockItem
    End If
Next r

MON2001_ReadStockOverview = True
End Function

Public Function MON2001_ConfirmZero(ByVal ses As Object, ByVal rowIdx As Long) As Boolean
Dim whtGrid As Object, subscreen As Object, btnFgr As Object
On Error GoTo Fail
MON2001_ConfirmZero = False

Set whtGrid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell")
If whtGrid Is Nothing Then Err.Raise vbObjectError + 4303, , "Khong tim thay bang WHT"

whtGrid.setCurrentCell rowIdx, ""
whtGrid.selectedRows = CStr(rowIdx)
whtGrid.pressToolbarContextButton "METHODS"
whtGrid.selectContextMenuItem "@M00006"
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4303, , "Timeout mo @M00006 Confirm WT"

Set btnFgr = SAPX_FindById(ses, "wnd[0]/usr/subSUB_ODP_1:/SCWM/SAPLUI_TO_CONF:0150/tabsTS_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_1_TAB:/SCWM/SAPLUI_TO_CONF:0155/cntlCC_ODP1_TB_MAT/shellcont/shell")
If Not btnFgr Is Nothing Then btnFgr.pressButton "ODP1_CONF_FGR"

Set subscreen = SAPX_FindById(ses, "wnd[0]/usr/subSUB_ODP_1:/SCWM/SAPLUI_TO_CONF:0150/tabsTS_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_1_TAB:/SCWM/SAPLUI_TO_CONF:0155/subSUB_ODP1_1_CONT:/SCWM/SAPLUI_TO_CONF:0160/cntlCC_ODP1_MAT/shellcont/shell")
If subscreen Is Nothing Then Err.Raise vbObjectError + 4303, , "Khong tim thay subscreen nhap chenh lech Confirm WT"

subscreen.modifyCheckbox 0, "DIF_ALL", True
subscreen.modifyCell 0, "NISTA", "0"
subscreen.modifyCell 0, "EXCCODE", "OVP2"
subscreen.clearSelection

ses.findById("wnd[0]").sendVKey 11
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4303, , "Timeout Save Confirm 0"

ses.findById("wnd[0]").sendVKey 12
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4303, , "Timeout Back sau Confirm 0"

On Error Resume Next
ses.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton "REFRESH"
If Err.Number <> 0 Then
    Err.Clear
    ses.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell").pressToolbarButton "REFRESH"
End If
On Error GoTo 0
SAPX_WaitIdle ses, 10

MON2001_ConfirmZero = True
Exit Function
Fail:
ModuleLog.WriteLog "CONFIRM_0_FAIL", Err.Number & ": " & Err.Description
MON2001_ConfirmZero = False
End Function

Public Function MON2001_CancelTask(ByVal ses As Object, ByVal rowIdx As Long) As Boolean
Dim whtGrid As Object
On Error GoTo Fail
MON2001_CancelTask = False

Set whtGrid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell")
If whtGrid Is Nothing Then Err.Raise vbObjectError + 4304, , "Khong tim thay bang WHT"

whtGrid.setCurrentCell rowIdx, ""
whtGrid.selectedRows = CStr(rowIdx)
whtGrid.pressToolbarContextButton "METHODS"
whtGrid.selectContextMenuItem "@M00008"
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 4304, , "Timeout Cancel WT @M00008"

On Error Resume Next
ses.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton "REFRESH"
If Err.Number <> 0 Then
    Err.Clear
    ses.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell").pressToolbarButton "REFRESH"
End If
On Error GoTo 0
SAPX_WaitIdle ses, 10

MON2001_CancelTask = True
Exit Function
Fail:
ModuleLog.WriteLog "CANCEL_WT_FAIL", Err.Number & ": " & Err.Description
MON2001_CancelTask = False
End Function

Public Function MON2001_ChangeODOAndCreateTask(ByVal ses As Object, ByVal docno As String, ByVal itemRow As Long, ByVal newQty As Double, ByVal newBatch As String) As Boolean
Dim ctl As Object, alv As Object
On Error GoTo Fail
MON2001_ChangeODOAndCreateTask = False

If Not SAPX_TcodeIs(ses, "/SCWM/PRDO") Then
    SAPX_Navigate ses, "/SCWM/PRDO"
End If

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell")
If Not ctl Is Nothing Then ctl.pressButton "OIP_CHANGE"
SAPX_WaitIdle ses, 10

Set alv = SAPX_FindById(ses, "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell")
If Not alv Is Nothing Then
    alv.modifyCell itemRow, "QTY_UI", CStr(newQty)
    alv.currentCellRow = itemRow
    alv.clearSelection
End If

ses.findById("wnd[0]").sendVKey 25
SAPX_WaitIdle ses, 10

Set ctl = SAPX_FindById(ses, "wnd[1]/usr/btnBUTTON_1")
If Not ctl Is Nothing Then ctl.Press: SAPX_WaitIdle ses, 10

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_DEFDLV")
If Not ctl Is Nothing Then ctl.Select

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_DEFDLV/ssubSUB_ODP_DEFDLV:/SCWM/SAPLUI_TODLV:4300/subSUB_ODP_DEFDLV_DATA:/SCWM/SAPLUI_TODLV:4311/ctxt/SCWM/S_ASP_TODLV_OD_DEFDLV-CHARG")
If Not ctl Is Nothing Then ctl.text = newBatch

Set ctl = SAPX_FindById(ses, "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/cntlCC_TB_OIP_DLV/shellcont/shell")
If Not ctl Is Nothing Then ctl.pressButton "OK_OI_DLV_CREATE_AND_SAVE"
SAPX_WaitIdle ses, 15

ses.findById("wnd[0]").sendVKey 12
SAPX_WaitIdle ses, 10

MON2001_ChangeODOAndCreateTask = True
Exit Function
Fail:
ModuleLog.WriteLog "CHANGE_ODO_FAIL", Err.Number & ": " & Err.Description
MON2001_ChangeODOAndCreateTask = False
End Function




