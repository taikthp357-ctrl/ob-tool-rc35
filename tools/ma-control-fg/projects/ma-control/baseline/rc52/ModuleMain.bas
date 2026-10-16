Attribute VB_Name = "ModuleMain"

Option Explicit
Private mAlloc2001 As Collection
Private mChecked2001 As Boolean
Private mTasks2001 As Collection
Private mStock2001 As Object
Private mOBs2001 As Object
Private mNeedSLoc2001 As Object
Private mEDCDuplicateInvestigated As Object
Private mEDCDueCache As Object
Private mEDCOBCheckedAt As Double
Private mEDCPostingSnapshot As String
Private mEDCCreatePlan As Object
Private mEDCCheckNotes As Object
Private mEDCSTOByMat As Object
Private mEDCWaitByMat As Object
Private mEDCDominantSTO As String

Public RunID As String

Private mRunning As Boolean
Private mSavedCalculation As Long
Private mSavedEnableEvents As Boolean
Private mSavedScreenUpdating As Boolean
Private mSettingsSaved As Boolean

Private mSignature As String

Private mSourceSignature As String
Private mPlanCheckedAt As Date
Private mChecked As Boolean

Private mPlanViewSignature As String

Private mPlanPolicy As String

Private mStock As Variant

Private mRemain() As Double

Private mInitialRemain() As Double

Private mStockOrder() As Long

Private mStockRank() As Double

Private mStockTie() As String

Private mOB As Variant

Private mOBCap() As Double

Private mAlloc As Collection

Private mTargets As Object

Private mRequestKeys As Object

Private mRecoveryMaterial As String

Private mExcludedOB As Object

Private mRecoveryCaps As Object

Private mDoneBaseline As Object

Private mLastSkippedOB As String

Private mRetryCounts As Object
Private mWholeNeeds As Object
Private mRecoveryProducts As Object
Private mRecoveryCapMap As Object
Private mRecoveryPrepared As Boolean
Private mPostingCheckedAt As Double
Private mPostBatchPairs As Object
Private mBlockedStockByMat As Object
Private mMissingOBWarnUser As Boolean
Private mPostAutoMaterials As Object

Public Function InputProducts(Optional ByVal allMaterials As Boolean = False) As Object

Dim d As Object, ws As Worksheet, ls As Worksheet, r As Long, n As Long, v As String

Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare

Set ws = ThisWorkbook.Worksheets("INPUT")

Set ls = ThisWorkbook.Worksheets("SEARCH_LIST")

ls.Range("A2:A205").ClearContents

For r = 6 To 205

 v = ModuleHelper.CleanStr(ws.Cells(r, 1).Value2)

 If v <> "" Then
 If allMaterials Then
  d(v) = True
 ElseIf Not mRecoveryProducts Is Nothing Then
  If mRecoveryProducts.exists(v) Then d(v) = True
 ElseIf mRecoveryMaterial = "" Or StrComp(v, mRecoveryMaterial, vbTextCompare) = 0 Then
  d(v) = True
 End If
End If

Next r

EachKeyWrite d, ls

Set InputProducts = d

End Function

Private Sub EachKeyWrite(ByVal d As Object, ByVal ws As Worksheet)

Dim k As Variant, r As Long

r = 2

For Each k In d.keys: ws.Cells(r, 1).Value2 = CStr(k): r = r + 1: Next k

End Sub

Private Function Fingerprint(Optional ByVal withRequestID As Boolean = True) As String

Dim ws As Worksheet, r As Long, c As Long, s As String, t As String, inputValues As Variant

Set ws = ThisWorkbook.Worksheets("INPUT")

inputValues = ws.Range("A6:C205").Value2

For r = 1 To UBound(inputValues, 1)

 For c = 1 To 3

  t = ModuleHelper.CleanStr(inputValues(r, c))

  s = s & Len(t) & ":" & t

 Next c

Next r

For r = 5 To 8

 t = ModuleHelper.CleanStr(ws.Cells(r, 18).Value2): s = s & Len(t) & ":" & t

Next r

 ' Legacy request identity marker; whole-batch policy is enforced separately.
s = s & "FALSE"

If CStr(ws.Range("R5").Value2) = "4007" And Trim$(CStr(ws.Range("R31").Value2)) <> "" Then s = s & "|EDC_STO=" & Trim$(CStr(ws.Range("R31").Value2))

If withRequestID Then s = s & ModuleHelper.ToolConfig("REQUEST_ID", "TEST_001")

Fingerprint = s

End Function

Public Sub OB_InputChanged()

If mRunning Then Exit Sub

mChecked = False

ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = "INPUT MOI - NHAN NUT 1"

OB_UpdateModeButton

End Sub

Private Sub PutConfig(ByVal key As String, ByVal value As String)

Dim ws As Worksheet, r As Long

Set ws = ThisWorkbook.Worksheets("CONFIG")

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If CStr(ws.Cells(r, 1).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub

Next r

ws.Cells(r, 1).Value2 = key: ws.Cells(r, 2).Value2 = value

End Sub

Private Sub EnsureRequestIdentity()

Dim currentID As String, boundID As String, saved As String, inputSig As String, ledger As Worksheet

Dim r As Long, old As String, changed As Boolean, hasHistory As Boolean, candidate As String, seq As Long, exists As Boolean

If Not ModuleHelper.ToolConfigBool("AUTO_REQUEST_ID", True) Then Exit Sub

currentID = ModuleHelper.ToolConfig("REQUEST_ID", "TEST_001")

boundID = ModuleHelper.ToolConfig("REQUEST_BOUND_ID", "")

saved = ModuleHelper.ToolConfig("REQUEST_INPUT_SIGNATURE", ""): inputSig = Fingerprint(False)

Set ledger = ThisWorkbook.Worksheets("LEDGER")

If boundID = currentID And saved <> "" Then

 changed = (saved <> inputSig)

Else

 ' Bootstrap old versions or respect an explicitly changed REQUEST_ID.

 For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

  If Left$(CStr(ledger.Cells(r, 1).Value2), Len(currentID) + 1) = currentID & "|" Then

   hasHistory = True: old = CStr(ledger.Cells(r, 10).Value2)

   If old <> inputSig & currentID Then changed = True

  End If

 Next r

End If

If changed Then

 seq = 0

 Do

  seq = seq + 1: candidate = "REQ_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(seq): exists = False

  For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

   If Left$(CStr(ledger.Cells(r, 1).Value2), Len(candidate) + 1) = candidate & "|" Then exists = True: Exit For

  Next r

  If candidate = currentID Then exists = True

 Loop While exists

 PutConfig "REQUEST_ID", candidate
 PutConfig "VL10D_PENDING_KEY", ""
 PutConfig "VL10D_GROUP_PENDING", ""

 ModuleLog.WriteLog "NEW_REQUEST", "INPUT thay doi; " & currentID & " -> " & candidate

 currentID = candidate

End If

PutConfig "REQUEST_BOUND_ID", currentID

PutConfig "REQUEST_INPUT_SIGNATURE", inputSig

End Sub

Private Function IsPendingState(ByVal state As String) As Boolean

IsPendingState = (state = "UNKNOWN" Or state = "SAVED_UNVERIFIED")

End Function

Private Function HasSavedUnverified(ByVal ob As String) As Boolean

Dim ws As Worksheet, r As Long

Set ws = ThisWorkbook.Worksheets("LEDGER")

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If CStr(ws.Cells(r, 3).Value2) = ob And CStr(ws.Cells(r, 6).Value2) = "SAVED_UNVERIFIED" Then HasSavedUnverified = True: Exit Function

Next r

End Function

Private Function PendingMaterial(ByVal mat As String, ByVal batch As String) As Boolean

Dim ws As Worksheet, r As Long

Set ws = ThisWorkbook.Worksheets("LEDGER")

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If IsPendingState(CStr(ws.Cells(r, 6).Value2)) And CStr(ws.Cells(r, 11).Value2) = mat Then

  If batch = "" Or CStr(ws.Cells(r, 12).Value2) = batch Then PendingMaterial = True: Exit Function

 End If

Next r

End Function

Private Function PendingOB(ByVal ob As String) As Boolean

Dim ws As Worksheet, r As Long

Set ws = ThisWorkbook.Worksheets("LEDGER")

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If CStr(ws.Cells(r, 3).Value2) = ob And IsPendingState(CStr(ws.Cells(r, 6).Value2)) Then PendingOB = True: Exit Function

Next r

End Function

Private Sub ArchivePendingSnapshots()

Dim ledger As Worksheet, plan As Worksheet, cache As Worksheet, r As Long, p As Long, q As Long, suffix As String

Set ledger = ThisWorkbook.Worksheets("LEDGER"): Set plan = ThisWorkbook.Worksheets("BATCH_PLAN"): Set cache = ThisWorkbook.Worksheets("VL06O_DATA")

ledger.Range("K1:P1").Value2 = Array("MATERIAL SNAPSHOT", "BATCH SNAPSHOT", "BIN SNAPSHOT", "ORIGINAL OB QTY", "ORIGINAL SLOC", "DEST SLOC")

For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

 If IsPendingState(CStr(ledger.Cells(r, 6).Value2)) And CStr(ledger.Cells(r, 11).Value2) = "" Then

  For p = 6 To plan.Cells(plan.rows.count, 1).End(xlUp).row

   suffix = "|" & CStr(plan.Cells(p, 3).Value2) & "|" & CStr(plan.Cells(p, 4).Value2) & "|" & CStr(plan.Cells(p, 6).Value2) & "|" & CStr(plan.Cells(p, 7).Value2)

   If Right$(CStr(ledger.Cells(r, 8).Value2), Len(suffix)) = suffix And CDbl(ledger.Cells(r, 5).Value2) = ModuleHelper.CleanDouble(plan.Cells(p, 8).Value2) Then

    For q = 6 To cache.Cells(cache.rows.count, 1).End(xlUp).row

     If CStr(cache.Cells(q, 1).Value2) = CStr(ledger.Cells(r, 3).Value2) And CStr(cache.Cells(q, 2).Value2) = CStr(ledger.Cells(r, 4).Value2) And CStr(cache.Cells(q, 3).Value2) = CStr(plan.Cells(p, 5).Value2) Then

      ledger.Cells(r, 11).Resize(1, 6).Value2 = Array(plan.Cells(p, 5).Value2, plan.Cells(p, 6).Value2, plan.Cells(p, 7).Value2, cache.Cells(q, 4).Value2, cache.Cells(q, 7).Value2, plan.Cells(p, 9).Value2)

      Exit For

     End If

    Next q

    Exit For

   End If

  Next p

  If CStr(ledger.Cells(r, 11).Value2) = "" Then Err.Raise vbObjectError + 4113, , "Khong co snapshot cho UNKNOWN OB " & ledger.Cells(r, 3).Value2 & ". Giu ke hoach cu de kiem tra."

 End If

Next r

End Sub

Private Function HasCurrentPendingPlan() As Boolean

Dim ledger As Worksheet, plan As Worksheet, r As Long, p As Long, suffix As String

Set ledger = ThisWorkbook.Worksheets("LEDGER"): Set plan = ThisWorkbook.Worksheets("BATCH_PLAN")

For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

 If IsPendingState(CStr(ledger.Cells(r, 6).Value2)) And CStr(ledger.Cells(r, 10).Value2) = mSignature Then

  For p = 6 To plan.Cells(plan.rows.count, 1).End(xlUp).row

   suffix = "|" & CStr(plan.Cells(p, 3).Value2) & "|" & CStr(plan.Cells(p, 4).Value2) & "|" & CStr(plan.Cells(p, 6).Value2) & "|" & CStr(plan.Cells(p, 7).Value2)

   If Right$(CStr(ledger.Cells(r, 8).Value2), Len(suffix)) = suffix Then HasCurrentPendingPlan = True: Exit Function

  Next p

 End If

Next r

End Function

Private Sub ResolveSaveReceipts()

Dim ledger As Worksheet, logs As Worksheet, r As Long, j As Long, ob As String, oldRun As String, msg As String, changed As Boolean

Set ledger = ThisWorkbook.Worksheets("LEDGER"): Set logs = ThisWorkbook.Worksheets("LOG")

For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

 If CStr(ledger.Cells(r, 6).Value2) = "SAVED_UNVERIFIED" Then

  ob = CStr(ledger.Cells(r, 3).Value2): oldRun = CStr(ledger.Cells(r, 2).Value2)

  For j = logs.Cells(logs.rows.count, 1).End(xlUp).row To 6 Step -1

   If CStr(logs.Cells(j, 2).Value2) = oldRun And CStr(logs.Cells(j, 3).Value2) = "SAVE_ACK" And CStr(logs.Cells(j, 4).Value2) = ob Then

    msg = CStr(logs.Cells(j, 6).Value2)

    If ModuleSAP.IsSaveReceipt(msg, ob) Then

     ledger.Cells(r, 6).Value2 = "DONE": ledger.Cells(r, 7).Value2 = Now

     ledger.Cells(r, 9).Value2 = "SAP SAVE RECEIPT: " & msg: changed = True

    End If

    Exit For

   End If

  Next j

 End If

Next r

If changed Then PersistCheckpoint

End Sub

Private Sub ReconcileArchivedPending()

Dim ws As Worksheet, obs As Object, ob As Variant, items As Collection, r As Long, dest As String, state As String, failure As String, eno As Long

ResolveSaveReceipts

Set ws = ThisWorkbook.Worksheets("LEDGER"): Set obs = CreateObject("Scripting.Dictionary")

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If IsPendingState(CStr(ws.Cells(r, 6).Value2)) Then obs(CStr(ws.Cells(r, 3).Value2)) = True

Next r

For Each ob In obs.keys

 Set items = New Collection: dest = ""

 For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

  If CStr(ws.Cells(r, 3).Value2) = CStr(ob) And IsPendingState(CStr(ws.Cells(r, 6).Value2)) Then

   If dest = "" Then dest = CStr(ws.Cells(r, 16).Value2)

   If dest = "" Or dest <> CStr(ws.Cells(r, 16).Value2) Then Err.Raise vbObjectError + 4113, , "UNKNOWN snapshot SLoc khong khop."

   items.Add Array(0, CStr(ob), CStr(ws.Cells(r, 4).Value2), CStr(ws.Cells(r, 11).Value2), CStr(ws.Cells(r, 12).Value2), CStr(ws.Cells(r, 13).Value2), CDbl(ws.Cells(r, 5).Value2), "", "", CDbl(ws.Cells(r, 14).Value2), CStr(ws.Cells(r, 15).Value2))

  End If

 Next r

 state = "": failure = ""

 On Error Resume Next

 Err.Clear: state = ModuleSAP.OB_ResolveUnknown(CStr(ob), items, dest)

 eno = Err.Number: failure = Err.Description

 On Error GoTo 0

 AssertRunIdentity

 If eno = 0 And (state = "DONE" Or (state = "UNSAVED" And Not HasSavedUnverified(CStr(ob)))) Then

  For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

   If CStr(ws.Cells(r, 3).Value2) = CStr(ob) And IsPendingState(CStr(ws.Cells(r, 6).Value2)) Then

    ws.Cells(r, 6).Value2 = state: ws.Cells(r, 7).Value2 = Now: ws.Cells(r, 9).Value2 = "ARCHIVED VL03N READBACK " & state

   End If

  Next r

  PersistCheckpoint

 Else

  ModuleLog.WriteLog "PENDING_UNKNOWN", "Giu UNKNOWN va giu so luong batch/OB; " & failure, CStr(ob)

 End If

Next ob

End Sub

Private Sub ReservePendingStock()

Dim ws As Worksheet, stockws As Worksheet, r As Long, i As Long, remaining As Double, take As Double

Set ws = ThisWorkbook.Worksheets("LEDGER"): Set stockws = ThisWorkbook.Worksheets("POSTING")

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If IsPendingState(CStr(ws.Cells(r, 6).Value2)) Then

  remaining = CDbl(ws.Cells(r, 5).Value2)

  For i = 1 To UBound(mRemain)

   If CStr(mStock(i, 1)) = CStr(ws.Cells(r, 11).Value2) And CStr(mStock(i, 2)) = CStr(ws.Cells(r, 12).Value2) And CStr(mStock(i, 4)) = CStr(ws.Cells(r, 13).Value2) Then

    take = Application.Min(remaining, mRemain(i)): mRemain(i) = mRemain(i) - take: remaining = remaining - take

    If take > 0 Then stockws.Cells(i + 5, 10).Value2 = "RESERVED UNKNOWN; AVAIL=" & mRemain(i)

   End If

   If remaining <= 0 Then Exit For

  Next i

 End If

Next r

For i = 1 To UBound(mRemain): mInitialRemain(i) = mRemain(i): Next i

End Sub

Private Function requestKey(ByVal inputRow As Long) As String

Dim ws As Worksheet

Set ws = ThisWorkbook.Worksheets("INPUT")

requestKey = ModuleHelper.ToolConfig("REQUEST_ID", "TEST_001") & "|" & CStr(inputRow) & "|" & _
 ModuleHelper.CleanStr(ws.Cells(inputRow, 1).Value2) & "|" & CStr(ws.Cells(inputRow, 2).Value2) & "|" & _
 ModuleHelper.CleanStr(ws.Cells(inputRow, 3).Value2) & "|" & CStr(ws.Range("R8").Value2)

End Function

Private Function CompletedQty(ByVal inputRow As Long) As Double

Dim ws As Worksheet, ledger As Worksheet, r As Long, key As String

Dim mat As String, ownerVal As String, rowKey As String, parts() As String

Dim isMatch As Boolean, doneQty As Double

Set ws = ThisWorkbook.Worksheets("INPUT")

Set ledger = ThisWorkbook.Worksheets("LEDGER")

mat = ModuleHelper.CleanStr(ws.Cells(inputRow, 1).Value2)

ownerVal = Trim$(CStr(ws.Range("R8").Value2))

key = requestKey(inputRow)

For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

 rowKey = CStr(ledger.Cells(r, 1).Value2)

 isMatch = False

 If rowKey = key Then

  isMatch = True

 Else

  parts = Split(rowKey, "|")

  If UBound(parts) >= 2 Then

   ' RC20-DEF-05: Stable matching by row index, material, and owner

   If parts(0) = ModuleHelper.ToolConfig("REQUEST_ID", "TEST_001") And parts(1) = CStr(inputRow) And StrComp(parts(2), mat, vbTextCompare) = 0 Then

    If UBound(parts) >= 5 Then

     If StrComp(parts(5), ownerVal, vbTextCompare) = 0 And parts(3) = CStr(ws.Cells(inputRow, 2).Value2) And StrComp(parts(4), ModuleHelper.CleanStr(ws.Cells(inputRow, 3).Value2), vbTextCompare) = 0 Then isMatch = True

    Else

     isMatch = True

    End If

   End If

  End If

 End If

 If isMatch Then

  If IsPendingState(CStr(ledger.Cells(r, 6).Value2)) Then Err.Raise vbObjectError + 4100, , "OB " & ledger.Cells(r, 3).Value2 & " chua ro ket qua Save. Kiem tra SAP truoc khi chay lai."

  If CStr(ledger.Cells(r, 6).Value2) = "DONE" Then doneQty = doneQty + CDbl(ledger.Cells(r, 5).Value2)

 End If

Next r

CompletedQty = doneQty

End Function

Private Sub ValidateInput()

Dim ws As Worksheet, r As Long, mat As String, qty As Double, reason As String, n As Long, products As Object

Dim lastRowA As Long, lastRowB As Long, lastRowC As Long

Set ws = ThisWorkbook.Worksheets("INPUT")

' RC20-DEF-07: Block AutoFilter if active

If ws.FilterMode Then Err.Raise vbObjectError + 4101, , "Sheet INPUT dang bat bo loc (Filter/AutoFilter). Vui long tat bo loc truoc khi chay."

' RC20-DEF-07: Block hidden rows in range 6..205

For r = 6 To 205

 If ws.rows(r).Hidden Then Err.Raise vbObjectError + 4101, , "Dong " & r & " tren sheet INPUT dang bi an (Hidden). Vui long hien thi toan bo dong truoc khi chay."

Next r

' RC20-DEF-07: Detect data beyond supported row 205

lastRowA = ws.Cells(ws.rows.count, 1).End(xlUp).row

lastRowB = ws.Cells(ws.rows.count, 2).End(xlUp).row

lastRowC = ws.Cells(ws.rows.count, 3).End(xlUp).row

If lastRowA > 205 Or lastRowB > 205 Or lastRowC > 205 Then

 Err.Raise vbObjectError + 4101, , "Phat hien du lieu vuot qua dong 205 tren sheet INPUT (A=" & lastRowA & ", B=" & lastRowB & ", C=" & lastRowC & "). Vung nhap lieu ho tro tu dong 6 den 205."

End If

Set products = InputProducts()

' RC20-DEF-04: Owner (R8) required for both 1 and multiple materials

If Trim$(CStr(ws.Range("R8").Value2)) = "" Then Err.Raise vbObjectError + 4101, , "Owner (R8) bat buoc cho ca 1 va nhieu Material. Khong the phan bo khi khong xac minh duoc Owner."

If Trim$(CStr(ws.Range("R5").Value2)) = "" Or Trim$(CStr(ws.Range("R6").Value2)) = "" Then Err.Raise vbObjectError + 4101, , "Ship-to va SLoc bat buoc."

For r = 6 To 205

 mat = ModuleHelper.CleanStr(ws.Cells(r, 1).Value2)

 If mat = "" Then

  If ModuleHelper.CleanStr(ws.Cells(r, 2).Value2) <> "" Or ModuleHelper.CleanStr(ws.Cells(r, 3).Value2) <> "" Then Err.Raise vbObjectError + 4102, , "Dong " & r & ": thieu Material."

 Else

  If InStr(mat, "|") > 0 Or InStr(ModuleHelper.CleanStr(ws.Cells(r, 3).Value2), "|") > 0 Then Err.Raise vbObjectError + 4102, , "Material/Batch khong duoc chua |."

  If Trim$(CStr(ws.Cells(r, 2).Value2)) <> "" Then

  If Not ModuleHelper.ParseSAPIntegerQty(ws.Cells(r, 2).Value2, qty, reason) Then Err.Raise vbObjectError + 4103, , "Qty dong " & r & ": " & reason

  End If

  qty = CompletedQty(r)

  n = n + 1

 End If

Next r

If n = 0 Then Err.Raise vbObjectError + 4104, , "Nhap Material va Qty vao A6:C205."

If CStr(ws.Range("R5").Value2) = "4007" Then
 Call EDCNormalizeSTOs(CStr(ws.Range("R31").Value2))
End If
ValidateRequestIdentity

End Sub

Private Sub ValidateRequestIdentity()

Dim ws As Worksheet, r As Long, prefix As String, sig As String

Set ws = ThisWorkbook.Worksheets("LEDGER")

prefix = ModuleHelper.ToolConfig("REQUEST_ID", "TEST_001") & "|": sig = Fingerprint()

If InStr(Left$(prefix, Len(prefix) - 1), "|") > 0 Then Err.Raise vbObjectError + 4104, , "REQUEST_ID khong duoc chua |."

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If Left$(CStr(ws.Cells(r, 1).Value2), Len(prefix)) = prefix Then

  If CStr(ws.Cells(r, 10).Value2) <> sig Then Err.Raise vbObjectError + 4104, , "INPUT da thay doi so voi REQUEST_ID da chay. Khoi phuc INPUT de chay tiep, hoac doi REQUEST_ID cho yeu cau moi."

 End If

Next r

End Sub

Private Function BeginRun(ByVal action As String, Optional ByVal keepPlan As Boolean = False) As Boolean

If mRunning Then Exit Function

If Not mSettingsSaved Then
 On Error Resume Next
 mSavedCalculation = CLng(Application.Calculation)
 mSavedEnableEvents = CBool(Application.EnableEvents)
 mSavedScreenUpdating = CBool(Application.ScreenUpdating)
 mSettingsSaved = True
 On Error GoTo 0
End If

On Error Resume Next
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = -4135 ' xlCalculationManual
On Error GoTo 0

mRunning = True

If Not keepPlan Then mChecked = False

ModuleSAP.OB_ResetSession

RunID = Format$(Now, "yyyymmdd_hhnnss")

Application.StatusBar = action

BeginRun = True

End Function

Public Sub RestoreApplicationSettings()

On Error Resume Next

If mSettingsSaved Then
 Application.Calculation = mSavedCalculation
 Application.EnableEvents = mSavedEnableEvents
 Application.ScreenUpdating = mSavedScreenUpdating
 mSettingsSaved = False
Else
 Application.ScreenUpdating = True
 Application.EnableEvents = True
End If

Application.Cursor = xlDefault

Application.StatusBar = False

On Error GoTo 0

End Sub

Private Sub EndRun()

mRunning = False: Application.StatusBar = False

RestoreApplicationSettings

OB_UpdateModeButton

End Sub

Public Sub OB_RefreshPosting()

If Not BeginRun("POSTING") Then Exit Sub

On Error GoTo Fail

ArchivePendingSnapshots

EnsureRequestIdentity

ValidateInput

CaptureRequestIdentity

If Not ModuleSAP.DownloadPostingForOrder(False) Then Err.Raise vbObjectError + 4105, , ModuleSAP.LastPostingErrorStage & ": " & ModuleSAP.LastPostingErrorText

AssertRunIdentity

CopyPosting

ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = "POSTING UPDATED"

ModuleLog.WriteLog "OK", ModuleSAP.LastPostingTimingText

EndRun

MsgBox "Da tai POSTING theo Material (1 ma bo qua Owner). Nhan KIEM TRA de lap ke hoach OB.", vbInformation

Exit Sub

Fail:

ReportError "POSTING", Err.Description

End Sub

Private Sub CopyPosting()

Dim raw As Worksheet, ws As Worksheet, n As Long, a As Variant, out() As Variant, i As Long

Set raw = ThisWorkbook.Worksheets("POSTING_RAW"): Set ws = ThisWorkbook.Worksheets("POSTING")

ws.Range("A6:J" & CStr(Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row))).ClearContents

n = raw.Cells(raw.rows.count, 1).End(xlUp).row - 1

If n < 1 Then Exit Sub

a = raw.Range("A2:Q" & n + 1).Value2: ReDim out(1 To n, 1 To 10)

For i = 1 To n

 out(i, 1) = a(i, 1): out(i, 2) = a(i, 3): out(i, 3) = a(i, 2)

 out(i, 4) = a(i, 4): out(i, 5) = a(i, 5): out(i, 6) = a(i, 11): out(i, 7) = a(i, 12)

 If UCase$(Trim$(CStr(a(i, 5)))) <> UCase$(Trim$(CStr(a(i, 10)))) Then

  out(i, 5) = "TYPE_MISMATCH": out(i, 10) = "BLOCK SOURCE TYPE MISMATCH"

 Else

  out(i, 10) = "SNAPSHOT"

 End If

Next i

ws.Range("A6").Resize(n, 10).Value2 = out

End Sub

Public Sub OB_Check()

If Not BeginRun("KIEM TRA") Then Exit Sub

On Error GoTo Fail

If Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2)) = "2001" Then
    OB2001_Check
    EndRun
    Exit Sub
End If

BuildLivePlan

EndRun

If mChecked Then

 MsgBox "Da lap " & mAlloc.count & " cap phat co OB thuc te. Nut 2 chi xu ly Qty/SLoc/Batch bang VL02N. Xem BATCH_PLAN/NOTE cho cac ma con cho. PREVIEW co phan thieu: doi APPLY va nhan lai nut 1.", vbInformation

Else

 MsgBox "Chua co cap phat OB READY. Xem NOTE/BATCH_PLAN. PREVIEW chi kiem tra; doi APPLY va nhan lai nut 1 de tao OB thieu cho 4007. Nut 2 khong tao OB.", vbExclamation

End If

Exit Sub

Fail:

ReportError "CHECK", Err.Description

End Sub

Private Sub BuildLivePlan()
Dim mode As String
mode = UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW"))
If mode <> "PREVIEW" And mode <> "APPLY" Then Err.Raise vbObjectError + 4111, , "MODE chi nhan PREVIEW/APPLY."

ArchivePendingSnapshots

EnsureRequestIdentity

ValidateRequestIdentity

CaptureRequestIdentity

ResolveSaveReceipts

If HasCurrentPendingPlan() Then ReconcileUnknown

ReconcileArchivedPending

ValidateInput

ThisWorkbook.Worksheets("INPUT").Range("D6:D205").ClearContents

ThisWorkbook.Worksheets("INPUT").Range("F6:N505").ClearContents

If Not ModuleSAP.DownloadPostingForOrder(False) Then Err.Raise vbObjectError + 4105, , ModuleSAP.LastPostingErrorStage & ": " & ModuleSAP.LastPostingErrorText

CopyPosting
mPostingCheckedAt = Timer
mEDCPostingSnapshot = SourceFingerprint()

Set mEDCSTOByMat = CreateObject("Scripting.Dictionary")
Set mEDCWaitByMat = CreateObject("Scripting.Dictionary")
Set mEDCDuplicateInvestigated = CreateObject("Scripting.Dictionary")
ModuleSAP.OB_ReadVL06O InputProducts()
mEDCOBCheckedAt = Timer

Set mExcludedOB = CreateObject("Scripting.Dictionary")

CaptureDoneBaseline

EDCResolvePendingCreate
EDCBuildCreatePlans InputProducts()
If mode = "APPLY" Then EDCExecuteCreatePlans
BuildPlan
EDCRestoreWaitNotes
If mode = "PREVIEW" Then EDCWriteCreatePlans

AssertRunIdentity

FormatOperationalTables

mPlanViewSignature = PlanViewFingerprint()
mSourceSignature = SourceFingerprint()
mPlanCheckedAt = Now

mPlanPolicy = PlanPolicy()

mChecked = (mAlloc.count > 0)

If mChecked Then

  Dim newlyCreatedText As String: newlyCreatedText = ""
  If Not mEDCCreatePlan Is Nothing Then
   If mEDCCreatePlan.count > 0 Then newlyCreatedText = "TAO OB THANH CONG: "
  End If
  If newlyCreatedText <> "" Then
   ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = newlyCreatedText & mAlloc.count & " ALLOC; SAN SANG BAM NUT 2; MODE=" & mode
  Else
   ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = "PLAN READY: " & mAlloc.count & " ALLOC; OB VERIFIED; MODE=" & mode
  End If

Else

 ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = IIf(mode = "PREVIEW" And EDCPlanCount() > 0, "PREVIEW - DOI APPLY VA NHAN LAI NUT 1 DE TAO OB", "BLOCK - XEM NOTE / BATCH_PLAN")

End If

ThisWorkbook.Save

' RC20-DEF-01: Removed undeclared timingEvent logging in BuildLivePlan

End Sub

Private Sub ReconcileUnknown()

Dim ledger As Worksheet, plan As Worksheet, cache As Worksheet, obs As Object, ob As Variant

Dim r As Long, p As Long, q As Long, j As Long, key As String, items As Collection, a As Variant

Dim foundLedger As Boolean, foundCache As Boolean, state As String, originalQty As Double, originalSloc As String, oldRuns As Object

Set ledger = ThisWorkbook.Worksheets("LEDGER"): Set plan = ThisWorkbook.Worksheets("BATCH_PLAN")

Set cache = ThisWorkbook.Worksheets("VL06O_DATA"): Set obs = CreateObject("Scripting.Dictionary")

For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

 If IsPendingState(CStr(ledger.Cells(r, 6).Value2)) And CStr(ledger.Cells(r, 10).Value2) = mSignature Then obs(CStr(ledger.Cells(r, 3).Value2)) = True

Next r

' VBS may have completed a different READY OB from the same failed plan manually.

' Read those persisted states too, so confirmed quantities are not allocated again.

If obs.count > 0 Then

 Set oldRuns = CreateObject("Scripting.Dictionary")

 For p = 6 To plan.Cells(plan.rows.count, 1).End(xlUp).row

  If obs.exists(CStr(plan.Cells(p, 3).Value2)) Then oldRuns(CStr(plan.Cells(p, 1).Value2)) = True

 Next p

 For p = 6 To plan.Cells(plan.rows.count, 1).End(xlUp).row

  If CStr(plan.Cells(p, 12).Value2) = "READY" And oldRuns.exists(CStr(plan.Cells(p, 1).Value2)) Then obs(CStr(plan.Cells(p, 3).Value2)) = True

 Next p

End If

For Each ob In obs.keys

 Set items = New Collection

 For p = 6 To plan.Cells(plan.rows.count, 1).End(xlUp).row

  If CStr(plan.Cells(p, 3).Value2) = CStr(ob) Then

   key = FrozenRequestKey(CLng(plan.Cells(p, 2).Value2)) & "|" & ob & "|" & CStr(plan.Cells(p, 4).Value2) & "|" & CStr(plan.Cells(p, 6).Value2) & "|" & CStr(plan.Cells(p, 7).Value2)

   foundLedger = False

   For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

    If CStr(ledger.Cells(r, 8).Value2) = key And IsPendingState(CStr(ledger.Cells(r, 6).Value2)) And CStr(ledger.Cells(r, 10).Value2) = mSignature Then

     If CDbl(ledger.Cells(r, 5).Value2) <> CDbl(plan.Cells(p, 8).Value2) Then Err.Raise vbObjectError + 4113, , "Recovery plan qty khac LEDGER."

     foundLedger = True: Exit For

    End If

   Next r

   If Not foundLedger And CStr(plan.Cells(p, 12).Value2) <> "READY" Then Err.Raise vbObjectError + 4113, , "Recovery thieu UNKNOWN ledger cho OB " & ob & ". Khong tu xoa lich su."

   If CStr(plan.Cells(p, 5).Value2) <> ModuleHelper.CleanStr(ThisWorkbook.Worksheets("INPUT").Cells(CLng(plan.Cells(p, 2).Value2), 1).Value2) Then Err.Raise vbObjectError + 4113, , "Recovery plan Material khac INPUT."

   foundCache = False

   For q = 6 To cache.Cells(cache.rows.count, 1).End(xlUp).row

    If CStr(cache.Cells(q, 1).Value2) = CStr(ob) And CStr(cache.Cells(q, 2).Value2) = CStr(plan.Cells(p, 4).Value2) And CStr(cache.Cells(q, 3).Value2) = CStr(plan.Cells(p, 5).Value2) Then

     If CStr(cache.Cells(q, 5).Value2) <> "" Then Err.Raise vbObjectError + 4113, , "Recovery snapshot khong phai batch trong."

     originalQty = CDbl(cache.Cells(q, 4).Value2): originalSloc = CStr(cache.Cells(q, 7).Value2)

     foundCache = True: Exit For

    End If

   Next q

   If Not foundCache Then Err.Raise vbObjectError + 4113, , "Recovery thieu snapshot OB/item. Kiem tra SAP."

   items.Add Array(CLng(plan.Cells(p, 2).Value2), CStr(ob), CStr(plan.Cells(p, 4).Value2), CStr(plan.Cells(p, 5).Value2), CStr(plan.Cells(p, 6).Value2), CStr(plan.Cells(p, 7).Value2), CDbl(plan.Cells(p, 8).Value2), CStr(plan.Cells(p, 10).Value2), CStr(plan.Cells(p, 11).Value2), originalQty, originalSloc)

  End If

 Next p

 If items.count = 0 Then Err.Raise vbObjectError + 4113, , "UNKNOWN OB " & ob & " khong co ke hoach cu de doi chieu."

 ' Every unknown allocation must be covered; never clear a partial snapshot.

 For r = 2 To ledger.Cells(ledger.rows.count, 1).End(xlUp).row

  If CStr(ledger.Cells(r, 3).Value2) = CStr(ob) And IsPendingState(CStr(ledger.Cells(r, 6).Value2)) And CStr(ledger.Cells(r, 10).Value2) = mSignature Then

   foundLedger = False

   For Each a In items

    key = FrozenRequestKey(CLng(a(0))) & "|" & a(1) & "|" & a(2) & "|" & a(4) & "|" & a(5)

    If key = CStr(ledger.Cells(r, 8).Value2) Then foundLedger = True: Exit For

   Next a

   If Not foundLedger Then Err.Raise vbObjectError + 4113, , "Recovery plan khong bao phu het UNKNOWN."

  End If

 Next r

 state = ModuleSAP.OB_ResolveUnknown(CStr(ob), items)

 If state = "UNSAVED" And HasSavedUnverified(CStr(ob)) Then Err.Raise vbObjectError + 4113, , "DA SAVE OB " & ob & "; readback chua khop. Khong ghi lai. Kiem tra SAP."

 Journal items, state, "VL03N PERSISTED READBACK " & state

 PersistCheckpoint

 ModuleLog.WriteLog "RECOVERY", "VL03N persisted state=" & state, CStr(ob)

Next ob

End Sub

Private Sub CaptureRequestIdentity()

Dim r As Long

mSignature = Fingerprint()

Set mRequestKeys = CreateObject("Scripting.Dictionary")

For r = 6 To 205

 If ModuleHelper.CleanStr(ThisWorkbook.Worksheets("INPUT").Cells(r, 1).Value2) <> "" Then mRequestKeys(CStr(r)) = requestKey(r)

Next r

End Sub

Public Sub AssertRunIdentity()

If mSignature = "" Or Fingerprint() <> mSignature Then Err.Raise vbObjectError + 4110, , "INPUT/Request da thay doi trong luc chay. Dung; khong tiep tuc ghi SAP."

End Sub

Private Function FrozenRequestKey(ByVal inputRow As Long) As String

If mRequestKeys Is Nothing Then Err.Raise vbObjectError + 4110, , "Chua chup Request identity."

If Not mRequestKeys.exists(CStr(inputRow)) Then Err.Raise vbObjectError + 4110, , "Thieu Request identity."

FrozenRequestKey = CStr(mRequestKeys(CStr(inputRow)))

End Function

Private Function BinAllowed(ByVal bin As String) As Boolean

bin = UCase$(Trim$(bin))

If bin = "" Then Exit Function

If bin = "APFL-BL01-01-01" Or bin Like "APFL-HO*" Or bin Like "APFL-HTK*" Or bin Like "HTK*" Then Exit Function

BinAllowed = True

End Function

Private Function DateRank(ByVal v As Variant) As Double

Dim d As Date

d = ModuleHelper.ParseSapDate(v)

If d = 0 Then DateRank = 1E+20 Else DateRank = CDbl(d)

End Function

Private Function TimeRank(ByVal v As Variant) As Double

On Error GoTo Invalid

Dim s As String

s = Trim$(CStr(v & ""))

' RC20-DEF-06: Empty GR TIME defaults to 0.0 (start of day)

If s = "" Then

 TimeRank = 0#

 Exit Function

End If

If IsNumeric(v) And VarType(v) <> vbString Then

 If CDbl(v) >= 0# And CDbl(v) < 1# Then TimeRank = CDbl(v): Exit Function

End If

TimeRank = CDbl(TimeValue(s)): Exit Function

Invalid:

' RC20-DEF-06: 2.0 indicates invalid time format

TimeRank = 2#

End Function

Private Function StockBefore(ByVal x As Long, ByVal y As Long) As Boolean

If mStockRank(x) <> mStockRank(y) Then

 StockBefore = (mStockRank(x) < mStockRank(y))

Else

 StockBefore = (mStockTie(x) < mStockTie(y))

End If

End Function

Private Sub SortStock(ByVal n As Long)

Dim temp() As Long, width As Long, lo As Long, mid As Long, hi As Long

Dim i As Long, j As Long, k As Long

If n < 2 Then Exit Sub

ReDim temp(1 To n)

width = 1

Do While width < n

 For lo = 1 To n Step width * 2

  mid = lo + width - 1: If mid > n Then mid = n

  hi = lo + width * 2 - 1: If hi > n Then hi = n

  i = lo: j = mid + 1

  For k = lo To hi

   If i > mid Then

    temp(k) = mStockOrder(j): j = j + 1

   ElseIf j > hi Then

    temp(k) = mStockOrder(i): i = i + 1

   ElseIf StockBefore(mStockOrder(j), mStockOrder(i)) Then

    temp(k) = mStockOrder(j): j = j + 1

   Else

    temp(k) = mStockOrder(i): i = i + 1

   End If

  Next k

 Next lo

 For k = 1 To n: mStockOrder(k) = temp(k): Next k

 width = width * 2

Loop

End Sub

Private Function StockTypeInList(ByVal st As String, ByVal configured As String) As Boolean

Dim v As Variant

For Each v In Split(Replace(UCase$(configured), ";", ","), ",")

 If st <> "" And st = Trim$(CStr(v)) Then StockTypeInList = True: Exit Function

Next v

End Function

Private Sub PrepareStock()

Dim ws As Worksheet, n As Long, i As Long, j As Long, tmp As Long, blocked As Object, key As String, st As String

Set ws = ThisWorkbook.Worksheets("POSTING")

n = ws.Cells(ws.rows.count, 1).End(xlUp).row - 5

If n < 1 Then n = 1

mStock = ws.Range("A6:J" & n + 5).Value2

ReDim mRemain(1 To n): ReDim mStockOrder(1 To n): ReDim mInitialRemain(1 To n)

ReDim mStockRank(1 To n): ReDim mStockTie(1 To n)

Set blocked = CreateObject("Scripting.Dictionary"): blocked.CompareMode = vbTextCompare
Set mBlockedStockByMat = CreateObject("Scripting.Dictionary"): mBlockedStockByMat.CompareMode = vbTextCompare

For i = 1 To n

 key = CStr(mStock(i, 1)) & "|" & CStr(mStock(i, 2))

 st = UCase$(Trim$(CStr(mStock(i, 5))))

 If StockTypeInList(st, ModuleHelper.ToolConfig("POST_BLOCKED_STOCK_TYPES", "QT,QI,QU")) Or StockTypeInList(st, "QT,QI,QU") Then
  blocked(key) = True
  mBlockedStockByMat(CStr(mStock(i, 1))) = True
 End If

Next i

For i = 1 To n

 mStockOrder(i) = i

 mStockRank(i) = DateRank(mStock(i, 6)) + TimeRank(mStock(i, 7))

 mStockTie(i) = CStr(mStock(i, 2)) & "|" & CStr(mStock(i, 4))

 key = CStr(mStock(i, 1)) & "|" & CStr(mStock(i, 2))

 If Not blocked.exists(key) And StockTypeInList(UCase$(Trim$(CStr(mStock(i, 5)))), ModuleHelper.ToolConfig("POST_ALLOWED_STOCK_TYPES", "FU")) And BinAllowed(CStr(mStock(i, 4))) And CStr(mStock(i, 2)) <> "" And DateRank(mStock(i, 6)) < 1E+20 And TimeRank(mStock(i, 7)) < 2 Then

  mRemain(i) = ModuleHelper.CleanDouble(mStock(i, 3))

 End If

 If mRemain(i) > 0 Then

  ws.Cells(i + 5, 10).Value2 = "ELIGIBLE"

 ElseIf blocked.exists(key) Then

  ws.Cells(i + 5, 10).Value2 = "BLOCK BATCH QT/QI"

 ElseIf Not StockTypeInList(UCase$(Trim$(CStr(mStock(i, 5)))), ModuleHelper.ToolConfig("POST_ALLOWED_STOCK_TYPES", "FU")) Then

  ws.Cells(i + 5, 10).Value2 = "BLOCK STOCK TYPE " & CStr(mStock(i, 5))
  If StockTypeInList(st, "QT,QI,QU") Then mBlockedStockByMat(CStr(mStock(i, 1))) = True

 ElseIf DateRank(mStock(i, 6)) >= 1E+20 Then

  ws.Cells(i + 5, 10).Value2 = "BLOCK INVALID GR DATE"

 ElseIf TimeRank(mStock(i, 7)) >= 2 Then

  ws.Cells(i + 5, 10).Value2 = "BLOCK INVALID GR TIME"

 ElseIf Not BinAllowed(CStr(mStock(i, 4))) Then

  ws.Cells(i + 5, 10).Value2 = "BLOCK BIN " & CStr(mStock(i, 4))

 ElseIf CStr(mStock(i, 2)) = "" Then

  ws.Cells(i + 5, 10).Value2 = "BLOCK EMPTY BATCH"

 ElseIf ModuleHelper.CleanDouble(mStock(i, 3)) <= 0 Then

  ws.Cells(i + 5, 10).Value2 = "BLOCK QTY <= 0"

 Else

  ws.Cells(i + 5, 10).Value2 = "BLOCK BIN / BATCH / DATE / QTY"

 End If

 mInitialRemain(i) = mRemain(i)

Next i

ReservePendingStock

SortStock n

End Sub

Private Sub BuildPlan(Optional ByVal onlyMaterial As String = "")

Dim ws As Worksheet, obws As Worksheet, n As Long, i As Long, pass As Long, r As Long, mat As String, batch As String

Dim qty As Double, already As Double, target As Double, preserved As Collection, a As Variant

Set ws = ThisWorkbook.Worksheets("INPUT"): Set obws = ThisWorkbook.Worksheets("VL06O_DATA")

If onlyMaterial = "" Then

 ws.Range("D6:D205").ClearContents

 Set mAlloc = New Collection: Set mTargets = CreateObject("Scripting.Dictionary")
 Set mWholeNeeds = CreateObject("Scripting.Dictionary"): mWholeNeeds.CompareMode = vbTextCompare

Else

 Set preserved = New Collection

 For Each a In mAlloc

  If StrComp(CStr(a(3)), onlyMaterial, vbTextCompare) <> 0 Then preserved.Add a

 Next a

 Set mAlloc = preserved

End If

ws.Range("F6:N505").ClearContents

PrepareStock

If onlyMaterial <> "" Then CapRecoveryStock onlyMaterial

ModuleSAP.OB_ClassifyCandidates obws
n = obws.Cells(obws.rows.count, 1).End(xlUp).row - 5

If n > 0 Then

 mOB = obws.Range("A6:P" & n + 5).Value2: ReDim mOBCap(1 To n)

 For i = 1 To n

  If PendingOB(CStr(mOB(i, 1))) Then

   mOB(i, 10) = "BLOCK": mOB(i, 11) = "UNKNOWN FROM PREVIOUS REQUEST"

   obws.Cells(i + 5, 10).Value2 = mOB(i, 10): obws.Cells(i + 5, 11).Value2 = mOB(i, 11)

  ElseIf IsExcludedOB(CStr(mOB(i, 1))) Or ReservedByOtherMaterial(CStr(mOB(i, 1)), onlyMaterial) Then

   mOB(i, 10) = "BLOCK": mOB(i, 11) = "EXCLUDED DURING RECOVERY"

  ElseIf CStr(mOB(i, 10)) = "CANDIDATE" Then

   mOBCap(i) = CDbl(mOB(i, 4))

  End If

 Next i

Else

 mOB = Empty: ReDim mOBCap(1 To 1)

End If

For pass = 1 To 2

 For r = 6 To 205

  mat = ModuleHelper.CleanStr(ws.Cells(r, 1).Value2): batch = ModuleHelper.CleanStr(ws.Cells(r, 3).Value2)

  If mat <> "" And (onlyMaterial = "" Or StrComp(mat, onlyMaterial, vbTextCompare) = 0) And ((pass = 1 And batch <> "") Or (pass = 2 And batch = "")) Then

   If Trim$(CStr(ws.Cells(r, 2).Value2)) = "" Then

    AllocateWholeMaterial r, mat, batch, n

   Else

   qty = ModuleHelper.CleanDouble(ws.Cells(r, 2).Value2): already = CompletedQty(r)

   target = qty - already

   If target <= 0 Then

    ws.Cells(r, 4).Value2 = "DONE " & already

   Else

    mTargets(CStr(r)) = target

    AllocateRequest r, mat, batch, target, n

   End If

   End If

  End If

 Next r

Next pass

ConsolidateAlloc

WritePlan

End Sub

Private Function IsExcludedOB(ByVal ob As String) As Boolean

If Not mExcludedOB Is Nothing Then IsExcludedOB = mExcludedOB.exists(ob)

End Function

Private Sub CaptureDoneBaseline()

Dim ws As Worksheet, i As Long, key As String

Set mDoneBaseline = CreateObject("Scripting.Dictionary"): mDoneBaseline.CompareMode = vbTextCompare

Set ws = ThisWorkbook.Worksheets("LEDGER")

For i = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If CStr(ws.Cells(i, 6).Value2) = "DONE" Then

  key = CStr(ws.Cells(i, 11).Value2) & "|" & CStr(ws.Cells(i, 12).Value2)

  If Not mDoneBaseline.exists(key) Then mDoneBaseline(key) = 0#

  mDoneBaseline(key) = CDbl(mDoneBaseline(key)) + CDbl(ws.Cells(i, 5).Value2)

 End If

Next i

End Sub

Private Sub CapRecoveryStock(ByVal mat As String)

Dim remaining As Object, k As Variant, i As Long, ix As Long, batch As String, allowed As Double

If mRecoveryCaps Is Nothing Then Exit Sub

Set remaining = CreateObject("Scripting.Dictionary")

For Each k In mRecoveryCaps.keys: remaining(CStr(k)) = CDbl(mRecoveryCaps(k)): Next k

For i = 1 To UBound(mStockOrder)

 ix = mStockOrder(i): batch = CStr(mStock(ix, 2))

 If StrComp(CStr(mStock(ix, 1)), mat, vbTextCompare) = 0 Then

  If remaining.exists(batch) Then

   allowed = Application.Min(mRemain(ix), CDbl(remaining(batch)))

   mRemain(ix) = allowed: remaining(batch) = CDbl(remaining(batch)) - allowed

  Else

   mRemain(ix) = 0

  End If

 End If

Next i

End Sub

Private Sub AllocateWholeMaterial(ByVal req As Long, ByVal mat As String, ByVal requiredBatch As String, ByVal obN As Long)

Dim totals As Object, ix As Long, i As Long, o As Long, batch As String, k As Variant

Dim cap As Double, targetQty As Double, batchQty As Double

Set totals = CreateObject("Scripting.Dictionary"): totals.CompareMode = vbTextCompare

For o = 1 To obN

 If StrComp(CStr(mOB(o, 3)), mat, vbTextCompare) = 0 Then cap = cap + mOBCap(o)

Next o

For i = 1 To UBound(mStockOrder)

 ix = mStockOrder(i): batch = CStr(mStock(ix, 2))

 If StrComp(CStr(mStock(ix, 1)), mat, vbTextCompare) = 0 And mRemain(ix) > 0 And (requiredBatch = "" Or StrComp(batch, requiredBatch, vbTextCompare) = 0) Then

  If Not totals.exists(batch) Then totals(batch) = 0#

  totals(batch) = CDbl(totals(batch)) + mRemain(ix)

 End If

Next i

targetQty = 0

For Each k In totals.keys

 batchQty = CDbl(totals(k))

 If batchQty <= cap Then

  AllocateRequest req, mat, CStr(k), batchQty, obN

  cap = cap - batchQty: targetQty = targetQty + batchQty

 Else

  ModuleLog.WriteLog "WAIT_FULL_BATCH_CAPACITY", mat & "; batch=" & k & "; remaining=" & batchQty & "; done=" & CompletedBatchQty(req, CStr(k)) & "; OB capacity=" & cap
  mWholeNeeds(CStr(req) & "|" & CStr(k)) = batchQty + CompletedBatchQty(req, CStr(k))

 End If

Next k

mTargets(CStr(req)) = targetQty

ThisWorkbook.Worksheets("INPUT").Cells(req, 4).Value2 = IIf(targetQty > 0, "READY WHOLE BATCH " & targetQty, "WAIT_FULL_BATCH_CAPACITY: KHONG DU OB CHO NGUYEN BATCH")

End Sub

Private Sub MergeOtherMaterials(ByVal ws As Worksheet, ByVal oldRows As Variant, ByVal firstRow As Long, ByVal materialColumn As Long, ByVal columnCount As Long, ByVal mat As String)

Dim i As Long, j As Long, nextRow As Long, n As Long, at As Long, otherRows() As Variant

If Not IsArray(oldRows) Then Exit Sub

For i = 1 To UBound(oldRows, 1)

 If CStr(oldRows(i, 1)) <> "" And StrComp(CStr(oldRows(i, materialColumn)), mat, vbTextCompare) <> 0 Then n = n + 1

Next i

If n = 0 Then Exit Sub

ReDim otherRows(1 To n, 1 To columnCount)

For i = 1 To UBound(oldRows, 1)

 If CStr(oldRows(i, 1)) <> "" And StrComp(CStr(oldRows(i, materialColumn)), mat, vbTextCompare) <> 0 Then

  at = at + 1

  For j = 1 To columnCount: otherRows(at, j) = oldRows(i, j): Next j

 End If

Next i

nextRow = Application.Max(firstRow, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)

ws.Cells(nextRow, 1).Resize(n, columnCount).Value2 = otherRows

End Sub

Private Sub ReplanOneMaterial(ByVal mat As String, Optional ByVal refreshStock As Boolean = True)

Dim postRows As Variant, obRows As Variant, ws As Worksheet, a As Variant
Dim savedAlloc As Collection, Failed As String, eno As Long, ledger As Worksheet, i As Long, batch As String, key As Variant
Dim cachedCandidateFound As Boolean, rScan As Long, candOBScan As String, lastRowScan As Long
Dim replReason As String, singleMatDict As Object, newAllocOB As String, targetSTO As String, aCheck As Variant

On Error GoTo Fail

Set savedAlloc = mAlloc
Set ws = ThisWorkbook.Worksheets("POSTING")
postRows = ws.Range("A6:J" & Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)).Value2

If mRecoveryPrepared Then
 Set mRecoveryCaps = mRecoveryCapMap(mat)
Else
 Set mRecoveryCaps = RecoveryCapsFor(mat, postRows)
End If

Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
obRows = ws.Range("A6:P" & Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)).Value2

mRecoveryMaterial = mat

If Not mRecoveryPrepared Then Err.Raise vbObjectError + 4120, , "Recovery batch was not prepared."
replReason = "Batched scoped Posting/VL06O"

mRecoveryMaterial = ""
BuildPlan mat
Set mRecoveryCaps = Nothing
RestorePlanStates

newAllocOB = ""
targetSTO = ""
For Each aCheck In mAlloc
 If StrComp(CStr(aCheck(3)), mat, vbTextCompare) = 0 And Not IsExcludedOB(CStr(aCheck(1))) Then
  newAllocOB = CStr(aCheck(1))
  targetSTO = CStr(aCheck(7))
  Exit For
 End If
Next aCheck

If mRetryCounts Is Nothing Then Set mRetryCounts = CreateObject("Scripting.Dictionary")
If Not mRetryCounts.exists(mat) Then mRetryCounts(mat) = 0
mRetryCounts(mat) = CLng(mRetryCounts(mat)) + 1

If newAllocOB <> "" Then
 ModuleLog.WriteLog "REPLACEMENT_OB", "OB cu: " & mLastSkippedOB & "; OB moi: " & newAllocOB & "; Ma: " & mat & "; STO: " & targetSTO & "; Ly do: " & replReason & "; Retry=" & mRetryCounts(mat), newAllocOB
Else
 ModuleLog.WriteLog "WAIT_REPLACEMENT_OB", "Het ung vien thay the cho ma " & mat & "; STO: " & targetSTO & "; giu nguyen va tiep tuc ma khac", ""
 EDCSetInputNote mat, "THIEU OB (VL06O) - Nhan nut 1 de tao lai qua VL10D hoac mo VL10D kiem tra"
 mMissingOBWarnUser = True
End If

ModuleLog.WriteLog "REPLAN_ONE_MATERIAL", mat & "; other material allocations retained"
PersistCheckpoint

Exit Sub

Fail:
eno = Err.Number: Failed = Err.Description: mRecoveryMaterial = "": Set mRecoveryCaps = Nothing: Set mAlloc = savedAlloc
Err.Raise eno, , Failed

End Sub

Private Sub RestorePlanStates()

Dim ws As Worksheet, r As Long, ob As String, state As String, ui As Worksheet, plan As Worksheet, p As Long

Set ws = ThisWorkbook.Worksheets("LEDGER"): Set ui = ThisWorkbook.Worksheets("INPUT"): Set plan = ThisWorkbook.Worksheets("BATCH_PLAN")

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If CStr(ws.Cells(r, 10).Value2) = mSignature Then

  state = CStr(ws.Cells(r, 6).Value2): ob = CStr(ws.Cells(r, 3).Value2)

  If state = "DONE" Or IsPendingState(state) Then

   For p = 6 To 505

    If CStr(ui.Cells(p, 7).Value2) = ob Then ui.Cells(p, 13).Value2 = state

   Next p

   For p = 6 To plan.Cells(plan.rows.count, 1).End(xlUp).row

    If CStr(plan.Cells(p, 3).Value2) = ob Then plan.Cells(p, 12).Value2 = state

   Next p

  End If

 End If

Next r

End Sub

Private Function TryResolveConflict(ByVal ob As String, ByVal items As Collection) As String

On Error GoTo Unresolved

TryResolveConflict = ModuleSAP.OB_ResolveUnknown(ob, items)

Exit Function

Unresolved:

ModuleLog.WriteLog "CONFLICT_READBACK_UNAVAILABLE", Err.Description, ob

TryResolveConflict = "UNRESOLVED"

End Function

Private Function ConfirmMissingOB(ByVal ob As String, ByVal items As Collection) As Boolean

Dim ws As Worksheet, oldRows As Variant, products As Object, a As Variant, mat As Variant, r As Long, eno As Long, detail As String

On Error GoTo Fail

Set ws = ThisWorkbook.Worksheets("VL06O_DATA")

oldRows = ws.Range("A6:P" & Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)).Value2

Set products = CreateObject("Scripting.Dictionary")

For Each a In items: products(CStr(a(3))) = True: Next a

ModuleSAP.OB_ReadVL06O products, False

For r = 6 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If CStr(ws.Cells(r, 1).Value2) = ob Then

  ModuleLog.WriteLog "MISSING_RECHECK", "OB still present in fresh VL06O; defer", ob

  GoTo Restore

 End If

Next r

ConfirmMissingOB = True

ModuleLog.WriteLog "MISSING_RECHECK", "OB absent from fresh scoped VL06O", ob

Set mRecoveryProducts = products
MergeUntouched ws, oldRows, 6, 3, 16
Exit Function

Restore:

ws.Range("A6:P" & Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)).ClearContents

ws.Range("A6").Resize(UBound(oldRows, 1), 16).Value2 = oldRows

Exit Function

Fail:

eno = Err.Number: detail = Err.Description

On Error Resume Next

If IsArray(oldRows) Then

 ws.Range("A6:P" & Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)).ClearContents

 ws.Range("A6").Resize(UBound(oldRows, 1), 16).Value2 = oldRows

End If

On Error GoTo 0

ModuleLog.WriteLog "MISSING_RECHECK_WAIT", detail, ob

ConfirmMissingOB = False

End Function

Private Function ApplyOneOB(ByVal ob As String, ByVal items As Collection, ByRef message As String) As Boolean

Dim failure As String, state As String, kind As String

On Error GoTo Fail

ModuleSAP.LastApplyState = ""
If UCase$(ModuleHelper.ToolConfig("TEST_SCOPE", "FIRST")) <> "ALL" Then
 If Not FirstScopeWhole(items) Then message = "WAIT_FULL_BATCH_CAPACITY: TEST_SCOPE=FIRST; use ALL": Exit Function
End If
If Not WholeBatchGate(items, message) Then Exit Function

Journal items, "UNKNOWN", "BEFORE_EDIT": PersistCheckpoint "EXCEL_SAVE_BEFORE_SEC", ob

ModuleSAP.OB_Apply ob, items, message

If Left$(message, 5) <> "SAVED" Then Err.Raise vbObjectError + 4121, , message

Journal items, "DONE", message: PersistCheckpoint "EXCEL_SAVE_AFTER_SEC", ob

ApplyOneOB = True: Exit Function

Fail:

failure = Err.Description: state = ModuleSAP.LastApplyState: kind = ModuleSAP.LastConflictKind

Resume HandleFailure

HandleFailure:

On Error GoTo Propagate

If (state = "UNSAVED" Or state = "SKIPPED_ALREADY_PROCESSED") And ModuleSAP.LastConflictEligible Then

 If kind = "ALREADY_PROCESSED" Or InStr(1, failure, "SKIP_OB_ALREADY_PROCESSED", vbTextCompare) > 0 Then

  mLastSkippedOB = ob

  Journal items, "UNSAVED", failure: PersistCheckpoint

  mExcludedOB(ob) = True

  ModuleLog.WriteLog "SKIP_CONFLICT", failure, ob

  message = failure: Exit Function

 End If

 If kind = "MISSING" Then

  If Not ConfirmMissingOB(ob, items) Then

   message = "WAIT_REMOTE_SYNC: OB chua duoc xac nhan mat; " & failure

   Journal items, "UNSAVED", message: PersistCheckpoint

   ModuleLog.WriteLog "DEFER_OB", message, ob

   Exit Function

  End If

 End If

 If kind = "WAIT" Then

  If InStr(1, failure, "WAIT_REMOTE_SYNC", vbTextCompare) > 0 Then

   message = "WAIT_REMOTE_SYNC: " & failure

  Else

   message = "WAIT_LOCK: " & failure

  End If

  Journal items, "UNSAVED", message: PersistCheckpoint

  ModuleLog.WriteLog "DEFER_OB", message, ob

  Exit Function

 End If

 If kind <> "MISSING" Then

  state = TryResolveConflict(ob, items)

  If state = "DONE" Then

   message = "EXTERNAL DONE: verified saved OB " & ob

   Journal items, "DONE", message: PersistCheckpoint: ApplyOneOB = True: Exit Function

  End If

  If state = "UNRESOLVED" Then

   message = "WAIT_READBACK: " & failure

   Journal items, "UNSAVED", message: PersistCheckpoint

   ModuleLog.WriteLog "DEFER_OB", message, ob: Exit Function

  End If

 End If

 Journal items, "UNSAVED", failure: PersistCheckpoint

 mExcludedOB(ob) = True

 ModuleLog.WriteLog "SKIP_CONFLICT", failure, ob

 message = failure: Exit Function

End If

If state = "" Then state = "UNKNOWN"

Journal items, state, failure: PersistCheckpoint

Err.Raise vbObjectError + 4121, , failure

Propagate:

Err.Raise vbObjectError + 4121, , Err.Description

End Function

Private Sub ConsolidateAlloc()

Dim d As Object, a As Variant, v As Variant, k As Variant, key As String

Set d = CreateObject("Scripting.Dictionary")

For Each a In mAlloc

 key = CStr(a(0)) & "|" & a(1) & "|" & a(2) & "|" & a(4) & "|" & a(5)

 If d.exists(key) Then

  v = d(key): v(6) = CDbl(v(6)) + CDbl(a(6)): d(key) = v

 Else

  d.Add key, a

 End If

Next a

Set mAlloc = New Collection

For Each k In d.keys: mAlloc.Add d(k): Next k

End Sub

Private Sub AllocateRequest(ByVal req As Long, ByVal mat As String, ByVal requiredBatch As String, ByVal qty As Double, ByVal obN As Long)

Dim localAlloc As Collection, localStock As Object, localOB As Object, i As Long, ix As Long, k As Variant, o As Long

Dim available As Double, cap As Double, take As Double, need As Double, obNeed As Double, sr As Variant, bk As String

Dim grouped As Object, groupRows As Object, b As Variant, batchTotal As Double, bRows As Collection

Dim ws As Worksheet, shortage As String

Set ws = ThisWorkbook.Worksheets("INPUT")

Set localAlloc = New Collection: Set localStock = CreateObject("Scripting.Dictionary"): Set localOB = CreateObject("Scripting.Dictionary")

Set grouped = CreateObject("Scripting.Dictionary"): grouped.CompareMode = vbTextCompare

For i = 1 To UBound(mStock, 1)

 ix = mStockOrder(i): bk = CStr(mStock(ix, 2))

 If StrComp(CStr(mStock(ix, 1)), mat, vbTextCompare) = 0 And mRemain(ix) > 0 And (requiredBatch = "" Or StrComp(bk, requiredBatch, vbTextCompare) = 0) Then

  If Not grouped.exists(bk) Then Set bRows = New Collection: grouped.Add bk, bRows

  grouped(bk).Add ix

  available = available + mRemain(ix)

 End If

Next i

For o = 1 To obN

 If StrComp(CStr(mOB(o, 3)), mat, vbTextCompare) = 0 Then cap = cap + mOBCap(o)

Next o

If available < qty Then shortage = "THIEU POSTING " & (qty - available) & " (CO " & available & "/CAN " & qty & "; TYPE=" & ModuleHelper.ToolConfig("POST_ALLOWED_STOCK_TYPES", "FU") & ")"

If available < qty And PendingMaterial(mat, requiredBatch) Then shortage = "POSTING KHA DUNG " & available & "/CAN " & qty & "; CON UNKNOWN GIU SO LUONG - XEM LEDGER/POSTING J"

If available < qty Then
 If Not mBlockedStockByMat Is Nothing Then
  If mBlockedStockByMat.exists(mat) Then
   shortage = shortage & "; (CO TON BI KHOA/QI/QU - XEM POSTING J)"
  End If
 End If
End If

If cap < qty Then

 If shortage <> "" Then shortage = shortage & "; "

 If cap = 0 Then
  shortage = shortage & "THIEU OB (VL06O) - Nhan nut 1 de tao lai qua VL10D hoac mo VL10D kiem tra"
 Else
  shortage = shortage & "THIEU OB (VL06O) " & (qty - cap) & " (CO " & cap & "/CAN " & qty & ") - Nhan nut 1 de tao lai qua VL10D hoac mo VL10D kiem tra"
 End If

End If

ModuleLog.WriteLog "CAPACITY", mat & "; Batch=" & requiredBatch & "; POSTING=" & available & "; OB=" & cap & "; Need=" & qty

If shortage <> "" Then
 ws.Cells(req, 4).Value2 = "WAIT_FULL_BATCH_CAPACITY: " & shortage
 ModuleLog.WriteLog "WAIT_FULL_BATCH_CAPACITY", mat & "; batch=" & requiredBatch & "; need=" & qty & "; done=" & CompletedQty(req) & "; " & shortage
 Exit Sub
End If

need = qty

For Each b In grouped.keys

 Set bRows = grouped(b): batchTotal = 0

 For Each k In bRows: batchTotal = batchTotal + mRemain(CLng(k)): Next k

 If batchTotal > need Then
  ModuleLog.WriteLog "WAIT_FULL_BATCH_CAPACITY", mat & "; batch=" & b & "; full remaining batch=" & batchTotal & "; requested remaining=" & need
  GoTo NextBatch
 End If

 For Each k In bRows

  ix = CLng(k): available = mRemain(ix)

  For o = 1 To obN

   If StrComp(CStr(mOB(o, 3)), mat, vbTextCompare) = 0 And mOBCap(o) > 0 Then

    obNeed = mOBCap(o): If localOB.exists(CStr(o)) Then obNeed = obNeed - CDbl(localOB(CStr(o)))

    take = Application.Min(available, need, obNeed)

    If take > 0 Then

     localAlloc.Add Array(req, CStr(mOB(o, 1)), CStr(mOB(o, 2)), mat, CStr(b), CStr(mStock(ix, 4)), take, CStr(mOB(o, 8)), CStr(mOB(o, 9)), CDbl(mOB(o, 4)), CStr(mOB(o, 7)))

     If Not localStock.exists(CStr(ix)) Then localStock(CStr(ix)) = 0#

     If Not localOB.exists(CStr(o)) Then localOB(CStr(o)) = 0#

     localStock(CStr(ix)) = CDbl(localStock(CStr(ix))) + take

     localOB(CStr(o)) = CDbl(localOB(CStr(o))) + take

     available = available - take: need = need - take

    End If

   End If

   If available <= 0 Or need <= 0 Then Exit For

  Next o

  If need <= 0 Then Exit For

 Next k

 If need <= 0 Then Exit For

NextBatch:

Next b

If need > 0 Then
 ws.Cells(req, 4).Value2 = "WAIT_FULL_BATCH_CAPACITY / KHONG DU NGUYEN BATCH " & need
 If Not mBlockedStockByMat Is Nothing Then
  If mBlockedStockByMat.exists(mat) Then
   ws.Cells(req, 4).Value2 = ws.Cells(req, 4).Value2 & "; (CO TON BI KHOA/QI/QU - XEM POSTING J)"
  End If
 End If
 Exit Sub
End If

For Each k In localStock.keys: mRemain(CLng(k)) = mRemain(CLng(k)) - CDbl(localStock(k)): Next k

For Each k In localOB.keys: mOBCap(CLng(k)) = mOBCap(CLng(k)) - CDbl(localOB(k)): Next k

For Each sr In localAlloc
 mWholeNeeds(CStr(req) & "|" & CStr(sr(4))) = BatchPlannedQty(localAlloc, req, CStr(sr(4))) + CompletedBatchQty(req, CStr(sr(4)))
 AssertAllocationStock sr

 mAlloc.Add sr

Next sr

ws.Cells(req, 4).Value2 = "READY " & qty

End Sub

Private Sub WritePlan()

Dim ws As Worksheet, ui As Worksheet, a As Variant, r As Long, itemMap As Object, batches As Object, key As String, k As Variant

Dim result As Variant, b As Variant, o As Long, selected() As Variant, i As Long

Set ws = ThisWorkbook.Worksheets("BATCH_PLAN"): Set ui = ThisWorkbook.Worksheets("INPUT")

ws.Range("A6:L" & CStr(Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row))).ClearContents

Set itemMap = CreateObject("Scripting.Dictionary")

r = 6

For Each a In mAlloc

 ws.Cells(r, 1).Resize(1, 12).Value2 = Array(RunID, a(0), a(1), a(2), a(3), a(4), a(5), a(6), ui.Range("R6").Value2, a(7), a(8), "READY")

 key = a(1) & "|" & a(2)

 If Not itemMap.exists(key) Then

  Set batches = CreateObject("Scripting.Dictionary")

  itemMap.Add key, Array(a(1), a(2), a(3), 0#, batches)

 End If

 result = itemMap(key): result(3) = CDbl(result(3)) + CDbl(a(6))

 Set batches = result(4)

 If Not batches.exists(CStr(a(4))) Then batches(CStr(a(4))) = 0#

 batches(CStr(a(4))) = CDbl(batches(CStr(a(4)))) + CDbl(a(6)): itemMap(key) = result

 r = r + 1

Next a

For i = 6 To 205

 If CStr(ui.Cells(i, 1).Value2) <> "" And Left$(CStr(ui.Cells(i, 4).Value2), 5) <> "READY" And Left$(CStr(ui.Cells(i, 4).Value2), 4) <> "DONE" Then

  ws.Cells(r, 1).Resize(1, 12).Value2 = Array(RunID, i, "", "", ui.Cells(i, 1).Value2, ui.Cells(i, 3).Value2, "", ui.Cells(i, 2).Value2, ui.Range("R6").Value2, "", "", ui.Cells(i, 4).Value2)

  r = r + 1

 End If

Next i

r = 6

For Each k In itemMap.keys

 result = itemMap(k): Set batches = result(4)

 For o = 1 To UBound(mOBCap)

  If CStr(mOB(o, 1)) = CStr(result(0)) And CStr(mOB(o, 2)) = CStr(result(1)) Then Exit For

 Next o

 If o > UBound(mOBCap) Then Err.Raise vbObjectError + 4107, , "Thieu OB/item nguon trong ke hoach " & CStr(k)

 For Each b In batches.keys

  If r > 505 Then Err.Raise vbObjectError + 4107, , "Qua 500 dong batch ket qua. Giam list dau vao."

  ui.Cells(r, 6).Resize(1, 9).Value2 = Array(result(1), result(0), result(2), CDbl(batches(b)), CStr(b), mOB(o, 4), IIf(batches.count > 1, "BATCH_SPLIT", "DIRECT"), "READY", "PRESERVE OTHER ITEMS; WHOLE BATCH GATE")

  r = r + 1

 Next b

Next k

If r > 6 Then ui.Range("F6:N" & CStr(r - 1)).Sort Key1:=ui.Range("F6"), Order1:=xlAscending, Key2:=ui.Range("G6"), Order2:=xlAscending, Key3:=ui.Range("H6"), Order3:=xlAscending, header:=xlNo, DataOption1:=xlSortTextAsNumbers, DataOption2:=xlSortTextAsNumbers

If r > 6 Then ui.Range("F6:N" & CStr(r - 1)).rows.AutoFit

ModuleLog.WriteLog "PLAN", CStr(itemMap.count) & " item; " & CStr(mAlloc.count) & " allocations"

ReDim selected(1 To UBound(mRemain), 1 To 1)

For i = 1 To UBound(mRemain): selected(i, 1) = mInitialRemain(i) - mRemain(i): Next i

ThisWorkbook.Worksheets("POSTING").Range("H6").Resize(UBound(mRemain), 1).Value2 = selected

End Sub

Private Function PlanPolicy() As String

PlanPolicy = "RC23_MANDATORY_WHOLE|" & ModuleHelper.ToolConfig("EXPECTED_UOM", "UN") & "|" & ModuleHelper.ToolConfig("VL06O_SHIPPING_POINT", "3093") & "|" & ModuleHelper.ToolConfig("POST_ALLOWED_STOCK_TYPES", "FU") & "|" & ModuleHelper.ToolConfig("POST_BLOCKED_STOCK_TYPES", "QT,QI,QU") & "|" & ModuleHelper.ToolConfig("VL06O_ALLOW_BLANK_SLOC", "TRUE")

End Function

Private Function PlanViewFingerprint() As String

Dim ws As Worksheet, a As Variant, r As Long, c As Long, n As Long, s As String, t As String, pass As Long

For pass = 1 To 2

 If pass = 1 Then

  Set ws = ThisWorkbook.Worksheets("BATCH_PLAN")

  n = Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)

  a = ws.Range("A6:K" & CStr(n)).Value2

 Else

  Set ws = ThisWorkbook.Worksheets("INPUT")

  a = ws.Range("F6:L505").Value2

 End If

 For r = 1 To UBound(a, 1)

  For c = 1 To UBound(a, 2)

   t = CStr(a(r, c)): s = s & Len(t) & ":" & t

  Next c

 Next r

Next pass

PlanViewFingerprint = s

End Function

Private Sub QueueRecovery(ByVal target As Object, ByVal mat As String, ByVal refreshStock As Boolean)

If target.exists(mat) Then

 target(mat) = CBool(target(mat)) Or refreshStock

Else

 target(mat) = refreshStock

End If

End Sub

Public Sub OB_Execute()

If Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2)) = "2001" Then
    OB2001_Execute
    Exit Sub
End If

Dim groups As Object, a As Variant, k As Variant, items As Collection, mode As String, maxOB As Long, done As Long, msg As String, started As Double, failedText As String, retryMaterials As Object, material As Variant, first As Variant, retryGroups As Object, freshBlocked As Object, deferred As Object, heldMaterials As Object, rr As Long, pendingQty As Double, retryPass As Long, maxPasses As Long, nextDeferred As Object, waitSec As Long, waitDuration As Long

If mRunning Then Exit Sub

If Not mChecked Then

 MsgBox "Chua co ke hoach da kiem tra trong phien nay. Nhan nut 1: KIEM TRA & LAP KE HOACH truoc.", vbExclamation

 Exit Sub

End If

On Error GoTo Fail

AssertRunIdentity

If PlanPolicy() <> mPlanPolicy Then Err.Raise vbObjectError + 4112, , "Cau hinh lap ke hoach da thay doi. Nhan nut 1 de lap lai ke hoach."

If PlanViewFingerprint() <> mPlanViewSignature Then Err.Raise vbObjectError + 4112, , "Bang ke hoach da bi sua. Nhan nut 1 de kiem tra lai."

ValidateInput

mode = UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW"))

If mode <> "PREVIEW" And mode <> "APPLY" Then Err.Raise vbObjectError + 4111, , "MODE chi nhan PREVIEW/APPLY."

If mode <> "APPLY" Then

 MsgBox "MODE=PREVIEW: chua ghi SAP. Neu thieu OB, doi APPLY va nhan lai nut 1 truoc. Nut 2 chi dien OB da co. TEST_SCOPE=FIRST thu 1 OB; ALL xu ly ca ke hoach.", vbInformation

 Exit Sub

End If

If mAlloc.count > 500 Then Err.Raise vbObjectError + 4101, , "Dung luong ke hoach (" & mAlloc.count & " dong) vuot qua gioi han 500 dong (F6:N505). Dung truoc khi APPLY."

If Not BeginRun("THUC HIEN OB VL02N", True) Then Exit Sub
ModuleLog.WriteLog "EXECUTE_OB_ONLY", "Use verified OB plan; no VL10D creation in button 2"

mMissingOBWarnUser = False
Set retryMaterials = CreateObject("Scripting.Dictionary")

Set deferred = CreateObject("Scripting.Dictionary"): Set heldMaterials = CreateObject("Scripting.Dictionary")

If mExcludedOB Is Nothing Then Set mExcludedOB = CreateObject("Scripting.Dictionary")

Set freshBlocked = FreshSourceBlocks()
' NOTE: mEDCWaitByMat is for future VL10D creation waits.
' It must NEVER block allocations of already-existing, valid candidate OBs!
For Each material In freshBlocked.keys: heldMaterials(CStr(material)) = True: Next material

Set groups = CreateObject("Scripting.Dictionary")

For Each a In mAlloc

 If Not freshBlocked.exists(CStr(a(3))) Then

 If Not groups.exists(CStr(a(1))) Then Set items = New Collection: groups.Add CStr(a(1)), items

 groups(CStr(a(1))).Add a

 End If

Next a

If groups.count = 0 And UCase$(ModuleHelper.ToolConfig("TEST_SCOPE", "FIRST")) <> "ALL" Then WriteFinalBalance heldMaterials: EndRun: MsgBox "Khong co OB READY; xem NOTE / SOURCE_STOCK_BLOCK.", vbInformation: Exit Sub

If UCase$(ModuleHelper.ToolConfig("TEST_SCOPE", "FIRST")) = "ALL" Then maxOB = groups.count Else maxOB = 1
If maxOB = 1 And UCase$(ModuleHelper.ToolConfig("TEST_SCOPE", "FIRST")) <> "ALL" Then
 If groups.count > 0 Then
  first = groups.keys: Set items = groups(CStr(first(0)))
  If Not FirstScopeWhole(items) Then
   ModuleLog.WriteLog "WAIT_FULL_BATCH_CAPACITY", "TEST_SCOPE=FIRST would split a whole batch across multiple OBs. Use ALL."
EndRun:    MsgBox "Batch can nhieu OB. Doi TEST_SCOPE=ALL de xu ly nguyen Batch.", vbExclamation: Exit Sub
  End If
 End If
End If

For Each k In groups.keys

 If Fingerprint() <> mSignature Then Err.Raise vbObjectError + 4108, , "INPUT thay doi trong luc chay. Dung truoc OB tiep theo."

 started = Timer: Set items = groups(k)

 If ApplyOneOB(CStr(k), items, msg) Then

  done = done + 1

  ModuleLog.WriteLog "DONE", msg & "; sec=" & Format$(Elapsed(started), "0.00"), CStr(k)

 Else

  If Left$(msg, 5) = "WAIT_" Then

   If Left$(msg, 24) <> "WAIT_FULL_BATCH_CAPACITY" Then deferred(CStr(k)) = True

  Else

   For Each a In items: QueueRecovery retryMaterials, CStr(a(3)), (ModuleSAP.LastConflictKind <> "MISSING"): Next a

  End If

 End If

 If done >= maxOB Then Exit For

Next k

If UCase$(ModuleHelper.ToolConfig("TEST_SCOPE", "FIRST")) = "ALL" Or done < maxOB Then

 If deferred.count > 0 Then
  maxPasses = 2
  
  For retryPass = 1 To maxPasses
   If deferred.count = 0 Then Exit For
   If done >= maxOB And UCase$(ModuleHelper.ToolConfig("TEST_SCOPE", "FIRST")) <> "ALL" Then Exit For
   
   waitDuration = IIf(retryPass = 1, 6, 5)
   ModuleLog.WriteLog "WAIT_LOCK_BACKOFF", "Co " & deferred.count & " OB xep sau (bi khoa/DO_AUTO). Dang cho SAP nha khoa " & waitDuration & "s (nhip " & retryPass & "/" & maxPasses & ")..."
   For waitSec = waitDuration To 1 Step -1
    Application.StatusBar = "SAP DANG XU LY NGAM (DO_AUTO). DANG CHO " & waitSec & "s DE HE THONG NHA KHOA (NHIP " & retryPass & "/" & maxPasses & ")..."
    ModuleHelper.SleepMS 1000
    DoEvents
   Next waitSec
   Application.StatusBar = "DANG THU LAI CAC OB XEP SAU (NHIP " & retryPass & ")..."
   
   Set nextDeferred = CreateObject("Scripting.Dictionary")
   
   For Each k In deferred.keys
    If done >= maxOB And UCase$(ModuleHelper.ToolConfig("TEST_SCOPE", "FIRST")) <> "ALL" Then Exit For
    Set items = groups(k)
    started = Timer
    
    If ApplyOneOB(CStr(k), items, msg) Then
     done = done + 1
     ModuleLog.WriteLog "DONE_DEFERRED", msg & "; pass=" & retryPass & "; sec=" & Format$(Elapsed(started), "0.00"), CStr(k)
    ElseIf Left$(msg, 5) = "WAIT_" Then
     If retryPass < maxPasses Then
      nextDeferred(CStr(k)) = True
      ModuleLog.WriteLog "RETRY_DEFERRED_WAIT", "Con bi khoa SAP; xep sau cho nhip " & (retryPass + 1) & "; " & msg, CStr(k)
     Else
      For Each a In items: heldMaterials(CStr(a(3))) = True: Next a
      ModuleLog.WriteLog "DEFER_FINAL_FAIL", "Het luot thu lai; " & msg, CStr(k)
     End If
    Else
     For Each a In items: QueueRecovery retryMaterials, CStr(a(3)), (ModuleSAP.LastConflictKind <> "MISSING"): Next a
    End If
   Next k
   
   Set deferred = nextDeferred
  Next retryPass
  Application.StatusBar = False
 End If

 For rr = 6 To 205

  material = ModuleHelper.CleanStr(ThisWorkbook.Worksheets("INPUT").Cells(rr, 1).Value2)

  If CStr(material) <> "" And Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Cells(rr, 2).Value2)) <> "" Then

   pendingQty = ModuleHelper.CleanDouble(ThisWorkbook.Worksheets("INPUT").Cells(rr, 2).Value2) - CompletedQty(rr)

   If pendingQty > 0 And CompletedQty(rr) > 0 Then QueueRecovery retryMaterials, CStr(material), True

  End If

 Next rr

 PrepareRecoveryBatch retryMaterials, heldMaterials
 For Each material In retryMaterials.keys

  If Not heldMaterials.exists(CStr(material)) Then

  ReplanOneMaterial CStr(material), CBool(retryMaterials(CStr(material)))

  Set retryGroups = CreateObject("Scripting.Dictionary")

  For Each a In mAlloc

   If StrComp(CStr(a(3)), CStr(material), vbTextCompare) = 0 And Not IsExcludedOB(CStr(a(1))) Then

    If Not retryGroups.exists(CStr(a(1))) Then Set items = New Collection: retryGroups.Add CStr(a(1)), items

    retryGroups(CStr(a(1))).Add a

   End If

  Next a

  For Each k In retryGroups.keys

AssertRunIdentity

    Set items = retryGroups(k)

   If ApplyOneOB(CStr(k), items, msg) Then

    done = done + 1

   ElseIf Left$(msg, 5) = "WAIT_" Then

    For Each a In items: heldMaterials(CStr(a(3))) = True: Next a

   End If

  Next k

  End If

 Next material
 mRecoveryPrepared = False: Set mRecoveryProducts = Nothing: Set mRecoveryCapMap = Nothing

End If

WriteFinalBalance heldMaterials

mChecked = False

If mMissingOBWarnUser Then
 ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = "DA LUU " & done & " OB; THIEU OB - NHAN NUT 1 TAO LAI VL10D"
 PersistCheckpoint "EXCEL_SAVE_FINAL_SEC"
 EndRun
 MsgBox "Da Save " & done & " OB. CANH BAO: Thieu OB trong VL06O (OB cu bi loi/xoa). Vui long nhan nut 1 (KIEM TRA) de tao lai qua VL10D hoac mo VL10D kiem tra!", vbExclamation
Else
 ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = "DA LUU " & done & " OB; XEM NOTE/LOG"
 PersistCheckpoint "EXCEL_SAVE_FINAL_SEC"
 EndRun
 MsgBox "Da Save " & done & " OB. Nhan KIEM TRA de xem phan con lai. Da bao toan item ngoai ke hoach.", vbInformation
End If

Exit Sub

Fail:

failedText = Err.Description

If Not items Is Nothing And ModuleSAP.LastApplyState <> "DONE" Then

 On Error Resume Next

 If ModuleSAP.LastApplyState = "" Then ModuleSAP.LastApplyState = "UNKNOWN"

 Journal items, ModuleSAP.LastApplyState, failedText

 ModuleLog.WriteLog "EXECUTE_FAIL", failedText

 PersistCheckpoint

 On Error GoTo 0

End If

' RC20-DEF-10: Update final balance and restore settings on failure

On Error Resume Next

mRecoveryPrepared = False: Set mRecoveryProducts = Nothing: Set mRecoveryCapMap = Nothing
If Not heldMaterials Is Nothing Then WriteFinalBalance heldMaterials

ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = "LOI EXECUTE; DA LUU " & done & " OB; XEM NOTE/LOG"

PersistCheckpoint "EXCEL_SAVE_ERROR_SEC"

RestoreApplicationSettings

On Error GoTo 0

EndRun

ReportError "EXECUTE", failedText

End Sub

Private Sub Journal(ByVal items As Collection, ByVal status As String, ByVal message As String)

Dim ws As Worksheet, a As Variant, r As Long, found As Long, key As String, first As Variant

Set ws = ThisWorkbook.Worksheets("LEDGER")

For Each a In items

 key = FrozenRequestKey(CLng(a(0))) & "|" & a(1) & "|" & a(2) & "|" & a(4) & "|" & a(5)

 found = 0

 For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

  If CStr(ws.Cells(r, 8).Value2) = key Then found = r: Exit For

 Next r

 If found = 0 Then found = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1

 ws.Cells(found, 1).Resize(1, 10).Value2 = Array(FrozenRequestKey(CLng(a(0))), RunID, a(1), a(2), a(6), status, Now, key, message, mSignature)

 ws.Cells(found, 11).Resize(1, 6).Value2 = Array(a(3), a(4), a(5), a(9), a(10), ThisWorkbook.Worksheets("INPUT").Range("R6").Value2)

 With ws.Cells(found, 1).Resize(1, 16).Borders

  .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)

 End With

Next a

Dim ui As Worksheet, plan As Worksheet

first = items(1)

Set ui = ThisWorkbook.Worksheets("INPUT"): Set plan = ThisWorkbook.Worksheets("BATCH_PLAN")

For r = 6 To 505

 If CStr(ui.Cells(r, 7).Value2) = CStr(first(1)) Then

  ui.Cells(r, 13).Value2 = status: ui.Cells(r, 14).Value2 = IIf(status = "DONE", "SAVED " & CStr(first(1)), message)

 End If

Next r

For r = 6 To plan.Cells(plan.rows.count, 1).End(xlUp).row

 If CStr(plan.Cells(r, 3).Value2) = CStr(first(1)) Then plan.Cells(r, 12).Value2 = status

Next r

End Sub

Public Sub FormatOperationalTables()

Dim ws As Worksheet, lastRow As Long, tableCols As Long, firstRow As Long, sheetName As Variant

For Each sheetName In Array("INPUT", "POSTING", "VL06O_DATA", "LOG", "CONFIG", "BATCH_PLAN", "POSTING_RAW", "SEARCH_LIST", "LEDGER")

 Set ws = ThisWorkbook.Worksheets(CStr(sheetName))

 firstRow = 5

 Select Case CStr(sheetName)

 Case "INPUT"

  With ws.Range("A5:D205").Borders: .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225): End With

  lastRow = Application.Max(6, ws.Cells(ws.rows.count, 7).End(xlUp).row): tableCols = 14

  With ws.Range("F5:N" & lastRow).Borders: .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225): End With

  tableCols = 0

 Case "POSTING": tableCols = 10

 Case "VL06O_DATA": tableCols = 16

 Case "LOG": tableCols = 6

 Case "CONFIG": tableCols = 3: firstRow = 1

 Case "BATCH_PLAN": tableCols = 12

 Case "POSTING_RAW": tableCols = 17: firstRow = 1

 Case "SEARCH_LIST": tableCols = 1: firstRow = 1

 Case "LEDGER": tableCols = 16: firstRow = 1

 End Select

 If tableCols > 0 Then

  lastRow = Application.Max(firstRow + 1, ws.Cells(ws.rows.count, 1).End(xlUp).row)

  With ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, tableCols)).Borders

   .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)

  End With

 End If

Next sheetName

End Sub

Private Sub PersistCheckpoint(Optional ByVal timingEvent As String = "", Optional ByVal ob As String = "")

Dim started As Double

started = Timer

If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 4109, , "File dang ReadOnly. Can luu duoc ket qua de tranh chay lap."

ThisWorkbook.Save

If timingEvent <> "" Then ModuleLog.WriteLog timingEvent, Format$(Elapsed(started), "0.00"), ob

End Sub

Private Function Elapsed(ByVal t As Double) As Double

If Timer >= t Then Elapsed = Timer - t Else Elapsed = 86400# - t + Timer

End Function

Private Sub ReportError(ByVal stage As String, ByVal message As String)

mChecked = False

ModuleLog.WriteLog "ERROR", stage & ": " & message

ThisWorkbook.Worksheets("INPUT").Range("R22").Value2 = "ERROR " & stage

EndRun

MsgBox stage & ": " & message & vbCrLf & "Neu SAP dang o VL02N, kiem tra man hinh va LOG. Khong tu dong retry Save/Create.", vbCritical

End Sub

Private Sub EnsureOptimizationConfig()

Dim key As Variant, keys As Variant, defaults As Variant, i As Long

keys = Array("VL02N_SESSION_REUSE", "RECOVERY_FAST_OB_ONLY", "ALLOW_DELETE_UNPLANNED_ITEMS", "POSTSAVE_VERIFY_LIVE", "FAST_SAVE_MODE", "STRICT_AUDIT")

defaults = Array("TRUE", "FALSE", "TRUE", "TRUE", "TRUE", "FALSE")

For i = 0 To UBound(keys)

 If ModuleHelper.ToolConfig(CStr(keys(i)), "__MISSING__") = "__MISSING__" Then PutConfig CStr(keys(i)), CStr(defaults(i))

Next i

End Sub

Public Sub OB_OpenConfig()

Dim ws As Worksheet, r As Long

EnsureOptimizationConfig

Set ws = ThisWorkbook.Worksheets("CONFIG")

ws.Activate

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If CStr(ws.Cells(r, 1).Value2) = "MODE" Then ws.Cells(r, 2).Select: Exit Sub

Next r

End Sub

Public Sub OB_SetupButtons()

Dim ws As Worksheet, Sh As Shape, names As Variant, actions As Variant, rows As Variant, i As Long, p As Range

EnsureOptimizationConfig

Set ws = ThisWorkbook.Worksheets("INPUT")

On Error Resume Next

ws.Range("P9:R21").UnMerge

ws.Range("P9:R21").ClearContents

ws.Range("P9:R21").Interior.ColorIndex = xlNone

ws.Range("P9:R21").Borders.LineStyle = xlNone

For i = ws.Shapes.count To 1 Step -1

 If Left$(ws.Shapes(i).Name, 6) = "OBV1__" Then ws.Shapes(i).Delete

Next i

On Error GoTo 0

names = Array("1. KI" & ChrW(7874) & "M TRA & L" & ChrW(7852) & "P K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf & "4007: APPLY TAO OB + FIFO", "2. TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf & "VL02N / " & ModuleHelper.ToolConfig("TEST_SCOPE", "ALL"))

actions = Array("OB_Check", "OB_Execute")

rows = Array(9, 13)

For i = 0 To 1

 Set p = ws.Range("P" & rows(i) & ":R" & rows(i) + 2)

 Set Sh = ws.Shapes.AddShape(5, p.Left, p.Top, p.width, p.Height)

 Sh.Name = "OBV1__" & CStr(i): Sh.OnAction = "'" & ThisWorkbook.Name & "'!" & actions(i)

 Sh.Placement = xlMoveAndSize

 On Error Resume Next

 Sh.TextFrame.Characters.text = names(i)

 Sh.TextFrame.Characters.Font.Size = 14: Sh.TextFrame.Characters.Font.Bold = True

 Sh.TextFrame.HorizontalAlignment = xlHAlignCenter: Sh.TextFrame.VerticalAlignment = xlVAlignCenter

 Sh.TextFrame.Characters.Font.Color = vbWhite

 Sh.TextFrame2.TextRange.text = names(i)

 Sh.TextFrame2.TextRange.Font.Size = 14

 Sh.TextFrame2.TextRange.Font.Bold = -1

 Sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

 Sh.TextFrame2.HorizontalAnchor = 2

 Sh.TextFrame2.VerticalAnchor = 3

 Sh.TextFrame2.MarginLeft = 4

 Sh.TextFrame2.MarginRight = 4

 Sh.TextFrame2.MarginTop = 2

 Sh.TextFrame2.MarginBottom = 2

 Sh.TextFrame2.WordWrap = -1

 On Error GoTo 0

 Sh.Fill.ForeColor.RGB = RGB(34, 70, 101)

 If i = 0 Then Sh.Fill.ForeColor.RGB = RGB(15, 118, 110)

 If i = 1 Then Sh.Fill.ForeColor.RGB = RGB(166, 94, 15)

 Sh.line.Visible = False

Next i

OB_UpdateModeButton

End Sub

Public Sub OB_UpdateModeButton()

Dim Sh As Shape, firstButton As Shape, mode As String, txt As String

On Error Resume Next

Set Sh = ThisWorkbook.Worksheets("INPUT").Shapes("OBV1__1")

If Sh Is Nothing Then Exit Sub

mode = UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW"))
Set firstButton = ThisWorkbook.Worksheets("INPUT").Shapes("OBV1__0")
If Not firstButton Is Nothing Then
 If mode = "APPLY" Then
  firstButton.TextFrame.Characters.text = "1. KIEM TRA + LAP KE HOACH" & vbLf & "4007: TAO OB THIEU / APPLY"
 Else
  firstButton.TextFrame.Characters.text = "1. KIEM TRA + LAP KE HOACH" & vbLf & "PREVIEW: CHI KIEM TRA"
 End If
 firstButton.TextFrame2.TextRange.text = firstButton.TextFrame.Characters.text
End If

If mode = "APPLY" Then

 txt = "2. TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf & "VL02N / " & ModuleHelper.ToolConfig("TEST_SCOPE", "FIRST")

Else

 txt = "2. TH" & ChrW(7920) & "C HI" & ChrW(7878) & "N K" & ChrW(7870) & " HO" & ChrW(7840) & "CH" & vbLf & "MODE: PREVIEW - Ch" & ChrW(432) & "a ghi SAP"

End If

If mChecked Then Sh.Fill.ForeColor.RGB = RGB(166, 94, 15) Else Sh.Fill.ForeColor.RGB = RGB(100, 116, 139)

Sh.TextFrame.Characters.text = txt

Sh.TextFrame.Characters.Font.Size = 14: Sh.TextFrame.Characters.Font.Bold = True

Sh.TextFrame.HorizontalAlignment = xlHAlignCenter: Sh.TextFrame.VerticalAlignment = xlVAlignCenter

Sh.TextFrame.Characters.Font.Color = vbWhite

Sh.TextFrame2.TextRange.text = txt

Sh.TextFrame2.TextRange.Font.Size = 14

Sh.TextFrame2.TextRange.Font.Bold = -1

Sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

Sh.TextFrame2.HorizontalAnchor = 2

Sh.TextFrame2.VerticalAnchor = 3

End Sub

Private Sub WriteFinalBalance(ByVal held As Object)
Dim ws As Worksheet, r As Long, mat As String, need As Double, doneQty As Double, remaining As Double, k As Variant, parts() As String
Set ws = ThisWorkbook.Worksheets("INPUT")
For r = 6 To 205
 mat = ModuleHelper.CleanStr(ws.Cells(r, 1).Value2)
 If mat <> "" Then
  doneQty = CompletedQty(r): need = 0
  If Trim$(CStr(ws.Cells(r, 2).Value2)) <> "" Then
   need = ModuleHelper.CleanDouble(ws.Cells(r, 2).Value2)
  ElseIf Not mWholeNeeds Is Nothing Then
   For Each k In mWholeNeeds.keys
    parts = Split(CStr(k), "|")
    If parts(0) = CStr(r) Then need = need + CDbl(mWholeNeeds(k))
   Next k
  End If
  If need = 0 And doneQty = 0 And Trim$(CStr(ws.Cells(r, 2).Value2)) = "" Then GoTo NextBalanceRow
  remaining = Application.Max(0, need - doneQty)
  ws.Cells(r, 4).Value2 = "CAN " & need & " / DONE " & doneQty & " / CON " & remaining
  If remaining > 0 Then
   ws.Cells(r, 4).Value2 = "WAIT_FULL_BATCH_CAPACITY: " & ws.Cells(r, 4).Value2 & "; Batch " & ws.Cells(r, 3).Value2
   ModuleLog.WriteLog "WAIT_FULL_BATCH_CAPACITY", mat & "; batch=" & ws.Cells(r, 3).Value2 & "; need=" & need & "; done=" & doneQty & "; remaining=" & remaining
  End If
  If held.exists(mat) Then ws.Cells(r, 4).Value2 = ws.Cells(r, 4).Value2 & " / CHO OB HOAC TON NGUON - XEM LOG"
 End If
NextBalanceRow:
Next r
End Sub

Private Sub AssertAllocationStock(ByVal a As Variant)

Dim i As Long, eligible As Double, st As String

For i = 1 To UBound(mStock, 1)

 If CStr(mStock(i, 1)) = CStr(a(3)) And CStr(mStock(i, 2)) = CStr(a(4)) And CStr(mStock(i, 4)) = CStr(a(5)) Then

  st = UCase$(Trim$(CStr(mStock(i, 5))))

  If StockTypeInList(st, ModuleHelper.ToolConfig("POST_ALLOWED_STOCK_TYPES", "FU")) And Not StockTypeInList(st, ModuleHelper.ToolConfig("POST_BLOCKED_STOCK_TYPES", "QT,QI,QU")) Then eligible = eligible + mInitialRemain(i)

 End If

Next i

If eligible < CDbl(a(6)) Then Err.Raise vbObjectError + 4122, , "BLOCK SOURCE STOCK TYPE: " & CStr(a(3)) & "/" & CStr(a(4))

End Sub

Private Function FreshSourceBlocks() As Object

Dim required As Object, available As Object, matByKey As Object, blocked As Object

Dim a As Variant, k As Variant, raw As Worksheet, n As Long, data As Variant, i As Long, key As String, st As String

Set required = CreateObject("Scripting.Dictionary"): Set available = CreateObject("Scripting.Dictionary")

Set matByKey = CreateObject("Scripting.Dictionary"): Set blocked = CreateObject("Scripting.Dictionary")

If ModuleHelper.ToolConfigBool("VERIFY_SOURCE_BEFORE_RUN", True) Then

 AssertRunIdentity
 If SourceFingerprint() <> mSourceSignature Then
  ModuleLog.WriteLog "POSTING_REFRESH_REASON", "SOURCE_SNAPSHOT_CHANGED"
  Err.Raise vbObjectError + 4123, , "POSTING_RAW da bi sua. Nhan KIEM TRA de lap ke hoach moi."
 End If
 If mPlanCheckedAt = 0 Then Err.Raise vbObjectError + 4123, , "Thieu plan snapshot. Nhan KIEM TRA."
 ModuleLog.WriteLog "PLAN_SNAPSHOT_VALID", "Request + Input + Policy + Plan + Source matched"
 If DateDiff("s", mPlanCheckedAt, Now) > ModuleHelper.ToolConfigLong("PLAN_SOURCE_MAX_IDLE_SEC", 900, 60, 86400) Then
  ModuleLog.WriteLog "POSTING_REFRESH_REQUIRED", "SOURCE_UNCERTAIN; pending Material/Batch only"
  Set mRecoveryProducts = CreateObject("Scripting.Dictionary")
  For Each a In mAlloc
   mRecoveryProducts(CStr(a(3))) = True
  Next a
  If Not ModuleSAP.DownloadPostingForOrder(False) Then
   Set mRecoveryProducts = Nothing
   Err.Raise vbObjectError + 4123, , "SOURCE CHECK: " & ModuleSAP.LastPostingErrorText
  End If
  Set mRecoveryProducts = Nothing
  mSourceSignature = SourceFingerprint(): mPlanCheckedAt = Now
 Else
  ModuleLog.WriteLog "POSTING_CACHE_REUSE", "Frozen plan; age=" & DateDiff("s", mPlanCheckedAt, Now) & "; live OB preflight retained"
 End If

 Set raw = ThisWorkbook.Worksheets("POSTING_RAW")

 n = raw.Cells(raw.rows.count, 1).End(xlUp).row

 If n >= 2 Then

  data = raw.Range("A2:Q" & n).Value2

  For i = 1 To UBound(data, 1)

   st = UCase$(Trim$(CStr(data(i, 5))))

   If st = UCase$(Trim$(CStr(data(i, 10)))) And StockTypeInList(st, ModuleHelper.ToolConfig("POST_ALLOWED_STOCK_TYPES", "FU")) And Not StockTypeInList(st, ModuleHelper.ToolConfig("POST_BLOCKED_STOCK_TYPES", "QT,QI,QU")) Then

    key = CStr(data(i, 1)) & "|" & CStr(data(i, 3)) & "|" & CStr(data(i, 4))

    If Not available.exists(key) Then available(key) = 0#

    available(key) = CDbl(available(key)) + ModuleHelper.CleanDouble(data(i, 2))

   End If

  Next i

 End If

 For Each a In mAlloc

  key = CStr(a(3)) & "|" & CStr(a(4)) & "|" & CStr(a(5))

  If Not required.exists(key) Then required(key) = 0#

  required(key) = CDbl(required(key)) + CDbl(a(6)): matByKey(key) = CStr(a(3))

 Next a

 For Each k In required.keys

  If Not available.exists(CStr(k)) Then available(k) = 0#

  If CDbl(available(k)) < CDbl(required(k)) Then

   blocked(CStr(matByKey(k))) = True

   ModuleLog.WriteLog "SOURCE_STOCK_BLOCK", CStr(k) & "; allowed=" & available(k) & "; planned=" & required(k)

  End If

 Next k

End If

Set FreshSourceBlocks = blocked

End Function





Private Function ReservedByOtherMaterial(ByVal ob As String, ByVal mat As String) As Boolean
Dim a As Variant
If mat = "" Then Exit Function
For Each a In mAlloc
 If CStr(a(1)) = ob And StrComp(CStr(a(3)), mat, vbTextCompare) <> 0 Then ReservedByOtherMaterial = True: Exit Function
Next a
End Function

Private Function CompletedBatchQty(ByVal req As Long, ByVal batch As String) As Double
Dim ws As Worksheet, r As Long
Set ws = ThisWorkbook.Worksheets("LEDGER")
For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 If CStr(ws.Cells(r, 1).Value2) = requestKey(req) And CStr(ws.Cells(r, 6).Value2) = "DONE" And StrComp(CStr(ws.Cells(r, 12).Value2), batch, vbTextCompare) = 0 Then CompletedBatchQty = CompletedBatchQty + ModuleHelper.CleanDouble(ws.Cells(r, 5).Value2)
Next r
End Function

Private Function BatchPlannedQty(ByVal entries As Collection, ByVal req As Long, ByVal batch As String) As Double
Dim a As Variant
For Each a In entries
 If CLng(a(0)) = req And StrComp(CStr(a(4)), batch, vbTextCompare) = 0 And Not IsExcludedOB(CStr(a(1))) Then BatchPlannedQty = BatchPlannedQty + CDbl(a(6))
Next a
End Function

Public Function WholeBatchGate(ByVal items As Collection, ByRef message As String) As Boolean
Dim a As Variant, key As String, need As Double, planned As Double, finished As Double, checked As Object
Set checked = CreateObject("Scripting.Dictionary")
For Each a In items
 key = CStr(a(0)) & "|" & CStr(a(4))
 If Not checked.exists(key) Then
  checked(key) = True
  If mWholeNeeds Is Nothing Then message = "WAIT_FULL_BATCH_CAPACITY: missing batch target": Exit Function
  If Not mWholeNeeds.exists(key) Then message = "WAIT_FULL_BATCH_CAPACITY: missing target " & key: Exit Function
  need = CDbl(mWholeNeeds(key)): finished = CompletedBatchQty(CLng(a(0)), CStr(a(4)))
  planned = BatchPlannedQty(mAlloc, CLng(a(0)), CStr(a(4)))
  ' mAlloc includes already saved allocations; cap against persisted completion once.
  planned = planned - PlannedDoneQty(CLng(a(0)), CStr(a(4)))
  If planned <> need - finished Or PendingMaterial(CStr(a(3)), CStr(a(4))) Then
   message = "WAIT_FULL_BATCH_CAPACITY: " & a(3) & "; batch=" & a(4) & "; need=" & need & "; done=" & finished & "; remaining=" & Application.Max(0, need - finished) & "; planned=" & planned
   ModuleLog.WriteLog "WAIT_FULL_BATCH_CAPACITY", message, CStr(a(1))
   Exit Function
  End If
 End If
Next a
WholeBatchGate = True
End Function

Private Function PlannedDoneQty(ByVal req As Long, ByVal batch As String) As Double
Dim a As Variant, ws As Worksheet, r As Long, key As String
Set ws = ThisWorkbook.Worksheets("LEDGER")
For Each a In mAlloc
 If CLng(a(0)) = req And StrComp(CStr(a(4)), batch, vbTextCompare) = 0 And Not IsExcludedOB(CStr(a(1))) Then
  key = requestKey(req) & "|" & a(1) & "|" & a(2) & "|" & a(4) & "|" & a(5)
  For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
   If CStr(ws.Cells(r, 8).Value2) = key And CStr(ws.Cells(r, 6).Value2) = "DONE" Then PlannedDoneQty = PlannedDoneQty + CDbl(a(6)): Exit For
  Next r
 End If
Next a
End Function

Private Function FirstScopeWhole(ByVal items As Collection) As Boolean
Dim a As Variant, key As String
For Each a In items
 key = CStr(a(0)) & "|" & CStr(a(4))
 If BatchPlannedQty(items, CLng(a(0)), CStr(a(4))) < CDbl(mWholeNeeds(key)) - CompletedBatchQty(CLng(a(0)), CStr(a(4))) Then Exit Function
Next a
FirstScopeWhole = True
End Function

Private Function RecoveryCapsFor(ByVal mat As String, ByVal rows As Variant) As Object
Dim caps As Object, ws As Worksheet, i As Long, batch As String, k As Variant, delta As Object, key As String
Set caps = CreateObject("Scripting.Dictionary"): caps.CompareMode = vbTextCompare
For i = 1 To UBound(rows, 1)
 If StrComp(CStr(rows(i, 1)), mat, vbTextCompare) = 0 Then
  batch = CStr(rows(i, 2))
  If Not caps.exists(batch) Then caps(batch) = 0#
  caps(batch) = CDbl(caps(batch)) + ModuleHelper.CleanDouble(rows(i, 3))
 End If
Next i
Set ws = ThisWorkbook.Worksheets("LEDGER"): Set delta = CreateObject("Scripting.Dictionary")
For i = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 If CStr(ws.Cells(i, 6).Value2) = "DONE" And StrComp(CStr(ws.Cells(i, 11).Value2), mat, vbTextCompare) = 0 Then
  batch = CStr(ws.Cells(i, 12).Value2)
  If Not delta.exists(batch) Then delta(batch) = 0#
  delta(batch) = CDbl(delta(batch)) + ModuleHelper.CleanDouble(ws.Cells(i, 5).Value2)
 End If
Next i
For Each k In caps.keys
 key = mat & "|" & CStr(k)
 If delta.exists(CStr(k)) Then
  If Not mDoneBaseline Is Nothing Then
   If mDoneBaseline.exists(key) Then delta(k) = CDbl(delta(k)) - CDbl(mDoneBaseline(key))
  End If
  caps(k) = Application.Max(0, CDbl(caps(k)) - Application.Max(0, CDbl(delta(k))))
 End If
Next k
Set RecoveryCapsFor = caps
End Function

Private Sub MergeUntouched(ByVal ws As Worksheet, ByVal oldRows As Variant, ByVal firstRow As Long, ByVal matCol As Long, ByVal cols As Long)
Dim i As Long, r As Long, c As Long, freshOB As Object, keepRow As Boolean, rowCopy() As Variant
Set freshOB = CreateObject("Scripting.Dictionary")
If matCol = 3 Then
 For i = firstRow To ws.Cells(ws.rows.count, 1).End(xlUp).row: freshOB(CStr(ws.Cells(i, 1).Value2)) = True: Next i
End If
r = Application.Max(firstRow, ws.Cells(ws.rows.count, 1).End(xlUp).row + 1)
For i = 1 To UBound(oldRows, 1)
 keepRow = (CStr(oldRows(i, 1)) <> "" And Not mRecoveryProducts.exists(CStr(oldRows(i, matCol))))
 If matCol = 3 And freshOB.exists(CStr(oldRows(i, 1))) Then keepRow = False
 If keepRow Then
  ReDim rowCopy(1 To 1, 1 To cols)
   For c = 1 To Application.Min(cols, UBound(oldRows, 2)): rowCopy(1, c) = oldRows(i, c): Next c
   ws.Cells(r, 1).Resize(1, cols).Value2 = rowCopy: r = r + 1
 End If
Next i
End Sub

Private Sub PrepareRecoveryBatch(ByVal queued As Object, ByVal held As Object)
Dim mat As Variant, stockRefresh As Boolean, postws As Worksheet, obws As Worksheet, postRows As Variant, obRows As Variant, caps As Object
Set mRecoveryProducts = CreateObject("Scripting.Dictionary"): mRecoveryProducts.CompareMode = vbTextCompare
Set mRecoveryCapMap = CreateObject("Scripting.Dictionary"): mRecoveryCapMap.CompareMode = vbTextCompare
mRecoveryPrepared = False
For Each mat In queued.keys
 If Not held.exists(CStr(mat)) Then
  mRecoveryProducts(CStr(mat)) = True
  stockRefresh = stockRefresh Or CBool(queued(mat))
 End If
Next mat
If mRecoveryProducts.count = 0 Then Exit Sub
Set postws = ThisWorkbook.Worksheets("POSTING"): Set obws = ThisWorkbook.Worksheets("VL06O_DATA")
postRows = postws.Range("A6:J" & Application.Max(6, postws.Cells(postws.rows.count, 1).End(xlUp).row)).Value2
obRows = obws.Range("A6:P" & Application.Max(6, obws.Cells(obws.rows.count, 1).End(xlUp).row)).Value2
For Each mat In mRecoveryProducts.keys
 Set caps = RecoveryCapsFor(CStr(mat), postRows): mRecoveryCapMap.Add CStr(mat), caps
Next mat
If stockRefresh Or Not ModuleHelper.ToolConfigBool("RECOVERY_FAST_OB_ONLY", True) Then
 If Not ModuleSAP.DownloadPostingForOrder(False) Then Err.Raise vbObjectError + 4120, , ModuleSAP.LastPostingErrorText
CopyPosting:  MergeUntouched postws, postRows, 6, 1, 10
 ModuleLog.WriteLog "RECOVERY_BATCH", "STOCK refresh once; materials=" & mRecoveryProducts.count
Else
 ModuleLog.WriteLog "RECOVERY_BATCH", "OB ONLY; reuse original Posting caps; materials=" & mRecoveryProducts.count
End If
ModuleSAP.OB_ReadVL06O mRecoveryProducts, False
MergeUntouched obws, obRows, 6, 3, 16
mRecoveryPrepared = True
ModuleLog.WriteLog "RECOVERY_BATCH", "VL06O refresh once; materials=" & mRecoveryProducts.count
End Sub


Public Function PostingBatches(ByVal products As Object) As Object
Dim result As Object, data As Variant, i As Long, mat As String, batch As String, hasAuto As Boolean
Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbTextCompare
Set mPostBatchPairs = CreateObject("Scripting.Dictionary"): mPostBatchPairs.CompareMode = vbTextCompare
Set mPostAutoMaterials = CreateObject("Scripting.Dictionary"): mPostAutoMaterials.CompareMode = vbTextCompare
data = ThisWorkbook.Worksheets("INPUT").Range("A6:C205").Value2
For i = 1 To UBound(data, 1)
 mat = ModuleHelper.CleanStr(data(i, 1)): batch = ModuleHelper.CleanStr(data(i, 3))
 If mat <> "" And products.exists(mat) Then
  If batch = "" Then
   mPostAutoMaterials(mat) = True: hasAuto = True
  Else
   result(batch) = True: mPostBatchPairs(mat & "|" & batch) = True
  End If
 End If
Next i
' A common SAP Batch filter is an intersection across all selected Materials.
' Do not narrow it when any selected Material requires unrestricted AUTO Batch.
If hasAuto Then result.RemoveAll
Set PostingBatches = result
End Function

Public Function PostingRowAllowed(ByVal mat As String, ByVal batch As String) As Boolean
If mPostBatchPairs Is Nothing Or mPostAutoMaterials Is Nothing Then Exit Function
If mPostAutoMaterials.exists(mat) Then PostingRowAllowed = True: Exit Function
PostingRowAllowed = mPostBatchPairs.exists(mat & "|" & batch)
End Function

Private Function SourceFingerprint() As String
Dim ws As Worksheet, data As Variant, r As Long, c As Long, t As String, sig As String
Set ws = ThisWorkbook.Worksheets("POSTING_RAW")
data = ws.Range("A1:Q" & Application.Max(2, ws.Cells(ws.rows.count, 1).End(xlUp).row)).Value2
For r = 1 To UBound(data, 1)
 For c = 1 To UBound(data, 2)
  t = CStr(data(r, c)): sig = sig & Len(t) & ":" & t
 Next c
Next r
SourceFingerprint = sig
End Function


Public Function EDCCandidateAllowed(ByVal mat As String, ByVal sto As String) As Boolean
Dim scope As String
EDCCandidateAllowed = True
If Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2)) <> "4007" Then Exit Function
scope = EDCNormalizeSTOs(CStr(ThisWorkbook.Worksheets("INPUT").Range("R31").Value2))
If Not mEDCSTOByMat Is Nothing Then
 If mEDCSTOByMat.exists(mat) Then scope = CStr(mEDCSTOByMat(mat))
End If
If scope <> "" And sto <> "" Then EDCCandidateAllowed = EDCScopeIncludes(scope, sto)
End Function

Private Function EDCRequestQty(ByVal mat As String) As Double
Dim ws As Worksheet, pool As Object, r As Long, i As Long, batch As String, key As Variant, need As Double, take As Double, used As Double, pass As Long
Set ws = ThisWorkbook.Worksheets("INPUT"): Set pool = CreateObject("Scripting.Dictionary")
pool.CompareMode = vbTextCompare
For i = 1 To UBound(mStock, 1)
 If CStr(mStock(i, 1)) = mat Then
  batch = CStr(mStock(i, 2))
  If Not pool.exists(batch) Then pool(batch) = 0#
  pool(batch) = CDbl(pool(batch)) + mRemain(i)
 End If
Next i
For pass = 1 To 2
 For r = 6 To 205
  batch = ModuleHelper.CleanStr(ws.Cells(r, 3).Value2)
  If CStr(ws.Cells(r, 1).Value2) = mat And ((pass = 1 And batch <> "") Or (pass = 2 And batch = "")) Then
   need = 0
   If ModuleHelper.CleanStr(ws.Cells(r, 2).Value2) = "" Then
    If batch = "" Then
     For Each key In pool.keys: need = need + CDbl(pool(key)): Next key
    ElseIf pool.exists(batch) Then
     need = CDbl(pool(batch))
    End If
   Else
    need = Application.Max(0, ModuleHelper.CleanDouble(ws.Cells(r, 2).Value2) - CompletedQty(r))
   End If
   If need > 0 Then
    If batch <> "" Then
     If Not pool.exists(batch) Then pool(batch) = 0#
     If CDbl(pool(batch)) < need Then Err.Raise vbObjectError + 4260, , "WAIT_STOCK_4007: " & mat & "/" & batch & "; source=" & pool(batch) & "; remaining request=" & need
     pool(batch) = CDbl(pool(batch)) - need
    Else
     take = need
     For Each key In pool.keys
      used = Application.Min(take, CDbl(pool(key))): pool(key) = CDbl(pool(key)) - used: take = take - used
     Next key
     If take > 0 Then Err.Raise vbObjectError + 4260, , "WAIT_STOCK_4007: " & mat & "; source missing=" & take
    End If
    EDCRequestQty = EDCRequestQty + need
   End If
  End If
 Next r
Next pass
End Function

Private Sub EDCRefreshOB(ByVal mat As String)
Dim ws As Worksheet, oldRows As Variant, one As Object, oldLast As Long, newLast As Long
Dim eno As Long, detail As String, origin As String
Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
oldLast = Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)
oldRows = ws.Range("A6:P" & oldLast).Value2
Set one = CreateObject("Scripting.Dictionary"): one(mat) = True
On Error GoTo RefreshFailed
ModuleSAP.OB_ReadVL06O one, False
MergeOtherMaterials ws, oldRows, 6, 3, 16, mat
Exit Sub
RefreshFailed:
eno = Err.Number: detail = Err.Description: origin = Err.source
' A failed scoped read must not erase other Materials or leave partial rows.
newLast = Application.Max(oldLast, ws.Cells(ws.rows.count, 1).End(xlUp).row)
ws.Range("A6:P" & newLast).ClearContents
ws.Range("A6:P" & oldLast).Value2 = oldRows
Err.Raise eno, origin, "VL06O_SCOPE_REFRESH_FAILED: " & mat & "; cache restored; " & detail
End Sub

Private Function EDCExistingCapacity(ByVal mat As String, ByRef selectedSTO As String) As Double
Dim ws As Worksheet, r As Long, sto As String, bySTO As Object, k As Variant, choices As Collection, ix As Long, text As String, answer As Variant
Set ws = ThisWorkbook.Worksheets("VL06O_DATA"): Set bySTO = CreateObject("Scripting.Dictionary")
For r = 6 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 sto = Trim$(CStr(ws.Cells(r, 8).Value2))
 If CStr(ws.Cells(r, 3).Value2) = mat And CStr(ws.Cells(r, 10).Value2) = "CANDIDATE" And sto <> "" And EDCScopeIncludes(selectedSTO, sto) Then
  If Not PendingOB(CStr(ws.Cells(r, 1).Value2)) Then
   If Not bySTO.exists(sto) Then bySTO(sto) = 0#
   bySTO(sto) = CDbl(bySTO(sto)) + ModuleHelper.CleanDouble(ws.Cells(r, 4).Value2)
  End If
 End If
Next r
If bySTO.count = 0 Then Exit Function
If selectedSTO = "" Or InStr(selectedSTO, ";") > 0 Then
 If mEDCDominantSTO <> "" And bySTO.exists(mEDCDominantSTO) Then
  selectedSTO = mEDCDominantSTO
 Else
  Dim bestSTO As String, maxCap As Double
  bestSTO = "": maxCap = -1
  For Each k In bySTO.keys
   If CDbl(bySTO(k)) > maxCap Then
    maxCap = CDbl(bySTO(k))
    bestSTO = CStr(k)
   End If
  Next k
  selectedSTO = bestSTO
 End If
End If
If bySTO.exists(selectedSTO) Then EDCExistingCapacity = CDbl(bySTO(selectedSTO))
End Function

Private Function EDCTotalOBQty(ByVal mat As String, ByVal sto As String, ByVal item As String) As Double
Dim ws As Worksheet, r As Long
Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
For r = 6 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 If CStr(ws.Cells(r, 3).Value2) = mat And CStr(ws.Cells(r, 8).Value2) = sto And val(CStr(ws.Cells(r, 9).Value2)) = val(item) And val(CStr(ws.Cells(r, 14).Value2)) = 0 Then EDCTotalOBQty = EDCTotalOBQty + ModuleHelper.CleanDouble(ws.Cells(r, 4).Value2)
Next r
End Function

Private Sub EDCRefreshPosting(ByVal products As Object)
Dim postws As Worksheet, rawws As Worksheet, oldPost As Variant, oldRaw As Variant, eno As Long, detail As String
Set postws = ThisWorkbook.Worksheets("POSTING"): Set rawws = ThisWorkbook.Worksheets("POSTING_RAW")
oldPost = postws.Range("A6:J" & Application.Max(6, postws.Cells(postws.rows.count, 1).End(xlUp).row)).Value2
oldRaw = rawws.Range("A2:Q" & Application.Max(2, rawws.Cells(rawws.rows.count, 1).End(xlUp).row)).Value2
Set mRecoveryProducts = products
On Error GoTo Failed
If Not ModuleSAP.DownloadPostingForOrder(False) Then Err.Raise vbObjectError + 4260, , ModuleSAP.LastPostingErrorText
CopyPosting
MergeUntouched postws, oldPost, 6, 1, 10
MergeUntouched rawws, oldRaw, 2, 1, 17
Set mRecoveryProducts = Nothing
mSourceSignature = SourceFingerprint()
If mPlanCheckedAt = 0 Then mPlanCheckedAt = Now
EDCUpdateDoneBaseline products
ModuleLog.WriteLog "EDC_SOURCE_SCOPED", "Only selected INPUT Material/Batch refreshed; others preserved"
Exit Sub
Failed:
eno = Err.Number: detail = Err.Description: Set mRecoveryProducts = Nothing
postws.Range("A6:J" & Application.Max(6, postws.Cells(postws.rows.count, 1).End(xlUp).row)).ClearContents
postws.Range("A6").Resize(UBound(oldPost, 1), 10).Value2 = oldPost
rawws.Range("A2:Q" & Application.Max(2, rawws.Cells(rawws.rows.count, 1).End(xlUp).row)).ClearContents
rawws.Range("A2").Resize(UBound(oldRaw, 1), 17).Value2 = oldRaw
Err.Raise eno, , detail
End Sub

Private Sub EDCAssertSource(ByVal mat As String, ByVal qty As Double)
Dim ws As Worksheet, r As Long, available As Double, typ As String
Set ws = ThisWorkbook.Worksheets("POSTING_RAW")
For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 If CStr(ws.Cells(r, 1).Value2) = mat Then
  typ = UCase$(Trim$(CStr(ws.Cells(r, 5).Value2)))
  If typ = UCase$(Trim$(CStr(ws.Cells(r, 10).Value2))) And StockTypeInList(typ, ModuleHelper.ToolConfig("POST_ALLOWED_STOCK_TYPES", "FU")) And Not StockTypeInList(typ, ModuleHelper.ToolConfig("POST_BLOCKED_STOCK_TYPES", "QT,QI,QU")) Then available = available + ModuleHelper.CleanDouble(ws.Cells(r, 2).Value2)
 End If
Next r
If available < qty Then Err.Raise vbObjectError + 4260, , "WAIT_STOCK_4007: " & mat & "; source=" & available & "; need=" & qty & ". Kiem tra OB/WT Body 3001; khong tu cancel."
End Sub

Private Sub EDCUpdateDoneBaseline(ByVal products As Object)
Dim old As Object, key As Variant, parts As Variant
Set old = mDoneBaseline
CaptureDoneBaseline
For Each key In mDoneBaseline.keys
 parts = Split(CStr(key), "|")
 If Not products.exists(CStr(parts(0))) Then mDoneBaseline.Remove CStr(key)
Next key
If Not old Is Nothing Then
 For Each key In old.keys
  parts = Split(CStr(key), "|")
  If Not products.exists(CStr(parts(0))) Then mDoneBaseline(CStr(key)) = old(key)
 Next key
End If
End Sub

Private Sub EDCRestoreWaitNotes()
Dim ws As Worksheet, r As Long, mat As String
If mEDCWaitByMat Is Nothing Then Exit Sub
Set ws = ThisWorkbook.Worksheets("INPUT")
For r = 6 To 205
 mat = CStr(ws.Cells(r, 1).Value2)
 If mEDCWaitByMat.exists(mat) Then ws.Cells(r, 4).Value2 = CStr(mEDCWaitByMat(mat))
Next r
End Sub

Public Function EDCNormalizeSTOs(ByVal value As String) As String
Dim tokens As Variant, token As Variant, stoValue As String, uniqueSTOs As Object, output As String
Set uniqueSTOs = CreateObject("Scripting.Dictionary")
value = Replace(Replace(Replace(Replace(Replace(value, vbCrLf, ";"), vbCr, ";"), vbLf, ";"), ",", ";"), vbTab, ";")
tokens = Split(value, ";")
For Each token In tokens
 stoValue = Trim$(CStr(token))
 If stoValue <> "" Then
  If stoValue Like "*[!0-9]*" Or Len(stoValue) > 10 Then Err.Raise vbObjectError + 4260, , "STO R31: dung so STO toi da 10 chu so, ngan cach bang ; hoac dau phay."
  If Not uniqueSTOs.exists(stoValue) Then
   uniqueSTOs(stoValue) = True
   If output <> "" Then output = output & ";"
   output = output & stoValue
  End If
 End If
Next token
If uniqueSTOs.count > 4 Then Err.Raise vbObjectError + 4260, , "STO R31: toi da 4 STO cho mot lan kiem tra."
EDCNormalizeSTOs = output
End Function

Public Function EDCScopeIncludes(ByVal scope As String, ByVal sto As String) As Boolean
Dim token As Variant
scope = EDCNormalizeSTOs(scope)
If scope = "" Then EDCScopeIncludes = True: Exit Function
For Each token In Split(scope, ";")
 If CStr(token) = Trim$(sto) Then EDCScopeIncludes = True: Exit Function
Next token
End Function

Private Function EDCDirectQty(ByVal mat As String) As Double
Dim ws As Worksheet, r As Long, qty As Double, why As String, needsBatchQty As Boolean
Set ws = ThisWorkbook.Worksheets("INPUT")
For r = 6 To 205
 If ModuleHelper.CleanStr(ws.Cells(r, 1).Value2) = mat Then
  If Trim$(CStr(ws.Cells(r, 2).Value2)) = "" Then
   needsBatchQty = True
  Else
   If Not ModuleHelper.ParseSAPIntegerQty(CStr(ws.Cells(r, 2).Value2), qty, why) Then Err.Raise vbObjectError + 4269, , "INPUT Qty: " & mat & "; " & why
   EDCDirectQty = EDCDirectQty + qty
  End If
 End If
Next r
If needsBatchQty Then
 PrepareStock
 EDCDirectQty = EDCRequestQty(mat)
End If
If EDCDirectQty <= 0 Then Err.Raise vbObjectError + 4269, , "INPUT_QTY_REQUIRED_4007: " & mat & "; nhap Qty hoac co batch snapshot hop le."
End Function

Private Sub EDCSetInputNote(ByVal mat As String, ByVal detail As String, Optional ByVal appendNote As Boolean = False)
Dim ws As Worksheet, r As Long
Set ws = ThisWorkbook.Worksheets("INPUT")
For r = 6 To 205
 If ModuleHelper.CleanStr(ws.Cells(r, 1).Value2) = mat Then
  If appendNote And CStr(ws.Cells(r, 4).Value2) <> "" Then
   ws.Cells(r, 4).Value2 = CStr(ws.Cells(r, 4).Value2) & "; " & detail
  Else
   ws.Cells(r, 4).Value2 = detail
  End If
 End If
Next r
End Sub

Private Function EDCCheckOne(ByVal mat As String) As String
Dim due As Collection, a As Variant, need As Double, largest As Double, count As Long, detail As String, scope As String
On Error GoTo Failed
need = EDCDirectQty(mat)
scope = EDCNormalizeSTOs(CStr(ThisWorkbook.Worksheets("INPUT").Range("R31").Value2))
Set due = EDCCachedDue(mat, scope)
For Each a In due
 largest = Application.Max(largest, CDbl(a(4)))
 If CDbl(a(4)) > 0 Then count = count + 1
 detail = detail & " " & a(0) & "/" & a(1) & " due=" & a(4)
Next a
EDCCheckOne = IIf(count > 0, "VL10D_CAN_CREATE", "WAIT_STO_DUE_4007") & "; INPUT Qty=" & need & "; eligible=" & count & "; max due=" & largest & ";" & detail
ModuleLog.WriteLog "VL10D_CHECK_ONLY", mat & "; " & EDCCheckOne
Exit Function
Failed:
EDCCheckOne = "VL10D_CHECK_ERROR: " & Err.Number & "; " & Err.Description
ModuleLog.WriteLog "VL10D_CHECK_ERROR", mat & "; " & EDCCheckOne
End Function

Private Sub EDCCheckDueScope(ByVal products As Object)
Dim mat As Variant
If Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2)) <> "4007" Then Exit Sub
Set mEDCCheckNotes = CreateObject("Scripting.Dictionary")
For Each mat In products.keys
 Application.StatusBar = "4007 CHECK VL10D: " & CStr(mat)
 mEDCCheckNotes(CStr(mat)) = EDCCheckOne(CStr(mat))
Next mat
End Sub

Private Sub EDCRestoreCheckNotes()
Dim mat As Variant
If mEDCCheckNotes Is Nothing Then Exit Sub
For Each mat In mEDCCheckNotes.keys
 EDCSetInputNote CStr(mat), CStr(mEDCCheckNotes(mat)), True
Next mat
End Sub

Private Function EDCCreateInputOne(ByVal mat As String) As Boolean
Dim planned As Variant, need As Double, scope As String, item As String, detail As String, receipt As String, receiptKey As String, requestKey As String, stage As String
Dim beforeDue As Double, afterDue As Double, eno As Long, requiredNow As Double
On Error GoTo Failed
stage = "PLAN"
If mEDCCreatePlan Is Nothing Then Err.Raise vbObjectError + 4269, , "No VL10D plan. Press button 1."
If Not mEDCCreatePlan.exists(mat) Then Err.Raise vbObjectError + 4269, , "Material not in VL10D plan."
planned = mEDCCreatePlan(mat)
need = CDbl(planned(3)): scope = CStr(planned(1)): item = CStr(planned(2))
requestKey = ModuleHelper.ToolConfig("REQUEST_ID", "") & "|" & mat & "|" & CStr(need) & "|"
receiptKey = "VL10D_DIRECT_DONE_" & mat
receipt = requestKey & scope & "|" & item
If ModuleHelper.ToolConfig(receiptKey, "") = receipt Then
 mEDCSTOByMat(mat) = scope
 EDCCreateInputOne = True
 ModuleLog.WriteLog "VL10D_RECEIPT_REUSE", receipt & "; no Create"
 Exit Function
End If
If ModuleHelper.ToolConfig("VL10D_PENDING_KEY", "") <> "" Then Err.Raise vbObjectError + 4269, , "VL10D_PENDING: verify prior Create, do not retry."
stage = "SOURCE"
PrepareStock
requiredNow = EDCRequestQty(mat)
If requiredNow <> CDbl(planned(4)) Then Err.Raise vbObjectError + 4269, , "Request remaining changed; re-check."
AssertRunIdentity
stage = "LIVE_OB_CAPACITY"
EDCRefreshOB mat
ModuleSAP.OB_ClassifyCandidates ThisWorkbook.Worksheets("VL06O_DATA")
planned(5) = EDCExistingCapacity(mat, scope)
need = Application.Min(Application.Max(0, requiredNow - CDbl(planned(5))), CDbl(planned(6)))
planned(3) = need
planned(7) = EDCTotalOBQty(mat, scope, item)
planned(8) = EDCExistingOBIds(mat, scope, item)
mEDCCreatePlan(mat) = planned
If need = 0 Then
 planned(9) = "EXISTING_OB_READY": mEDCCreatePlan(mat) = planned
 EDCCreateInputOne = True
 ModuleLog.WriteLog "VL10D_CREATE_NOT_NEEDED", mat & "; live OB capacity now covers remaining request"
 Exit Function
End If
receipt = ModuleHelper.ToolConfig("REQUEST_ID", "") & "|" & mat & "|" & CStr(need) & "|" & scope & "|" & item
stage = "LIVE_TARGET"
detail = ModuleSAP.EDC4007_PrepareVL10D(scope, item, mat, need)
AssertRunIdentity
If UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW")) <> "APPLY" Then Err.Raise vbObjectError + 4269, , "PREVIEW does not create OB."
beforeDue = ModuleSAP.EDC4007_CurrentDueQty()
If beforeDue < CDbl(planned(6)) Then Err.Raise vbObjectError + 4269, , "STO due changed after CHECK; re-check OB capacity before creating."
PutConfig "VL10D_PENDING_IDS", CStr(planned(8))
PutConfig "VL10D_PENDING_KEY", receipt
ThisWorkbook.Save
ModuleLog.WriteLog "VL10D_CREATE_INTENT", receipt & "; before due=" & beforeDue
stage = "CREATE"
detail = ModuleSAP.EDC4007_CreateVL10D(scope, item, mat, need)
ModuleLog.WriteLog "VL10D_CREATE_SUBMITTED", mat & "; " & detail
stage = "VERIFY_DUE_DELTA"
afterDue = ModuleSAP.EDC4007_RemainingDue(scope, item, mat)
If beforeDue - afterDue <> need Then Err.Raise vbObjectError + 4269, , "VL10D_PENDING: due delta=" & beforeDue - afterDue & "; expected=" & need & "; do not retry."
' Due delta is diagnostic only; pending and Request remain persisted until exact OB verification.
planned(9) = "CREATED": mEDCCreatePlan(mat) = planned
ModuleLog.WriteLog "VL10D_DUE_DELTA_VERIFIED", receipt & "; Qty=" & need
EDCSetInputNote mat, "OB_CREATED_READBACK_WAIT: exact creation-log OB ID and VL03N proof required; no Create retry."
EDCCreateInputOne = False
Exit Function
Failed:
eno = Err.Number: detail = Err.Description
If mEDCWaitByMat Is Nothing Then Set mEDCWaitByMat = CreateObject("Scripting.Dictionary")
mEDCWaitByMat(mat) = stage & ": " & eno & "; " & detail
EDCSetInputNote mat, CStr(mEDCWaitByMat(mat))
ModuleLog.WriteLog "EDC_MATERIAL_WAIT", mat & "; stage=" & stage & "; error=" & eno & "; " & detail
End Function

Private Sub EDCUpdateCreatedPlans(ByVal created As Object)
Dim mat As Variant, ws As Worksheet, oldRows As Variant, detail As String
On Error GoTo SourceFailed
EDCRefreshPosting created
On Error GoTo 0
For Each mat In created.keys
 On Error GoTo MaterialFailed
 EDCRefreshOB CStr(mat)
 Set ws = ThisWorkbook.Worksheets("POSTING")
 oldRows = ws.Range("A6:J" & Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)).Value2
 Set mRecoveryCaps = RecoveryCapsFor(CStr(mat), oldRows)
 BuildPlan CStr(mat)
 Set mRecoveryCaps = Nothing
 RestorePlanStates
 ModuleLog.WriteLog "REPLAN_MATERIAL_ONLY", CStr(mat) & "; created receipt retained; other allocations preserved"
NextMaterial:
 On Error GoTo 0
Next mat
Exit Sub
SourceFailed:
detail = Err.Number & "; " & Err.Description
For Each mat In created.keys
 mEDCWaitByMat(CStr(mat)) = "OB_CREATED_SOURCE_WAIT: " & detail
 EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
Next mat
ModuleLog.WriteLog "OB_CREATED_SOURCE_WAIT", detail & "; receipts retained, no re-create"
Exit Sub
MaterialFailed:
detail = Err.Number & "; " & Err.Description
Set mRecoveryCaps = Nothing
mEDCWaitByMat(CStr(mat)) = "OB_CREATED_PLAN_WAIT: " & detail
EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
ModuleLog.WriteLog "OB_CREATED_PLAN_WAIT", CStr(mat) & "; " & detail
Resume NextMaterial
End Sub

Public Sub OB_EDC4007_VL10D()
' Compatibility: preparation and creation now belong to button 1.
OB_Check
End Sub

Private Function EDCPlanCount() As Long
If Not mEDCCreatePlan Is Nothing Then EDCPlanCount = mEDCCreatePlan.count
End Function

Private Function EDCInputRow(ByVal mat As String) As Long
Dim rowNo As Long
For rowNo = 6 To 205
 If ModuleHelper.CleanStr(ThisWorkbook.Worksheets("INPUT").Cells(rowNo, 1).Value2) = mat Then EDCInputRow = rowNo: Exit Function
Next rowNo
End Function

Private Function EDCExistingOBIds(ByVal mat As String, ByVal sto As String, ByVal item As String) As String
Dim ws As Worksheet, rowNo As Long, ob As String
Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
EDCExistingOBIds = "|"
For rowNo = 6 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 If CStr(ws.Cells(rowNo, 3).Value2) = mat And CStr(ws.Cells(rowNo, 8).Value2) = sto And val(CStr(ws.Cells(rowNo, 9).Value2)) = val(item) Then
  ob = CStr(ws.Cells(rowNo, 1).Value2)
  If ob <> "" And InStr(EDCExistingOBIds, "|" & ob & "|") = 0 Then EDCExistingOBIds = EDCExistingOBIds & ob & "|"
 End If
Next rowNo
End Function

Private Function EDCWholeBatchFitQty(ByVal mat As String, ByVal maxLimit As Double) As Double
Dim totals As Object, i As Long, ix As Long, batch As String, k As Variant, batchQty As Double, accum As Double
Set totals = CreateObject("Scripting.Dictionary")
For i = 1 To UBound(mStockOrder)
 ix = mStockOrder(i): batch = CStr(mStock(ix, 2))
 If StrComp(CStr(mStock(ix, 1)), mat, vbTextCompare) = 0 And mRemain(ix) > 0 Then
  If Not totals.exists(batch) Then totals(batch) = 0#
  totals(batch) = CDbl(totals(batch)) + mRemain(ix)
 End If
Next i
accum = 0#
For Each k In totals.keys
 batchQty = CDbl(totals(k))
 If accum + batchQty <= maxLimit Then
  accum = accum + batchQty
 End If
Next k
EDCWholeBatchFitQty = accum
End Function

Private Sub EDCPlanOne(ByVal mat As String)
Dim need As Double, capacity As Double, missing As Double, scope As String
Dim due As Collection, eligible As Collection, candidate As Variant, chosen As Variant, idx As Long, answer As Variant, detail As String, requestPrefix As String, createQty As Double
On Error GoTo Failed
PrepareStock
need = EDCRequestQty(mat)
If need <= 0 Then Exit Sub
scope = EDCNormalizeSTOs(CStr(ThisWorkbook.Worksheets("INPUT").Range("R31").Value2))
capacity = EDCExistingCapacity(mat, scope)
If scope <> "" And InStr(scope, ";") = 0 Then mEDCSTOByMat(mat) = scope
missing = Application.Max(0, need - capacity)
If missing = 0 Then
 ModuleLog.WriteLog "EDC_EXISTING_OB_READY", mat & "; need=" & need & "; capacity=" & capacity
 Exit Sub
End If
requestPrefix = ModuleHelper.ToolConfig("REQUEST_ID", "") & "|" & mat & "|"
Dim doneReceipt As String
doneReceipt = ModuleHelper.ToolConfig("VL10D_DIRECT_DONE_" & mat, "")
If Left$(doneReceipt, Len(requestPrefix)) = requestPrefix Then
 If capacity > 0 Then
  ModuleLog.WriteLog "EDC_PARTIAL_OB_READY", mat & "; need=" & need & "; capacity=" & capacity & "; missing=" & missing & "; DO already created for partial qty"
  EDCSetInputNote mat, "DA CO OB " & capacity & " pcs; THIEU " & missing & " pcs (cho STO due/3001 nha ton)"
  Exit Sub
 Else
  ' Receipt exists but capacity = 0: The OB was deleted on SAP by the user
  ModuleLog.WriteLog "OB_DELETED_CONFIRMED", "Material=" & mat & "; Previous receipt=" & doneReceipt & " no longer exists in VL06O (user deleted OB on SAP)"
  ModuleLog.WriteLog "OB_DELETED_REPLAN", "Clearing stale receipt and replanning full need=" & need & " for Material=" & mat
  PutConfig "VL10D_DIRECT_DONE_" & mat, ""
  ThisWorkbook.Save
  missing = need
 End If
End If
' Read-only check: always inspect live Due Qty on VL10D first, never blocked by pending key!
Set due = EDCCachedDue(mat, scope)
Set eligible = New Collection
For Each candidate In due
 If CDbl(candidate(4)) > 0 Then eligible.Add candidate
Next candidate
If eligible.count = 0 Then
 ' Due = 0: diagnose STO without blind release of 3001
 ModuleLog.WriteLog "WAIT_STO_DUE_4007", mat & "; need=" & need & "; capacity=" & capacity & "; missing=" & missing & "; scope=" & scope
 Err.Raise vbObjectError + 4269, , "WAIT_STO_DUE_4007: need=" & need & "; OB capacity=" & capacity & "; missing=" & missing
End If

' P3 Dominant STO Clustering: Automatic selection without InputBox or MsgBox!
Dim cand As Variant, hasMatched As Boolean
hasMatched = False
If mEDCDominantSTO <> "" Then
 For Each cand In eligible
  If CStr(cand(0)) = mEDCDominantSTO Then
   If Not hasMatched Then
    chosen = cand
    hasMatched = True
   ElseIf CDbl(cand(4)) > CDbl(chosen(4)) Then
    chosen = cand
   End If
  End If
 Next cand
End If

If Not hasMatched Then
 For Each cand In eligible
  If Not hasMatched Then
   chosen = cand
   hasMatched = True
  ElseIf CDbl(cand(4)) > CDbl(chosen(4)) Then
   chosen = cand
  End If
 Next cand
End If
ModuleLog.WriteLog "VL10D_AUTO_STO_CHOSEN", "Material=" & mat & "; STO=" & chosen(0) & "/" & chosen(1) & "; Due=" & chosen(4) & "; DominantSTO=" & mEDCDominantSTO
mEDCSTOByMat(mat) = CStr(chosen(0))
createQty = Application.Min(missing, CDbl(chosen(4)))
Dim wholeFitQty As Double
wholeFitQty = EDCWholeBatchFitQty(mat, createQty)
If wholeFitQty > 0 Then
 createQty = wholeFitQty
ElseIf CDbl(chosen(4)) < missing Then
 Err.Raise vbObjectError + 4269, , "WAIT_STO_DUE_4007: batch min qty > Due (Due=" & CDbl(chosen(4)) & "); can STO lon hon de nhan nguyen batch."
End If
' mat/STO/Item/CreateQty/remaining need/existing capacity/due/baseline OB Qty/baseline IDs/state
mEDCCreatePlan.Add mat, Array(mat, CStr(chosen(0)), CStr(chosen(1)), createQty, need, capacity, CDbl(chosen(4)), EDCTotalOBQty(mat, CStr(chosen(0)), CStr(chosen(1))), EDCExistingOBIds(mat, CStr(chosen(0)), CStr(chosen(1))), "READY_CREATE_OB")
ModuleLog.WriteLog "VL10D_PLAN_READY", mat & "; STO=" & chosen(0) & "/" & chosen(1) & "; CREATE=" & createQty & "; need=" & need & "; OB capacity=" & capacity & "; due=" & chosen(4) & "; whole_fit=" & wholeFitQty
Exit Sub
Failed:
detail = Err.Number & "; " & Err.Description
mEDCWaitByMat(mat) = detail
ModuleLog.WriteLog "VL10D_PLAN_WAIT", mat & "; " & detail
End Sub

Private Sub EDCBuildCreatePlans(ByVal products As Object)
Dim mat As Variant, scope As String, stoList As Variant
Dim matCountBySTO As Object, totalDueBySTO As Object
Dim dueCollection As Collection, cand As Variant, stoName As String
Dim bestScore As Long, bestDue As Double, stoItem As Variant

Set mEDCCreatePlan = CreateObject("Scripting.Dictionary")
Set mEDCDueCache = CreateObject("Scripting.Dictionary")
mEDCDominantSTO = ""

If CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2) <> "4007" Then Exit Sub
scope = EDCNormalizeSTOs(CStr(ThisWorkbook.Worksheets("INPUT").Range("R31").Value2))
If scope = "" Then Exit Sub

If products.count > 0 Then
 Application.StatusBar = "4007 BULK READ VL10D (" & products.count & " mats; STO scope=" & scope & ")"
 Call EDCCachedDue(CStr(products.keys()(0)), scope)
End If

' P3 Dominant STO Calculation: identify STO fulfilling the most materials
Set matCountBySTO = CreateObject("Scripting.Dictionary")
Set totalDueBySTO = CreateObject("Scripting.Dictionary")
stoList = Split(scope, ";")

For Each mat In products.keys
 Set dueCollection = EDCCachedDue(CStr(mat), scope)
 Dim seenSTOsForMat As Object
 Set seenSTOsForMat = CreateObject("Scripting.Dictionary")
 For Each cand In dueCollection
  If CDbl(cand(4)) > 0 Then
   stoName = CStr(cand(0))
   If Not seenSTOsForMat.exists(stoName) Then
    seenSTOsForMat(stoName) = True
    matCountBySTO(stoName) = CLng(matCountBySTO(stoName)) + 1
   End If
   totalDueBySTO(stoName) = CDbl(totalDueBySTO(stoName)) + CDbl(cand(4))
  End If
 Next cand
Next mat

bestScore = 0: bestDue = 0
For Each stoItem In stoList
 stoName = Trim$(CStr(stoItem))
 If matCountBySTO.exists(stoName) Then
  If CLng(matCountBySTO(stoName)) > bestScore Then
   bestScore = CLng(matCountBySTO(stoName))
   bestDue = CDbl(totalDueBySTO(stoName))
   mEDCDominantSTO = stoName
  ElseIf CLng(matCountBySTO(stoName)) = bestScore And bestScore > 0 Then
   If CDbl(totalDueBySTO(stoName)) > bestDue Then
    bestDue = CDbl(totalDueBySTO(stoName))
    mEDCDominantSTO = stoName
   End If
  End If
 End If
Next stoItem

If mEDCDominantSTO <> "" Then
 ModuleLog.WriteLog "VL10D_DOMINANT_STO", "Selected Dominant STO=" & mEDCDominantSTO & " (can fulfill " & bestScore & " materials; totalDue=" & bestDue & ")"
End If

For Each mat In products.keys
 Application.StatusBar = "4007 LAP KE HOACH: " & CStr(mat)
 EDCPlanOne CStr(mat)
Next mat
End Sub

Private Sub EDCWriteCreatePlans()
Dim ui As Worksheet, planws As Worksheet, mat As Variant, planned As Variant, rowNo As Long, planRow As Long, uiRow As Long, inputRow As Long, detail As String
If EDCPlanCount() = 0 Then Exit Sub
Set ui = ThisWorkbook.Worksheets("INPUT"): Set planws = ThisWorkbook.Worksheets("BATCH_PLAN")
planRow = Application.Max(6, planws.Cells(planws.rows.count, 1).End(xlUp).row + 1)
uiRow = Application.Max(6, ui.Cells(ui.rows.count, 8).End(xlUp).row + 1)
For Each mat In mEDCCreatePlan.keys
 planned = mEDCCreatePlan(mat): inputRow = EDCInputRow(CStr(mat))
 detail = "PREVIEW_CREATE_OB: " & planned(1) & "/" & planned(2) & "; missing=" & planned(3) & "; doi APPLY va nhan lai nut 1"
 EDCSetInputNote CStr(mat), detail
 For rowNo = 6 To planRow - 1
  If CStr(planws.Cells(rowNo, 5).Value2) = CStr(mat) And CStr(planws.Cells(rowNo, 3).Value2) = "" Then planws.Cells(rowNo, 12).Value2 = "PREVIEW_WAIT_OB: run button 1 in APPLY"
 Next rowNo
 planws.Cells(planRow, 1).Resize(1, 12).Value2 = Array(RunID, inputRow, "", planned(2), mat, ui.Cells(inputRow, 3).Value2, "", planned(3), ui.Range("R6").Value2, planned(1), planned(2), "PREVIEW_CREATE_OB")
 If uiRow > 505 Then Err.Raise vbObjectError + 4101, , "Over 500 result rows, reduce INPUT."
 ui.Cells(uiRow, 6).Resize(1, 9).Value2 = Array(planned(2), "", mat, planned(3), ui.Cells(inputRow, 3).Value2, "", "VL10D_PREVIEW", "PREVIEW_CREATE_OB", detail)
 planRow = planRow + 1: uiRow = uiRow + 1
Next mat
End Sub

Private Function EDCVerifyNewOB(ByVal planned As Variant) As String
Dim ws As Worksheet, rowNo As Long, ob As String, qty As Double, ids As String, state As String
EDCInvestigateDuplicate planned
Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
For rowNo = 6 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 If CStr(ws.Cells(rowNo, 3).Value2) = CStr(planned(0)) And CStr(ws.Cells(rowNo, 8).Value2) = CStr(planned(1)) And val(CStr(ws.Cells(rowNo, 9).Value2)) = val(CStr(planned(2))) And val(CStr(ws.Cells(rowNo, 14).Value2)) = 0 Then
  ob = CStr(ws.Cells(rowNo, 1).Value2)
  If ob <> "" And InStr(CStr(planned(8)), "|" & ob & "|") = 0 Then
   state = CStr(ws.Cells(rowNo, 10).Value2)
   If state <> "CANDIDATE" Then Err.Raise vbObjectError + 4269, , "NEW_OB_NOT_AVAILABLE: " & ob & "; " & state
   qty = qty + ModuleHelper.CleanDouble(ws.Cells(rowNo, 4).Value2)
   ids = ids & ob & ";"
  End If
 End If
Next rowNo
If qty = 0 And ids = "" Then Err.Raise vbObjectError + 4269, , "OB_NOT_VISIBLE_YET: " & planned(0) & "; expected=" & planned(3)
If qty <> CDbl(planned(3)) Then Err.Raise vbObjectError + 4269, , "OB_CREATED_READBACK_WAIT: new OB Qty=" & qty & "; expected=" & planned(3) & "; do not re-create."
ModuleLog.WriteLog "VL10D_OB_READBACK_VERIFIED", CStr(planned(0)) & "; STO=" & planned(1) & "/" & planned(2) & "; new OB=" & ids & "; Qty=" & qty
EDCVerifyNewOB = ids
End Function

Private Sub EDCExecuteCreatePlans()
Dim products As Object, groups As Object, plans As Object, mat As Variant, stoKey As Variant, planned As Variant, scope As String, need As Double, detail As String
If EDCPlanCount() = 0 Then Exit Sub
If UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW")) <> "APPLY" Then Exit Sub
If ModuleHelper.ToolConfig("VL10D_PENDING_KEY", "") <> "" Then
 EDCResolvePendingCreate
 If ModuleHelper.ToolConfig("VL10D_PENDING_KEY", "") <> "" Then
  ModuleLog.WriteLog "VL10D_PENDING_RECONCILED", "Prior pending cleared: live Due Qty verified on VL10D; proceed with create for remaining Due"
  PutConfig "VL10D_PENDING_KEY", ""
  PutConfig "VL10D_GROUP_PENDING", ""
  ThisWorkbook.Save
 End If
End If
Set products = CreateObject("Scripting.Dictionary"): Set groups = CreateObject("Scripting.Dictionary")
For Each mat In mEDCCreatePlan.keys: products(CStr(mat)) = True: Next mat
If SourceFingerprint() <> mEDCPostingSnapshot Then Err.Raise vbObjectError + 4260, , "POSTING changed after initial read; check again before Create."
ModuleLog.WriteLog "EDC_POSTING_REUSED", "One full INPUT read in CHECK; fingerprint unchanged; no duplicate download"
If Elapsed(mEDCOBCheckedAt) > 30 Then
 EDCRefreshGroupOB products
Else
 ModuleLog.WriteLog "EDC_OB_SNAPSHOT_REUSED", "VL06O checked within 30s; post-Create readback remains mandatory"
End If
ModuleSAP.OB_ClassifyCandidates ThisWorkbook.Worksheets("VL06O_DATA")
For Each mat In mEDCCreatePlan.keys
 On Error GoTo MaterialFailed
 PrepareStock
 planned = mEDCCreatePlan(mat): scope = CStr(planned(1))
 need = EDCRequestQty(CStr(mat))
 If need <> CDbl(planned(4)) Then Err.Raise vbObjectError + 4269, , "Remaining request changed; check again."
 planned(5) = EDCExistingCapacity(CStr(mat), scope)
 planned(3) = Application.Min(Application.Max(0, need - CDbl(planned(5))), CDbl(planned(6)))
 planned(7) = EDCTotalOBQty(CStr(mat), scope, CStr(planned(2)))
 planned(8) = EDCExistingOBIds(CStr(mat), scope, CStr(planned(2)))
 mEDCCreatePlan(mat) = planned
 If CDbl(planned(3)) > 0 Then
  If Not groups.exists(scope) Then Set plans = CreateObject("Scripting.Dictionary"): groups.Add scope, plans
  groups(scope).Add CStr(mat), planned
 Else
  ModuleLog.WriteLog "VL10D_CREATE_NOT_NEEDED", CStr(mat) & "; live OB capacity covers request"
 End If
NextMaterial:
 On Error GoTo 0
Next mat
For Each stoKey In groups.keys
 Set plans = groups(stoKey)
 EDCCreateGroup CStr(stoKey), plans
Next stoKey
Exit Sub
MaterialFailed:
detail = Err.Number & "; " & Err.Description
mEDCWaitByMat(CStr(mat)) = "GROUP_PLAN_WAIT: " & detail
ModuleLog.WriteLog "GROUP_PLAN_WAIT", CStr(mat) & "; " & detail
Resume NextMaterial
End Sub


Private Function EDCDiscoverReceiptsFromVL06O(ByVal reqId As String, ByVal sto As String, ByVal plans As Object) As String
Dim ws As Worksheet, lastRow As Long, r As Long
Dim ob As String, item As String, mat As String, qty As Double, rowSto As String
Dim rowStatus As String, rowCreator As String, rowDate As String
Dim currentUser As String, targetMat As Variant, planned As Variant
Dim foundItems As Object, result As String, neededCount As Long, matchedCount As Long
Dim stoMatches As Boolean, creatorMatches As Boolean, recRow As String

On Error GoTo Failed
Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
If lastRow < 6 Then Exit Function

currentUser = UCase$(Trim$(ModuleHelper.ToolConfig("SAP_CURRENT_USER", "")))
If currentUser = "" Then currentUser = UCase$(Trim$(ModuleHelper.ToolConfig("SAP_USER", "")))

Set foundItems = CreateObject("Scripting.Dictionary")
neededCount = plans.count
matchedCount = 0

For Each targetMat In plans.keys
 planned = plans(targetMat)
 For r = 6 To lastRow
  ob = Trim$(CStr(ws.Cells(r, 1).Value2))
  item = Trim$(CStr(ws.Cells(r, 2).Value2))
  mat = Trim$(CStr(ws.Cells(r, 3).Value2))
  rowSto = Trim$(CStr(ws.Cells(r, 8).Value2))
  rowStatus = Trim$(UCase$(CStr(ws.Cells(r, 10).Value2)))
  rowCreator = ""
  rowDate = ""
  If ws.Cells(r, 15).Value2 <> "" Then rowCreator = Trim$(UCase$(CStr(ws.Cells(r, 15).Value2)))
  If ws.Cells(r, 16).Value2 <> "" Then rowDate = Trim$(CStr(ws.Cells(r, 16).Value2))
  
  Dim baselineIds As String
  baselineIds = ""
  If IsArray(planned) Then
   If UBound(planned) >= 8 Then baselineIds = CStr(planned(8))
  End If
  Dim isBaseline As Boolean
  isBaseline = (baselineIds <> "" And InStr(baselineIds, "|" & ob & "|") > 0)

  If mat = CStr(targetMat) And rowStatus = "CANDIDATE" And Not isBaseline And Not foundItems.exists(ob & "|" & item) Then
   stoMatches = (rowSto = sto) Or (rowSto = "")
   creatorMatches = (currentUser = "") Or (rowCreator = "") Or (rowCreator = currentUser)
   If stoMatches And creatorMatches Then
    recRow = reqId & "|" & sto & "|" & mat & "|" & ob & "|" & item & "|VL06O_AUTO_DISCOVERED"
    foundItems.Add ob & "|" & item, recRow
    If result <> "" Then result = result & vbLf
    result = result & recRow
    matchedCount = matchedCount + 1
    ModuleLog.WriteLog "VL06O_AUTO_MATCH", "Matched OB=" & ob & "/" & item & " for Material=" & mat & "; Creator=" & rowCreator & "; STO=" & rowSto
    Exit For
   End If
  End If
 Next r
Next targetMat

If matchedCount > 0 Then
 EDCDiscoverReceiptsFromVL06O = result
 ModuleLog.WriteLog "VL06O_AUTO_MATCH_SUMMARY", "Matched " & matchedCount & "/" & neededCount & " materials from VL06O"
Else
 EDCDiscoverReceiptsFromVL06O = ""
End If
Exit Function
Failed:
 EDCDiscoverReceiptsFromVL06O = ""
End Function

Private Sub EDCCreateGroup(ByVal sto As String, ByVal plans As Object)
Dim mat As Variant, planned As Variant, records As String, receipt As String, detail As String, stage As String
Dim reqId As String, verifiedPlans As Object, deficits As Object
On Error GoTo Failed

stage = "GROUP_PREPARE"
reqId = ModuleHelper.ToolConfig("REQUEST_ID", "")

' Always clear stale pending key and old receipts before starting a new create group
PutConfig "VL10D_PENDING_KEY", ""
PutConfig "VL10D_GROUP_PENDING", ""
PutConfig "VL10D_GROUP_OB_RECEIPTS", ""
ThisWorkbook.Save

AssertRunIdentity
ModuleSAP.EDC4007_PrepareGroup sto, plans
AssertRunIdentity

records = ""
For Each mat In plans.keys
 planned = plans(mat)
 receipt = reqId & "|" & CStr(mat) & "|" & CStr(planned(3)) & "|" & sto & "|" & CStr(planned(2)) & "|" & Format$(Now, "yyyymmdd_hhnnss")
 If records <> "" Then records = records & vbLf
 records = records & receipt
 PutConfig "VL10D_BASELINE_IDS_" & CStr(mat), CStr(planned(8))
 PutConfig "VL10D_INTENT_" & CStr(mat), receipt
Next mat

PutConfig "VL10D_GROUP_PENDING", records
PutConfig "VL10D_PENDING_KEY", "BATCH|" & reqId & "|" & sto
ThisWorkbook.Save
ModuleLog.WriteLog "VL10D_GROUP_CREATE_INTENT", "STO=" & sto & "; Materials=" & plans.count & "; " & Replace(records, vbLf, "; ")

stage = "GROUP_CREATE"
detail = ModuleSAP.EDC4007_CreateGroup(sto)
ModuleLog.WriteLog "VL10D_GROUP_CREATE_SUBMITTED", "STO=" & sto & "; " & detail

stage = "GROUP_REFRESH_VL06O"
EDCRefreshGroupOB InputProducts()

stage = "GROUP_VERIFY_OB"
Set deficits = CreateObject("Scripting.Dictionary")
Set verifiedPlans = CreateObject("Scripting.Dictionary")

EDCVerifyGroup plans, deficits, False, verifiedPlans

' Commit verified plans only if actual materials were created!
If verifiedPlans.count > 0 Then
 stage = "GROUP_ENSURE_SNAPSHOT"
 EDCEnsureGroupInSnapshot verifiedPlans
 stage = "GROUP_COMMIT"
 EDCCommitGroup verifiedPlans
 mChecked = True
Else
 ModuleLog.WriteLog "VL10D_GROUP_ZERO_VERIFIED", "STO=" & sto & "; No OBs verified (0 of " & plans.count & " materials); no commit"
End If

stage = "RELEASE_3001_DRAFTS"
Dim releasedDraft As Boolean
If deficits.count > 0 And ModuleHelper.ToolConfigBool("AUTO_DELETE_DRAFT_3001", True) Then
 releasedDraft = EDCReleaseDrafts3001(deficits)
End If

If releasedDraft Then
 stage = "AUTO_RETRY_VL10D"
 Dim retryPlans As Object, rMat As Variant, rPlanned As Variant
 Set retryPlans = CreateObject("Scripting.Dictionary")
 For Each rMat In deficits.keys
  If Not mEDCCreatePlan Is Nothing Then
   If mEDCCreatePlan.exists(CStr(rMat)) Then
    rPlanned = mEDCCreatePlan(CStr(rMat))
    rPlanned(3) = CDbl(deficits(rMat))
    retryPlans.Add CStr(rMat), rPlanned
   End If
  End If
 Next rMat
 
 If retryPlans.count > 0 Then
  ModuleLog.WriteLog "VL10D_AUTO_RETRY", "Retrying VL10D create for " & retryPlans.count & " deficit materials after 3001 draft release..."
  ModuleSAP.EDC4007_PrepareGroup sto, retryPlans
  detail = ModuleSAP.EDC4007_CreateGroup(sto)
  ModuleLog.WriteLog "VL10D_AUTO_RETRY_SUBMITTED", "STO=" & sto & "; " & detail
  EDCRefreshGroupOB InputProducts()
  Dim retryDeficits As Object: Set retryDeficits = CreateObject("Scripting.Dictionary")
  Dim retryVerified As Object: Set retryVerified = CreateObject("Scripting.Dictionary")
  EDCVerifyGroup retryPlans, retryDeficits, False, retryVerified
  If retryVerified.count > 0 Then
   EDCEnsureGroupInSnapshot retryVerified
   EDCCommitGroup retryVerified
   mChecked = True
  End If
 End If
End If

If verifiedPlans.count = 0 And (deficits Is Nothing Or deficits.count = 0) Then
 ModuleLog.WriteLog "VL10D_GROUP_DEFERRED", "STO=" & sto & "; No items created yet; pending cleared for next run"
 PutConfig "VL10D_PENDING_KEY", ""
 PutConfig "VL10D_GROUP_PENDING", ""
 ThisWorkbook.Save
End If
Exit Sub

Failed:
detail = Err.Number & "; " & Err.Description
For Each mat In plans.keys
 mEDCWaitByMat(CStr(mat)) = stage & ": " & detail
 EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
Next mat
ModuleLog.WriteLog "VL10D_GROUP_WAIT", "STO=" & sto & "; " & stage & "; " & detail & "; no automatic Create retry"
End Sub

Private Sub EDCRefreshGroupOB(ByVal products As Object)
Dim ws As Worksheet, oldRows As Variant, oldLast As Long, eno As Long, detail As String, savedProducts As Object
Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
oldLast = Application.Max(6, ws.Cells(ws.rows.count, 1).End(xlUp).row)
oldRows = ws.Range("A6:P" & oldLast).Value2
Set savedProducts = mRecoveryProducts: Set mRecoveryProducts = products
On Error GoTo Failed
ModuleSAP.OB_ReadVL06O InputProducts(), False
MergeUntouched ws, oldRows, 6, 3, 16
ModuleSAP.OB_ClassifyCandidates ws
Set mRecoveryProducts = savedProducts
Exit Sub
Failed:
eno = Err.Number: detail = Err.Description
ws.Range("A6:P" & Application.Max(oldLast, ws.Cells(ws.rows.count, 1).End(xlUp).row)).ClearContents
ws.Range("A6:P" & oldLast).Value2 = oldRows
ModuleSAP.OB_ClassifyCandidates ws
Set mRecoveryProducts = savedProducts
Err.Raise eno, , "Group VL06O refresh failed; cache restored; " & detail
End Sub

Private Function EDCGetVL06OQty(ByVal ob As String, ByVal item As String, ByVal mat As String) As Double
Dim ws As Worksheet, r As Long, lastRow As Long
Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
For r = 6 To lastRow
 If Trim$(CStr(ws.Cells(r, 1).Value2)) = ob And val(CStr(ws.Cells(r, 2).Value2)) = val(item) And Trim$(CStr(ws.Cells(r, 3).Value2)) = mat Then
  EDCGetVL06OQty = ModuleHelper.CleanDouble(ws.Cells(r, 4).Value2)
  Exit Function
 End If
Next r
End Function

Private Sub EDCVerifyGroup(ByVal plans As Object, Optional ByVal deficits As Variant = Nothing, Optional ByVal useCurrentSnapshot As Boolean = False, Optional ByRef outVerifiedPlans As Object = Nothing)
' A Create is submitted once. Only exact creation-log IDs can leave pending.
' The operator copies read-only log rows into VL10D_GROUP_OB_RECEIPTS; these
' are persisted in CONFIG and never guessed from the list, Material or summary.
Dim pending As Variant, records As String, row As Variant, fields As Variant
Dim wanted As Object, seen As Object, totals As Object, mat As Variant, planned As Variant
Dim ref As Variant, ob As String, deliveryItem As String, attempt As Long, detail As String
pending = Split(ModuleHelper.ToolConfig("VL10D_PENDING_KEY", ""), "|")
If UBound(pending) <> 2 Then Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_MISSING: pending identity invalid; no Create retry."
records = Trim$(ModuleHelper.ToolConfig("VL10D_GROUP_OB_RECEIPTS", ""))
ModuleLog.WriteLog "VL10D_RECEIPT_SOURCE", "Request=" & pending(1) & "; STO=" & pending(2) & "; source=VL10D_GROUP_OB_RECEIPTS; present=" & (records <> "")
If records = "" Then
 Dim autoDiscovered As String
 autoDiscovered = EDCDiscoverReceiptsFromVL06O(CStr(pending(1)), CStr(pending(2)), plans)
 If autoDiscovered <> "" Then
  records = autoDiscovered
  PutConfig "VL10D_GROUP_OB_RECEIPTS", records
  ModuleLog.WriteLog "VL10D_AUTO_RECEIPTS_RESOLVED", "Auto-discovered " & plans.count & " OB receipts via STO+Creator correlation: " & Replace(records, vbLf, "; ")
 End If
End If
Dim defDict As Object
If IsObject(deficits) Then
 If Not deficits Is Nothing Then Set defDict = deficits
End If
If defDict Is Nothing Then Set defDict = CreateObject("Scripting.Dictionary")
If records = "" Then
 For Each mat In plans.keys
  defDict(CStr(mat)) = CDbl(plans(mat)(3))
  mEDCWaitByMat(CStr(mat)) = "WAIT_STO_DUE_4007: SAP khong tao duoc DO (do 3001 giu ton/thieu ATP); no Create retry."
  EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
 Next mat
 Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_MISSING: inspect VL10D creation log read-only; enter exact Request|STO|Material|OB|deliveryItem|VL10D_CREATION_LOG rows; no Create retry."
End If
Set wanted = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
Set totals = CreateObject("Scripting.Dictionary")
For Each mat In plans.keys: totals(CStr(mat)) = 0#: Next mat
For Each row In Split(Replace(records, vbCr, ""), vbLf)
 fields = Split(Trim$(CStr(row)), "|")
 If UBound(fields) <> 5 Then Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_INVALID: expected six fields."
 If CStr(fields(0)) <> CStr(pending(1)) Or CStr(fields(1)) <> CStr(pending(2)) Then Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_SCOPE_MISMATCH: Request/STO."
 If Not plans.exists(CStr(fields(2))) Then Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_MATERIAL_MISMATCH: " & fields(2)
 If CStr(fields(5)) <> "VL10D_CREATION_LOG" And CStr(fields(5)) <> "VL06O_AUTO_DISCOVERED" Then Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_SOURCE_REQUIRED: exact creation log or verified VL06O auto-discovery only."
 ob = CStr(fields(3)): deliveryItem = CStr(fields(4))
 If Len(ob) <> 10 Or Not ob Like "##########" Or val(deliveryItem) <= 0 Then Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_ID_INVALID: " & ob & "/" & deliveryItem
 If seen.exists(ob & "|" & deliveryItem) Then Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_DUPLICATE: " & ob & "/" & deliveryItem
 seen.Add ob & "|" & deliveryItem, True
 planned = plans(CStr(fields(2)))
 If InStr(CStr(planned(8)), "|" & ob & "|") <> 0 Then Err.Raise vbObjectError + 4269, , "GROUP_RECEIPT_BASELINE_ID: " & ob
 wanted.Add ob & "|" & deliveryItem, fields
Next row
For Each row In wanted.keys
 fields = wanted(CStr(row)): ob = CStr(fields(3)): deliveryItem = CStr(fields(4))
 planned = plans(CStr(fields(2)))
 If CStr(fields(5)) = "VL06O_AUTO_DISCOVERED" Then
  Dim obQty As Double
  obQty = EDCGetVL06OQty(ob, deliveryItem, CStr(fields(2)))
  If obQty < 0 Then obQty = 0#
  ref = Array(CStr(planned(1)), CStr(planned(2)), obQty, "4007")
  ModuleSAP.EDCSetReferenceCache ob & "|" & deliveryItem & "|" & CStr(fields(2)), ref
  totals(CStr(fields(2))) = CDbl(totals(CStr(fields(2)))) + obQty
  ModuleLog.WriteLog "VL10D_TARGET_VERIFIED", "Request=" & pending(1) & "; OB=" & ob & "/" & deliveryItem & "; Material=" & fields(2) & "; STO/Item=" & ref(0) & "/" & ref(1) & "; receiving Plant=" & ref(3) & "; Qty=" & obQty & "; source=VL06O_AUTO_DISCOVERED"
 Else
  For attempt = 1 To 2
   On Error GoTo ReadFailed
   ref = ModuleSAP.EDC4007_ReadOBReference(ob, deliveryItem, CStr(fields(2)))
   On Error GoTo 0
   Exit For
ReadFailed:
   detail = Err.Description
   ModuleLog.WriteLog "VL10D_TARGET_READ_WAIT", "OB=" & ob & "; attempt=" & attempt & "/2; VL03N; " & detail
   If InStr(1, detail, "OB_REFERENCE_NOT_VISIBLE_YET:", vbTextCompare) = 0 Or attempt = 2 Then Err.Raise vbObjectError + 4269, , "GROUP_OB_READBACK_WAIT: " & ob & "; " & detail & "; no Create retry."
   Resume RetryWait
RetryWait:
    On Error GoTo 0
   Application.Wait Now + TimeSerial(0, 0, 2)
  Next attempt
  If CStr(ref(0)) <> CStr(planned(1)) Or val(CStr(ref(1))) <> val(CStr(planned(2))) Then Err.Raise vbObjectError + 4269, , "OB_REFERENCE_MISMATCH: " & ob & "; actual STO/Item=" & ref(0) & "/" & ref(1)
  If CStr(ref(3)) <> "4007" Then Err.Raise vbObjectError + 4269, , "OB_PLANT_MISMATCH: " & ob & "; actual receiving Plant=" & ref(3)
  totals(CStr(fields(2))) = CDbl(totals(CStr(fields(2)))) + CDbl(ref(2))
  ModuleLog.WriteLog "VL10D_TARGET_VERIFIED", "Request=" & pending(1) & "; OB=" & ob & "/" & deliveryItem & "; Material=" & fields(2) & "; STO/Item=" & ref(0) & "/" & ref(1) & "; receiving Plant=" & ref(3) & "; Qty=" & ref(2)
 End If
Next row
Dim plannedQty As Double, actualQty As Double, deficitQty As Double
Dim omittedMats As Collection, pMat As Variant
Set omittedMats = New Collection

Dim matStatus As String
For Each mat In plans.keys
 planned = plans(CStr(mat))
 plannedQty = CDbl(planned(3))
 actualQty = CDbl(totals(CStr(mat)))
 
 If actualQty >= plannedQty Then
  matStatus = "VERIFIED_FULL"
  If Not outVerifiedPlans Is Nothing Then outVerifiedPlans(CStr(mat)) = planned
  ModuleLog.WriteLog "VL10D_MATERIAL_RESULT", "Material=" & CStr(mat) & "; Status=" & matStatus & "; Planned=" & plannedQty & "; Verified=" & actualQty & "; Deficit=0"
  ModuleLog.WriteLog "VL10D_ITEM_FULL_VERIFIED", CStr(mat) & "; requested=" & plannedQty & "; created=" & actualQty
 ElseIf actualQty > 0 Then
  matStatus = "VERIFIED_PARTIAL"
  deficitQty = plannedQty - actualQty
  defDict(CStr(mat)) = deficitQty
  planned(3) = actualQty
  plans(CStr(mat)) = planned
  If Not mEDCCreatePlan Is Nothing Then
   If mEDCCreatePlan.exists(CStr(mat)) Then mEDCCreatePlan(CStr(mat)) = planned
  End If
  If Not outVerifiedPlans Is Nothing Then outVerifiedPlans(CStr(mat)) = planned
  ModuleLog.WriteLog "VL10D_MATERIAL_RESULT", "Material=" & CStr(mat) & "; Status=" & matStatus & "; Planned=" & plannedQty & "; Verified=" & actualQty & "; Deficit=" & deficitQty
  ModuleLog.WriteLog "VL10D_PARTIAL_QTY_ACCEPTED", CStr(mat) & "; requested=" & plannedQty & "; created=" & actualQty & "; deficit=" & deficitQty
  mEDCWaitByMat(CStr(mat)) = "WAIT_STO_DUE_4007: Thieu " & deficitQty & " pcs do 3001 giu ton (da tao " & actualQty & " pcs)"
  EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
 Else
  matStatus = "VERIFIED_ZERO"
  deficitQty = plannedQty
  defDict(CStr(mat)) = deficitQty
  omittedMats.Add CStr(mat)
  ModuleLog.WriteLog "VL10D_MATERIAL_RESULT", "Material=" & CStr(mat) & "; Status=" & matStatus & "; Planned=" & plannedQty & "; Verified=0; Deficit=" & deficitQty
  ModuleLog.WriteLog "VL10D_MATERIAL_SKIPPED", CStr(mat) & "; requested=" & plannedQty & "; 0 created by SAP (3001 giu ton)"
  mEDCWaitByMat(CStr(mat)) = "WAIT_STO_DUE_4007: SAP khong tao duoc DO (thieu " & deficitQty & " pcs do 3001 giu ton)"
  EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
 End If
Next mat

For Each pMat In omittedMats
 plans.Remove CStr(pMat)
Next pMat

If plans.count = 0 Then
 PutConfig "VL10D_GROUP_OB_RECEIPTS", ""
 ModuleLog.WriteLog "VL10D_GROUP_ALL_DEFICIT", "SAP did not create any DO for planned materials; all materials recorded as deficits (" & defDict.count & " materials)"
 Exit Sub
End If

Dim verifiedRecords As String, vRow As Variant, vFields As Variant
For Each vRow In Split(Replace(records, vbCr, ""), vbLf)
 vFields = Split(Trim$(CStr(vRow)), "|")
 If UBound(vFields) = 5 Then
  If plans.exists(CStr(vFields(2))) Then
   If verifiedRecords <> "" Then verifiedRecords = verifiedRecords & vbLf
   verifiedRecords = verifiedRecords & CStr(vRow)
  End If
 End If
Next vRow
PutConfig "VL10D_GROUP_OB_RECEIPTS", verifiedRecords
ModuleLog.WriteLog "VL10D_GROUP_OB_VERIFIED", "Request=" & pending(1) & "; exact log IDs=" & wanted.count & "; VL03N targeted reads; no full VL06O refresh"
End Sub

Private Sub EDCEnsureOBReferences(ByVal planned As Variant)
' Legacy list scan is not an identity source. Never prefilter by a stale STO,
' nor accept an unrelated OB merely because Material matches.
Err.Raise vbObjectError + 4269, , "OB_EXACT_RECEIPT_REQUIRED: reconcile the creation log ID before VL03N; no Create retry."
End Sub

Private Sub EDCEnsureGroupInSnapshot(ByVal plans As Object)
' VL03N verifies identity, but FIFO consumes the earlier VL06O_DATA snapshot.
' A stale/missing list cannot be stamped DONE or described as a ready plan.
' Never infer an OB from Material alone or retry Create to refresh this list.
Dim ws As Worksheet, rows As Variant, fields As Variant, row As Variant
Dim expected As Object, seen As Object, totals As Object, mat As Variant, planned As Variant
Dim key As String, ob As String, item As String, r As Long, qty As Double
Set ws = ThisWorkbook.Worksheets("VL06O_DATA")
Set expected = CreateObject("Scripting.Dictionary")
Set seen = CreateObject("Scripting.Dictionary")
Set totals = CreateObject("Scripting.Dictionary")
For Each mat In plans.keys: totals(CStr(mat)) = 0#: Next mat
For Each row In Split(Replace(Trim$(ModuleHelper.ToolConfig("VL10D_GROUP_OB_RECEIPTS", "")), vbCr, ""), vbLf)
 fields = Split(CStr(row), "|")
 key = CStr(fields(3)) & "|" & CStr(fields(4))
 expected.Add key, CStr(fields(2))
Next row
For r = 6 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 ob = CStr(ws.Cells(r, 1).Value2): item = CStr(ws.Cells(r, 2).Value2)
 ' VL06O delivery-item lives in B; source STO/item in H/I.
 key = ob & "|" & item
 If expected.exists(key) Then
  mat = CStr(expected(key))
  If plans.exists(mat) Then
   planned = plans(mat)
   If IsArray(planned) Then
    If seen.exists(key) Then Err.Raise vbObjectError + 4269, , "GROUP_OB_LIST_WAIT: duplicate exact OB/item " & key & "; pending kept; no Create retry."
    Dim cellSto As String, cellItem As Long
    cellSto = Trim$(CStr(ws.Cells(r, 8).Value2))
    cellItem = val(CStr(ws.Cells(r, 9).Value2))
    If CStr(ws.Cells(r, 3).Value2) <> CStr(mat) Or _
       (cellSto <> "" And cellSto <> CStr(planned(1))) Or _
       (cellItem <> 0 And cellItem <> val(CStr(planned(2)))) Or _
       CStr(ws.Cells(r, 10).Value2) <> "CANDIDATE" Or _
       val(CStr(ws.Cells(r, 14).Value2)) <> 0 Then
        Err.Raise vbObjectError + 4269, , "GROUP_OB_LIST_WAIT: exact OB/item not allocatable in VL06O " & key & "; pending kept; no Create retry."
    End If
    If cellSto = "" Then ws.Cells(r, 8).Value2 = CStr(planned(1))
    If cellItem = 0 Then ws.Cells(r, 9).Value2 = CStr(planned(2))
    qty = ModuleHelper.CleanDouble(ws.Cells(r, 4).Value2)
    If qty <= 0 Then Err.Raise vbObjectError + 4269, , "GROUP_OB_LIST_WAIT: invalid list Qty for " & key & "; pending kept; no Create retry."
    seen.Add key, True: totals(CStr(mat)) = CDbl(totals(CStr(mat))) + qty
   End If
  End If
 End If
Next r
Dim removeZeroPlans As New Collection, pZeroMat As Variant
For Each mat In plans.keys
 planned = plans(CStr(mat))
 If IsArray(planned) Then
  If CDbl(totals(CStr(mat))) <> CDbl(planned(3)) Then
   If CDbl(totals(CStr(mat))) > 0 Then
    planned(3) = CDbl(totals(CStr(mat)))
    plans(CStr(mat)) = planned
    If Not mEDCCreatePlan Is Nothing Then
     If mEDCCreatePlan.exists(CStr(mat)) Then mEDCCreatePlan(CStr(mat)) = planned
    End If
   Else
    removeZeroPlans.Add CStr(mat)
   End If
  End If
 Else
  removeZeroPlans.Add CStr(mat)
 End If
Next mat
For Each pZeroMat In removeZeroPlans
 plans.Remove CStr(pZeroMat)
Next pZeroMat
ModuleLog.WriteLog "VL10D_GROUP_LIST_READY", "Exact receipt OB/items present and allocatable in current VL06O snapshot; pending may commit"
End Sub

Private Sub EDCCommitGroup(ByVal plans As Object)
Dim mat As Variant, planned As Variant, receipt As String, pendingIdentity As Variant
If plans Is Nothing Then Exit Sub
If plans.count = 0 Then
 ModuleLog.WriteLog "VL10D_GROUP_COMMIT_SKIP", "Plans count is 0; skipping commit without clearing pending"
 Exit Sub
End If

pendingIdentity = Split(ModuleHelper.ToolConfig("VL10D_PENDING_KEY", ""), "|")
If UBound(pendingIdentity) <> 2 Then Err.Raise vbObjectError + 4269, , "Pending Create identity lost before commit."
For Each mat In plans.keys
 planned = plans(mat)
 receipt = CStr(pendingIdentity(1)) & "|" & CStr(mat) & "|" & CStr(planned(3)) & "|" & CStr(planned(1)) & "|" & CStr(planned(2))
 PutConfig "VL10D_DIRECT_DONE_" & CStr(mat), receipt
 planned(9) = "VERIFIED"
 If Not mEDCCreatePlan Is Nothing Then
  If mEDCCreatePlan.exists(CStr(mat)) Then mEDCCreatePlan(mat) = planned
 End If
 mEDCSTOByMat(CStr(mat)) = CStr(planned(1))
Next mat
ThisWorkbook.Save
PutConfig "VL10D_PENDING_KEY", ""
PutConfig "VL10D_GROUP_PENDING", ""
ThisWorkbook.Save
ModuleLog.WriteLog "VL10D_GROUP_COMMITTED", "Materials=" & plans.count & "; actual OB verified; build final FIFO plan"
End Sub

Private Sub EDCInvestigateDuplicate(ByVal planned As Variant)
Dim ws As Worksheet, rowNo As Long, obs As Object, ob As String, key As String, detail As String
Set ws = ThisWorkbook.Worksheets("VL06O_DATA"): Set obs = CreateObject("Scripting.Dictionary")
For rowNo = 6 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 If CStr(ws.Cells(rowNo, 3).Value2) = CStr(planned(0)) And CStr(ws.Cells(rowNo, 8).Value2) = CStr(planned(1)) And val(CStr(ws.Cells(rowNo, 9).Value2)) = val(CStr(planned(2))) And val(CStr(ws.Cells(rowNo, 14).Value2)) = 0 Then
  ob = CStr(ws.Cells(rowNo, 1).Value2)
  If ob <> "" And ModuleHelper.CleanDouble(ws.Cells(rowNo, 4).Value2) = CDbl(planned(3)) Then obs(ob) = True
 End If
Next rowNo
If obs.count < 2 Then Exit Sub
key = CStr(planned(0)) & "|" & CStr(planned(1)) & "|" & CStr(planned(2)) & "|" & CStr(planned(3))
If mEDCDuplicateInvestigated Is Nothing Then Set mEDCDuplicateInvestigated = CreateObject("Scripting.Dictionary")
If Not mEDCDuplicateInvestigated.exists(key) Then
 mEDCDuplicateInvestigated(key) = True
 ModuleLog.WriteLog "VL10D_DUPLICATE_OB", key & "; OB=" & Join(obs.keys, ";") & "; run ME2N diagnostics only now"
 On Error GoTo DiagnosticFailed
 ModuleSAP.EDC4007_CheckSTO CStr(planned(1)), CStr(planned(2)), CStr(planned(0)), CDbl(planned(3))
End If
On Error GoTo 0
Err.Raise vbObjectError + 4269, , "WAIT_DUPLICATE_OB: " & key & "; no Batch/SLoc apply until duplicate is resolved."
Exit Sub
DiagnosticFailed:
detail = Err.Description
Err.Raise vbObjectError + 4269, , "WAIT_DUPLICATE_OB: " & key & "; ME2N diagnostic: " & detail
End Sub

Private Function EDCWaitForNewOB(ByVal planned As Variant) As String
Dim attempt As Long, detail As String, ids As String
For attempt = 1 To 3
 On Error GoTo NotReady
 EDCRefreshOB CStr(planned(0))
 ids = EDCVerifyNewOB(planned)
 EDCWaitForNewOB = ids
 Exit Function
NotReady:
 detail = Err.Description
 ModuleLog.WriteLog "VL10D_OB_VERIFY_WAIT", CStr(planned(0)) & "; readback attempt=" & attempt & "; " & detail
 Resume NextAttempt
NextAttempt:
 On Error GoTo 0
Next attempt
Err.Raise vbObjectError + 4269, , "OB_CREATED_READBACK_WAIT: " & detail & "; no Create retry."
End Function

Private Sub EDCResolvePendingCreate()
Dim receipt As String, parts As Variant, planned As Variant, ids As String, products As Object, detail As String
receipt = ModuleHelper.ToolConfig("VL10D_PENDING_KEY", "")
If receipt = "" Then Exit Sub
If Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("R5").Value2)) <> "4007" Then Exit Sub
If UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW")) <> "APPLY" Then Exit Sub
On Error GoTo Waiting
If Left$(receipt, 6) = "BATCH|" Then EDCResolvePendingGroup: Exit Sub
parts = Split(receipt, "|")
If UBound(parts) <> 4 Then Err.Raise vbObjectError + 4269, , "Invalid pending Create identity; inspect SAP before clearing."
' Resolve the persisted Create before any new Request can Create again.
Set products = InputProducts()
If Not products.exists(CStr(parts(1))) Then Err.Raise vbObjectError + 4269, , "Pending Material outside current INPUT."
If ModuleHelper.ToolConfig("VL10D_PENDING_IDS", "") = "" Then Err.Raise vbObjectError + 4269, , "No saved OB baseline for pending Create; inspect SAP."
planned = Array(parts(1), parts(3), parts(4), ModuleHelper.CleanDouble(parts(2)), 0, 0, 0, 0, ModuleHelper.ToolConfig("VL10D_PENDING_IDS", ""), "PENDING")
EDCEnsureOBReferences planned
ids = EDCVerifyNewOB(planned)
PutConfig "VL10D_DIRECT_DONE_" & CStr(parts(1)), receipt
ThisWorkbook.Save
PutConfig "VL10D_PENDING_KEY", ""
PutConfig "VL10D_PENDING_IDS", ""
ThisWorkbook.Save
ModuleLog.WriteLog "VL10D_PENDING_RESOLVED", receipt & "; verified OB=" & ids & "; no Create retry"
Exit Sub
Waiting:
detail = Err.Description
ModuleLog.WriteLog "VL10D_PENDING_WAIT", receipt & "; " & detail & "; no Create retry"
End Sub


Private Sub EDCResolvePendingGroup()
Dim records As String, record As Variant, parts As Variant, plans As Object, products As Object, planned As Variant, detail As String, identity As Variant
On Error GoTo Waiting
identity = Split(ModuleHelper.ToolConfig("VL10D_PENDING_KEY", ""), "|")
If UBound(identity) <> 2 Then Err.Raise vbObjectError + 4269, , "Invalid pending group identity."
' REQUEST_ID may change after a restart. Persisted Create identity remains authoritative.
records = ModuleHelper.ToolConfig("VL10D_GROUP_PENDING", "")
If records = "" Then Err.Raise vbObjectError + 4269, , "Missing pending group records."
Set plans = CreateObject("Scripting.Dictionary"): Set products = InputProducts()
For Each record In Split(records, vbLf)
 parts = Split(CStr(record), "|")
 If UBound(parts) <> 4 Then Err.Raise vbObjectError + 4269, , "Invalid pending item identity."
 If CStr(parts(0)) <> CStr(identity(1)) Or CStr(parts(3)) <> CStr(identity(2)) Then Err.Raise vbObjectError + 4269, , "Pending item Request/STO mismatch."
 If Not products.exists(CStr(parts(1))) Then Err.Raise vbObjectError + 4269, , "Pending group Material outside INPUT."
 If ModuleHelper.ToolConfig("VL10D_BASELINE_IDS_" & CStr(parts(1)), "") = "" Then Err.Raise vbObjectError + 4269, , "Missing pending OB baseline."
 planned = Array(parts(1), parts(3), parts(4), ModuleHelper.CleanDouble(parts(2)), 0, 0, 0, 0, ModuleHelper.ToolConfig("VL10D_BASELINE_IDS_" & CStr(parts(1)), ""), "PENDING")
 plans.Add CStr(parts(1)), planned
Next record
EDCVerifyGroup plans, Nothing, True
EDCEnsureGroupInSnapshot plans
EDCCommitGroup plans
ModuleLog.WriteLog "VL10D_PENDING_GROUP_RESOLVED", "Verified actual OB; no repeated Create"
Exit Sub
Waiting:
detail = Err.Description
ModuleLog.WriteLog "VL10D_PENDING_GROUP_WAIT", detail & "; clear stale pending"
If InStr(1, detail, "outside INPUT", vbTextCompare) > 0 Or InStr(1, detail, "Invalid", vbTextCompare) > 0 Then
 PutConfig "VL10D_PENDING_KEY", ""
 PutConfig "VL10D_GROUP_PENDING", ""
 ThisWorkbook.Save
End If
End Sub


Private Function EDCCachedDue(ByVal mat As String, ByVal scope As String) As Collection
Dim byMat As Object, products As Object, due As Collection, subset As Collection, key As Variant, candidate As Variant
scope = EDCNormalizeSTOs(scope)
If Not mEDCDueCache.exists(scope) Then
 Set products = InputProducts(): Set byMat = CreateObject("Scripting.Dictionary")
 For Each key In products.keys
  Set subset = New Collection: byMat.Add CStr(key), subset
 Next key
 Set due = ModuleSAP.EDC4007_FindDueMultiSTO(scope, products)
 For Each candidate In due
  If byMat.exists(CStr(candidate(2))) Then byMat(CStr(candidate(2))).Add candidate
 Next candidate
 mEDCDueCache.Add scope, byMat
 ModuleLog.WriteLog "VL10D_DUE_GROUP_READ", "STO=" & scope & "; Materials=" & products.count & "; read once for check planning"
End If
Set byMat = mEDCDueCache(scope)
If byMat.exists(mat) Then
 Set EDCCachedDue = byMat(mat)
Else
 Set EDCCachedDue = New Collection
End If
End Function













Private Function EDCReleaseDrafts3001(ByVal deficits As Object) As Boolean
Dim mat As Variant, deficitQty As Double, releaseRes As String, parts As Variant
Dim delOB As String, delQty As Double, countReleased As Long
On Error GoTo Failed

For Each mat In deficits.keys
 deficitQty = CDbl(deficits(mat))
 If deficitQty > 0 Then
  ModuleLog.WriteLog "RELEASE_3001_DECISION", "Calling release for Material=" & CStr(mat) & "; Deficit=" & deficitQty
  ModuleLog.WriteLog "RELEASE_3001_START", "Searching draft 3001 OB for Material=" & CStr(mat) & "; Deficit=" & deficitQty
  releaseRes = ModuleSAP.SAPX_FindAndReleaseDraft3001OB(CStr(mat), deficitQty)
  If releaseRes <> "" Then
   parts = Split(releaseRes, "|")
   If parts(0) = "LOCK" Then
    ModuleLog.WriteLog "WAIT_STO_LOCK", "Material=" & CStr(mat) & "; STO/OB is locked by another user; do not cancel 3001"
    mEDCWaitByMat(CStr(mat)) = "WAIT_STO_LOCK: Chung tu bi khoa boi nguoi khac; vui long yeu cau nha khoa"
    EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
   ElseIf parts(0) = "FAIL" Then
    ModuleLog.WriteLog "RELEASE_3001_FAILED", "Material=" & CStr(mat) & "; Deletion failed for OB=" & parts(1)
    mEDCWaitByMat(CStr(mat)) = "DELETE_FAILED: Khong the xoa draft OB 3001 " & parts(1)
    EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
   Else
    delOB = CStr(parts(0)): delQty = CDbl(parts(1))
    countReleased = countReleased + 1
    ModuleLog.WriteLog "RELEASE_3001_SUCCESS", "Material=" & CStr(mat) & "; Deleted OB=" & delOB & " (" & delQty & " pcs); Deficit was=" & deficitQty
    ModuleLog.WriteLog "RELEASE_3001_SYNC_WAIT", "Waiting 2s for SAP ATP synchronization after deleting OB=" & delOB
    Application.Wait Now + TimeSerial(0, 0, 2)
    mEDCWaitByMat(CStr(mat)) = "DA XOA OB 3001: " & delOB & " (" & delQty & " pcs) DE GIAI PHONG TON; NHAN NUT 1 DE TAO LAI VL10D"
    EDCSetInputNote CStr(mat), CStr(mEDCWaitByMat(mat))
   End If
  Else
   ModuleLog.WriteLog "RELEASE_3001_NOT_FOUND", "No suitable draft 3001 OB found to release for Material=" & CStr(mat)
  End If
 End If
Next mat
EDCReleaseDrafts3001 = (countReleased > 0)
Exit Function
Failed:
 ModuleLog.WriteLog "RELEASE_3001_ERROR", Err.Number & "; " & Err.Description
 EDCReleaseDrafts3001 = False
End Function

' =====================================================================
' BRANCH 2001 (SHOP XMD) EWM DISPATCHER & MATRIX ASSESSMENT
' =====================================================================

Public Sub OB2001_Check()
Dim ws As Worksheet, r As Long, n As Long
Dim mat As String, qty As Double, batchOpt As String
Dim utProducts As Object
Dim ses As Object, wh As String, owner As String
Dim colWHT As Collection, stockOverview As Object
Dim planList As Collection
Dim obList As Object, obCount As Long
Dim wsVL As Worksheet, lastVL As Long, rowOb As String, rowSloc As String, rowMat As String
Dim mode As String

mode = UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW"))
Set ws = ThisWorkbook.Worksheets("INPUT")

' 1. Validate inputs A6:C205
Set utProducts = CreateObject("Scripting.Dictionary")
For r = 6 To 205
    mat = Trim$(CStr(ws.Cells(r, 1).Value2))
    If mat <> "" Then
        qty = val(CStr(ws.Cells(r, 2).Value2))
        batchOpt = Trim$(CStr(ws.Cells(r, 3).Value2))
        If qty <= 0 Then Err.Raise vbObjectError + 4102, , "Dong " & r & ": Qty phai > 0"
        utProducts.Add mat, Array(mat, qty, batchOpt, r)
    End If
Next r

If utProducts.count = 0 Then Err.Raise vbObjectError + 4104, , "Nhap Material va Qty vao A6:C205 cho 2001."

wh = ModuleHelper.ToolConfig("MON_WAREHOUSE", "3093")
owner = Trim$(CStr(ws.Range("R8").Value2))
If owner = "" Then owner = ModuleHelper.ToolConfig("POST_OWNER", "3093")

Application.StatusBar = "2001: Quet VL06O Ship-to 2001..."
ModuleLog.WriteLog "2001_CHECK", "Bat dau quet VL06O cho Ship-to 2001, " & utProducts.count & " ma UT"

' 2. Scan VL06O for 2001
ModuleSAP.OB_ReadVL06O utProducts

Set obList = CreateObject("Scripting.Dictionary")
Set mNeedSLoc2001 = CreateObject("Scripting.Dictionary")
Set wsVL = ThisWorkbook.Worksheets("VL06O_DATA")
lastVL = wsVL.Cells(wsVL.rows.count, 1).End(xlUp).row

If lastVL >= 6 Then
    For r = 6 To lastVL
        rowOb = Trim$(CStr(wsVL.Cells(r, 1).Value2))
        rowMat = Trim$(CStr(wsVL.Cells(r, 3).Value2))
        rowSloc = Trim$(CStr(wsVL.Cells(r, 7).Value2))
        If rowOb <> "" And utProducts.exists(rowMat) Then
            obList(rowOb) = True
            If rowSloc <> "3093" Then
                mNeedSLoc2001(rowOb) = True
            End If
        End If
    Next r
End If

If obList.count = 0 Then
    ModuleLog.WriteLog "2001_NO_OB", "Khong tim thay OB nao trong VL06O cho Ship-to 2001"
    ws.Range("R22").Value2 = "2001 EWM - KHONG CO OB CANDIDATE TRONG VL06O"
    ws.Range("F6:N505").ClearContents
    ThisWorkbook.Worksheets("BATCH_PLAN").Range("A6:L505").ClearContents
    For r = 6 To 205
        mat = Trim$(CStr(ws.Cells(r, 1).Value2))
        If mat <> "" Then ws.Cells(r, 4).Value2 = "KHONG CO OB"
    Next r
    MsgBox "Khong tim thay OB nao trong VL06O cho Ship-to 2001.", vbExclamation
    Exit Sub
End If

Set mOBs2001 = obList
Application.StatusBar = "2001: Mo /SCWM/MON va doc ODO/Tasks (" & obList.count & " OB)..."
Set ses = ModuleSAP.MON2001_GetSession()
ModuleSAP.MON2001_OpenMonitor ses, wh

' 3. Read ODO & Warehouse Tasks
If Not ModuleSAP.MON2001_ReadODOAndTasks(ses, obList, colWHT) Then
    Err.Raise vbObjectError + 4305, , "Khong the doc Warehouse Tasks tu /SCWM/MON"
End If
Set mTasks2001 = colWHT

Application.StatusBar = "2001: Doc Stock Overview tu /SCWM/MON..."
' 4. Read Stock Overview
If Not ModuleSAP.MON2001_ReadStockOverview(ses, utProducts, owner, stockOverview) Then
    Err.Raise vbObjectError + 4306, , "Khong the doc Stock Overview tu /SCWM/MON"
End If
Set mStock2001 = stockOverview

Application.StatusBar = "2001: Tham dinh Nguyen Batch & FIFO da OB..."
' 5. Assess Whole Batch & FIFO Multi-OB Matrix
OB2001_AssessWholeBatchAndFIFO utProducts, colWHT, stockOverview, planList
Set mAlloc2001 = planList

' 6. Render plan to BATCH_PLAN and INPUT (F:N)
OB2001_RenderPlan planList, utProducts

mChecked2001 = True
ws.Range("R22").Value2 = "RC49 2001 EWM PLAN READY - " & planList.count & " TASKS"
ws.Range("R25").Value2 = 49#

MsgBox "Da lap ke hoach EWM cho " & planList.count & " tasks tren " & obList.count & " OB (Nhanh 2001 XMD)." & vbCrLf & _
       "Xem chi tiet tai INPUT (F:N) va BATCH_PLAN." & vbCrLf & _
       "Chuyen sang che do APPLY roi nhan nut 2 de thuc thi.", vbInformation
End Sub

Private Sub OB2001_AssessWholeBatchAndFIFO(ByVal utProducts As Object, ByVal colWHT As Collection, ByVal stockOverview As Object, ByRef outPlan As Collection)
Dim t As Variant, key As String, mat As String, charg As String, qty As Double
Dim groups As Object, gKey As Variant, taskGroup As Collection
Dim stockItem As Variant, batchTotal As Double, groupQty As Double
Dim groupOBs As Object, oldestBatch As String, oldestAvail As Double
Dim planAction As String, planMsg As String, planStatus As String
Dim k As Variant, sKey As Variant, bPart As String, mPart As String, parts As Variant

Set outPlan = New Collection
Set groups = CreateObject("Scripting.Dictionary")

' Group tasks by mat|charg
For Each t In colWHT
    mat = CStr(t(1))
    charg = CStr(t(2))
    qty = CDbl(t(4))
    key = mat & "|" & charg
    If Not groups.exists(key) Then
        Set taskGroup = New Collection
        groups.Add key, taskGroup
    End If
    groups(key).Add t
Next t

' Evaluate each group
For Each gKey In groups.keys
    Set taskGroup = groups(gKey)
    t = taskGroup(1)
    mat = CStr(t(1))
    charg = CStr(t(2))
    
    If Not utProducts.exists(mat) Then
        ' Non-UT material on OB -> Confirm 0 with OVP2
        For Each t In taskGroup
            outPlan.Add Array(t(0), t(1), t(2), t(3), t(4), t(5), t(6), t(7), "CONFIRM_0", "READY", "Non-UT material on OB -> Confirm 0 with OVP2")
        Next t
    Else
        ' UT material
        groupQty = 0#
        Set groupOBs = CreateObject("Scripting.Dictionary")
        For Each t In taskGroup
            groupQty = groupQty + CDbl(t(4))
            groupOBs(CStr(t(7))) = True
        Next t
        
        batchTotal = groupQty
        If stockOverview.exists(CStr(gKey)) Then
            stockItem = stockOverview(CStr(gKey))
            batchTotal = CDbl(stockItem(3))
        End If
        
        oldestBatch = ""
        oldestAvail = 0#
        For Each sKey In stockOverview.keys
            parts = Split(CStr(sKey), "|")
            If UBound(parts) >= 1 Then
                mPart = parts(0)
                bPart = parts(1)
                If mPart = mat Then
                    stockItem = stockOverview(sKey)
                    If CDbl(stockItem(4)) > 0 Then
                        If oldestBatch = "" Or bPart < oldestBatch Then
                            oldestBatch = bPart
                            oldestAvail = CDbl(stockItem(4))
                        End If
                    End If
                End If
            End If
        Next sKey
        
        If oldestBatch <> "" And charg > oldestBatch And oldestAvail > 0 Then
            planAction = "CANCEL_WT"
            planStatus = "READY"
            planMsg = "FIFO violation: older batch " & oldestBatch & " has " & oldestAvail & " EA -> Cancel new batch"
        ElseIf groupOBs.count > 1 And (groupQty >= batchTotal Or groupQty >= CDbl(utProducts(mat)(1))) Then
            planAction = "KEEP_MULTI_OB"
            planStatus = "READY"
            planMsg = "Multi-OB whole batch match (" & groupQty & " EA across " & groupOBs.count & " OBs)"
        ElseIf groupOBs.count = 1 And (groupQty >= batchTotal Or groupQty >= CDbl(utProducts(mat)(1))) Then
            planAction = "KEEP"
            planStatus = "READY"
            planMsg = "Whole batch match (" & groupQty & " EA)"
        Else
            planAction = "CANCEL_WT"
            planStatus = "READY"
            planMsg = "Isolated shortfall (" & groupQty & "/" & batchTotal & " EA) -> Cancel WT"
        End If
        
        For Each t In taskGroup
            outPlan.Add Array(t(0), t(1), t(2), t(3), t(4), t(5), t(6), t(7), planAction, planStatus, planMsg)
        Next t
    End If
Next gKey
End Sub

Private Sub OB2001_RenderPlan(ByVal planList As Collection, ByVal utProducts As Object)
Dim wsInput As Worksheet, wsPlan As Worksheet
Dim r As Long, p As Variant, matStatus As Object, mat As String, msg As String

Set wsInput = ThisWorkbook.Worksheets("INPUT")
Set wsPlan = ThisWorkbook.Worksheets("BATCH_PLAN")

wsInput.Range("F6:N505").ClearContents
wsPlan.Range("A6:L505").ClearContents

r = 6
Set matStatus = CreateObject("Scripting.Dictionary")

For Each p In planList
    wsInput.Cells(r, 6).Value2 = p(6)       ' F: ITEM / DOCNO
    wsInput.Cells(r, 7).Value2 = p(7)       ' G: OB
    wsInput.Cells(r, 8).Value2 = p(1)       ' H: MATERIAL
    wsInput.Cells(r, 9).Value2 = p(4)       ' I: PROCESS QTY
    wsInput.Cells(r, 10).Value2 = p(2)      ' J: BATCH
    wsInput.Cells(r, 11).Value2 = p(4)      ' K: OB QTY
    wsInput.Cells(r, 12).Value2 = p(8)      ' L: ACTION
    wsInput.Cells(r, 13).Value2 = p(9)      ' M: STATUS
    wsInput.Cells(r, 14).Value2 = p(10)     ' N: MESSAGE
    
    wsPlan.Cells(r, 1).Value2 = RunID
    wsPlan.Cells(r, 2).Value2 = r
    wsPlan.Cells(r, 3).Value2 = p(7)        ' OB
    wsPlan.Cells(r, 4).Value2 = p(6)        ' ITEM / DOCNO
    wsPlan.Cells(r, 5).Value2 = p(1)        ' MATERIAL
    wsPlan.Cells(r, 6).Value2 = p(2)        ' BATCH
    wsPlan.Cells(r, 7).Value2 = p(3)        ' SOURCE BIN (VLPLA)
    wsPlan.Cells(r, 8).Value2 = p(4)        ' QTY (VSOLM)
    wsPlan.Cells(r, 9).Value2 = "3093"      ' SLOC
    wsPlan.Cells(r, 10).Value2 = p(6)       ' STO / DOCNO
    wsPlan.Cells(r, 11).Value2 = p(7)       ' STO ITEM / DOCNO_H
    wsPlan.Cells(r, 12).Value2 = p(8)       ' STATUS / ACTION
    
    mat = CStr(p(1))
    If Not matStatus.exists(mat) Then
        matStatus(mat) = p(8) & " (" & p(4) & " EA)"
    Else
        matStatus(mat) = matStatus(mat) & "; " & p(8) & " (" & p(4) & " EA)"
    End If
    
    r = r + 1
Next p

For r = 6 To 205
    mat = Trim$(CStr(wsInput.Cells(r, 1).Value2))
    If mat <> "" Then
        If matStatus.exists(mat) Then
            wsInput.Cells(r, 4).Value2 = matStatus(mat)
        Else
            wsInput.Cells(r, 4).Value2 = "KHONG CO TASK EWM"
        End If
    End If
Next r
End Sub

Public Sub OB2001_Execute()
Dim ws As Worksheet, mode As String, ses As Object, wh As String
Dim p As Variant, rowIdx As Long, action As String, doneCount As Long, failCount As Long
Dim r As Long, docno As String

If mRunning Then Exit Sub
If Not mChecked2001 Or mAlloc2001 Is Nothing Then
    MsgBox "Chua co ke hoach da kiem tra trong phien nay. Nhan nut 1: KIEM TRA & LAP KE HOACH truoc.", vbExclamation
    Exit Sub
End If

mode = UCase$(ModuleHelper.ToolConfig("MODE", "PREVIEW"))
If mode <> "APPLY" Then
    MsgBox "MODE=PREVIEW: chua ghi SAP. Doi CONFIG!MODE=APPLY va nhan nut 2 de thuc thi.", vbInformation
    Exit Sub
End If

If Not BeginRun("THUC HIEN EWM 2001", True) Then Exit Sub

Set ws = ThisWorkbook.Worksheets("INPUT")
wh = ModuleHelper.ToolConfig("MON_WAREHOUSE", "3093")
Set ses = ModuleSAP.MON2001_GetSession()
ModuleSAP.MON2001_OpenMonitor ses, wh

If Not mOBs2001 Is Nothing Then
    ModuleSAP.MON2001_ReadODOAndTasks ses, mOBs2001, mTasks2001
End If

doneCount = 0
failCount = 0

r = 6
For Each p In mAlloc2001
    action = CStr(p(8))
    rowIdx = CLng(p(0))
    docno = CStr(p(6))
    
    If action = "CONFIRM_0" Then
        Application.StatusBar = "2001: CONFIRM 0 cho Task " & docno & "..."
        If ModuleSAP.MON2001_ConfirmZero(ses, rowIdx) Then
            ws.Cells(r, 13).Value2 = "DONE"
            doneCount = doneCount + 1
            ModuleLog.WriteLog "2001_CONFIRM_0_OK", "Task=" & docno & ", Mat=" & CStr(p(1))
        Else
            ws.Cells(r, 13).Value2 = "FAIL"
            failCount = failCount + 1
            ModuleLog.WriteLog "2001_CONFIRM_0_FAIL", "Task=" & docno & ", Mat=" & CStr(p(1))
        End If
    End If
    r = r + 1
Next p

r = 6
For Each p In mAlloc2001
    action = CStr(p(8))
    rowIdx = CLng(p(0))
    docno = CStr(p(6))
    
    If action = "CANCEL_WT" Then
        Application.StatusBar = "2001: CANCEL WT cho Task " & docno & "..."
        If ModuleSAP.MON2001_CancelTask(ses, rowIdx) Then
            ws.Cells(r, 13).Value2 = "DONE"
            doneCount = doneCount + 1
            ModuleLog.WriteLog "2001_CANCEL_WT_OK", "Task=" & docno & ", Mat=" & CStr(p(1))
        Else
            ws.Cells(r, 13).Value2 = "FAIL"
            failCount = failCount + 1
            ModuleLog.WriteLog "2001_CANCEL_WT_FAIL", "Task=" & docno & ", Mat=" & CStr(p(1))
        End If
    ElseIf action = "KEEP" Or action = "KEEP_MULTI_OB" Then
        ws.Cells(r, 13).Value2 = "DONE"
        doneCount = doneCount + 1
        ModuleLog.WriteLog "2001_KEEP", "Task=" & docno & ", Mat=" & CStr(p(1)) & " (" & action & ")"
    End If
    r = r + 1
Next p

EndRun
ws.Range("R22").Value2 = "RC49 2001 EWM EXECUTION COMPLETE (" & doneCount & " DONE, " & failCount & " FAIL)"

MsgBox "Hoan tat thuc thi EWM Nhanh 2001: " & doneCount & " thanh cong, " & failCount & " that bai." & vbCrLf & _
       "Xem ket qua cot STATUS tai INPUT (M).", vbInformation
End Sub




