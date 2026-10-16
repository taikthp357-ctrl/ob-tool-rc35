Attribute VB_Name = "ModuleHelper"

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Public Sub SleepMS(ByVal ms As Long)
    If ms > 0 Then SleepAPI ms
End Sub


Public Function CleanStr(ByVal v As Variant) As String

If IsError(v) Or IsNull(v) Or IsEmpty(v) Then

CleanStr = ""

Else

CleanStr = Trim$(CStr(v))

End If

End Function

Public Function CleanDouble(ByVal v As Variant) As Double

    Dim parsedQty As Double, reasonText As String

    If ParseSAPIntegerQty(v, parsedQty, reasonText) Then
        CleanDouble = parsedQty
    Else
        On Error Resume Next
        CleanDouble = val(Replace(Replace(CStr(v), " ", ""), ",", "."))
        On Error GoTo 0
    End If

End Function

Public Function ParseSAPIntegerQty(ByVal rawValue As Variant, ByRef qty As Double, _
    Optional ByRef reasonText As String) As Boolean

    Dim s As String, integerPart As String, fractionPart As String

    Dim dotPos As Long, commaPos As Long, decimalPos As Long

    Dim decimalSep As String, groupSep As String

    qty = 0

    reasonText = ""

    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then

        reasonText = "rong"

        Exit Function

    End If

    If IsNumeric(rawValue) And VarType(rawValue) <> vbString Then

        If CDbl(rawValue) <= 0 Then

            reasonText = "phai lon hon 0"

        ElseIf CDbl(rawValue) <> Fix(CDbl(rawValue)) Then

            reasonText = "Qty chi nhan so nguyen, khong nhan so le"

        Else

            qty = Fix(CDbl(rawValue))

            ParseSAPIntegerQty = True

        End If

        Exit Function

    End If

    s = Trim$(CStr(rawValue & ""))

    s = Replace(s, " ", "")

    s = Replace(s, ChrW(160), "")

    If s = "" Then

        reasonText = "rong"

        Exit Function

    End If

    If s Like "*[!0-9.,]*" Then

        reasonText = "chi duoc dung chu so va dau . ,"

        Exit Function

    End If

    dotPos = InStrRev(s, ".")

    commaPos = InStrRev(s, ",")

    If dotPos > 0 And commaPos > 0 Then

        ' When both separators exist, only a zero decimal suffix is accepted.

        ' Examples: 1.000,000 and 1,000.000 both mean integer 1000.

        If dotPos > commaPos Then

            decimalPos = dotPos

            decimalSep = "."

            groupSep = ","

        Else

            decimalPos = commaPos

            decimalSep = ","

            groupSep = "."

        End If

        integerPart = Left$(s, decimalPos - 1)

        fractionPart = mid$(s, decimalPos + 1)

        If Not SAPIntegerQty_DigitsOnly(fractionPart) Or _
           Not SAPIntegerQty_AllZeros(fractionPart) Then

            reasonText = "Qty chi nhan so nguyen; phan le phai bang 0"

            Exit Function

        End If

        If InStr(integerPart, groupSep) > 0 Then

            If Not SAPIntegerQty_IsGrouped(integerPart, groupSep) Then

                reasonText = "dau phan cach nghin khong hop le"

                Exit Function

            End If

            s = Replace(integerPart, groupSep, "")

        Else

            If Not SAPIntegerQty_DigitsOnly(integerPart) Then

                reasonText = "dinh dang so luong khong hop le"

                Exit Function

            End If

            s = integerPart

        End If

    ElseIf dotPos > 0 Or commaPos > 0 Then

        If dotPos > 0 Then decimalSep = "." Else decimalSep = ","

        ' A valid three-digit grouping is always thousands, never decimals.

        ' Therefore both 1.000 and 1,000 are parsed as 1000.

        If SAPIntegerQty_IsGrouped(s, decimalSep) Then

            s = Replace(s, decimalSep, "")

        Else

            decimalPos = InStrRev(s, decimalSep)

            integerPart = Left$(s, decimalPos - 1)

            fractionPart = mid$(s, decimalPos + 1)

            If InStr(integerPart, decimalSep) > 0 Or _
               Not SAPIntegerQty_DigitsOnly(integerPart) Or _
               Not SAPIntegerQty_DigitsOnly(fractionPart) Or _
               Not SAPIntegerQty_AllZeros(fractionPart) Then

                reasonText = "Qty chi nhan so nguyen; dau . , phai tach nhom 3 so"

                Exit Function

            End If

            s = integerPart

        End If

    End If

    If Not SAPIntegerQty_DigitsOnly(s) Then

        reasonText = "dinh dang so luong khong hop le"

        Exit Function

    End If

    On Error GoTo InvalidQty

    qty = CDbl(s)

    If qty <= 0 Then

        reasonText = "phai lon hon 0"

        qty = 0

        Exit Function

    End If

    ParseSAPIntegerQty = True

    Exit Function

InvalidQty:

    qty = 0

    reasonText = "vuot qua gioi han so luong"

End Function

Private Function SAPIntegerQty_DigitsOnly(ByVal valueText As String) As Boolean

    SAPIntegerQty_DigitsOnly = (valueText <> "" And Not (valueText Like "*[!0-9]*"))

End Function

Private Function SAPIntegerQty_AllZeros(ByVal valueText As String) As Boolean

    SAPIntegerQty_AllZeros = (valueText <> "" And Replace(valueText, "0", "") = "")

End Function

Private Function SAPIntegerQty_IsGrouped(ByVal valueText As String, ByVal separatorText As String) As Boolean

    Dim parts As Variant, i As Long

    parts = Split(valueText, separatorText)

    If UBound(parts) < 1 Then Exit Function

    If Len(CStr(parts(0))) < 1 Or Len(CStr(parts(0))) > 3 Or _
       Not SAPIntegerQty_DigitsOnly(CStr(parts(0))) Then Exit Function

    For i = 1 To UBound(parts)

        If Len(CStr(parts(i))) <> 3 Or Not SAPIntegerQty_DigitsOnly(CStr(parts(i))) Then Exit Function

    Next i

    SAPIntegerQty_IsGrouped = True

End Function

Public Function ParseSapDate(ByVal v As Variant) As Date

On Error GoTo Invalid

If IsNumeric(v) And VarType(v) <> vbString And Not IsEmpty(v) Then

 If CDbl(v) >= 61 And CDbl(v) < 2958466# Then ParseSapDate = CDate(v)

 Exit Function

End If

If IsDate(v) And Not (VarType(v) = vbString) Then

ParseSapDate = CDate(v)

Exit Function

End If

Dim s As String

s = Trim$(CStr(v & ""))

If s = "" Then GoTo Invalid

If InStr(s, " ") > 0 Then s = Split(s, " ")(0)

' RC20-DEF-06: Explicit SAP_DATE_FORMAT configuration and 8-digit numeric date (YYYYMMDD)
Dim fmt As String
fmt = UCase$(Trim$(ToolConfig("SAP_DATE_FORMAT", "")))
If fmt = "YYYYMMDD" Or (Len(s) = 8 And IsNumeric(s) And InStr(s, ".") = 0 And InStr(s, "/") = 0 And InStr(s, "-") = 0) Then
    Dim y8 As Integer, m8 As Integer, d8 As Integer, dt8 As Date
    y8 = CInt(Left$(s, 4)): m8 = CInt(mid$(s, 5, 2)): d8 = CInt(Right$(s, 2))
    If y8 >= 1900 And y8 <= 2100 And m8 >= 1 And m8 <= 12 And d8 >= 1 And d8 <= 31 Then
        dt8 = DateSerial(y8, m8, d8)
        If Year(dt8) = y8 And Month(dt8) = m8 And Day(dt8) = d8 Then
            ParseSapDate = dt8: Exit Function
        End If
    End If
End If

Dim SEP As String

If InStr(s, ".") > 0 Then

SEP = "."

ElseIf InStr(s, "/") > 0 Then

SEP = "/"

ElseIf InStr(s, "-") > 0 Then

SEP = "-"

Else

If IsDate(s) Then

ParseSapDate = CDate(s)

Exit Function

End If

GoTo Invalid

End If

Dim p() As String

p = Split(s, SEP)

If UBound(p) <> 2 Then GoTo Invalid

Dim d As Integer, m As Integer, y As Integer

If Len(Trim$(p(0))) = 4 Then

y = CInt(p(0)): m = CInt(p(1)): d = CInt(p(2))

Else

d = CInt(p(0)): m = CInt(p(1)): y = CInt(p(2))

End If

' RC20-DEF-06: Support explicit SAP_DATE_FORMAT configuration

Dim fmtSep As String, fmtParts() As String, k As Long, parsedDate As Date

If fmt <> "" Then

    If fmt = "YYYYMMDD" And Len(s) = 8 And IsNumeric(s) Then

        y = CInt(Left$(s, 4)): m = CInt(mid$(s, 5, 2)): d = CInt(Right$(s, 2))

        parsedDate = DateSerial(y, m, d)

        If Year(parsedDate) = y And Month(parsedDate) = m And Day(parsedDate) = d Then

            ParseSapDate = parsedDate: Exit Function

        End If

    Else

        If InStr(fmt, ".") > 0 Then fmtSep = "." Else If InStr(fmt, "/") > 0 Then fmtSep = "/" Else If InStr(fmt, "-") > 0 Then fmtSep = "-"

        If fmtSep <> "" And SEP = fmtSep Then

            fmtParts = Split(fmt, fmtSep)

            If UBound(fmtParts) = 2 And UBound(p) = 2 Then

                For k = 0 To 2

                    If Left$(fmtParts(k), 1) = "Y" Then y = CInt(p(k))

                    If Left$(fmtParts(k), 1) = "M" Then m = CInt(p(k))

                    If Left$(fmtParts(k), 1) = "D" Then d = CInt(p(k))

                Next k

                parsedDate = DateSerial(y, m, d)

                If Year(parsedDate) = y And Month(parsedDate) = m And Day(parsedDate) = d Then

                    ParseSapDate = parsedDate: Exit Function

                End If

            End If

        End If

    End If

End If

parsedDate = DateSerial(y, m, d)

If Year(parsedDate) <> y Or Month(parsedDate) <> m Or Day(parsedDate) <> d Then GoTo Invalid

ParseSapDate = parsedDate

Exit Function

Invalid:

ParseSapDate = 0

End Function

Public Function ToolConfigBool(ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean

Dim raw As String

raw = UCase$(Trim$(ToolConfig(keyName, IIf(defaultValue, "TRUE", "FALSE"))))

Select Case raw

Case "TRUE", "YES", "Y", "1", "ON"

ToolConfigBool = True

Case "FALSE", "NO", "N", "0", "OFF"

ToolConfigBool = False

Case Else

ToolConfigBool = defaultValue

End Select

End Function

Public Function ToolConfigLong(ByVal keyName As String, ByVal defaultValue As Long, _
Optional ByVal minValue As Long = -2147483647, _
Optional ByVal maxValue As Long = 2147483647) As Long

Dim raw As String, resultValue As Long

raw = ToolConfig(keyName, CStr(defaultValue))

If IsNumeric(raw) Then resultValue = CLng(val(raw)) Else resultValue = defaultValue

If resultValue < minValue Then resultValue = minValue

If resultValue > maxValue Then resultValue = maxValue

ToolConfigLong = resultValue

End Function

Public Function ToolConfig(ByVal keyName As String, ByVal defaultValue As String) As String

Dim ws As Worksheet, r As Long

Set ws = ThisWorkbook.Worksheets("CONFIG")

For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row

 If UCase$(Trim$(CStr(ws.Cells(r, 1).Value2))) = UCase$(keyName) Then

  ToolConfig = Trim$(CStr(ws.Cells(r, 2).Value2))

  If ToolConfig = "" Then ToolConfig = defaultValue

  Exit Function

 End If

Next r

ToolConfig = defaultValue

End Function

Public Sub PutConfig(ByVal key As String, ByVal value As String)
Dim ws As Worksheet, r As Long
Set ws = ThisWorkbook.Worksheets("CONFIG")
For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
 If CStr(ws.Cells(r, 1).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub
Next r
ws.Cells(r, 1).Value2 = key: ws.Cells(r, 2).Value2 = value
End Sub

Public Function NormalizeUoM(ByVal uomText As String) As String
    Dim s As String
    s = Trim$(uomText)
    s = Replace(s, Chr(160), "") ' Non-breaking space
    s = Replace(s, Chr(9), "")   ' Tab
    s = Replace(s, Chr(10), "")  ' LF
    s = Replace(s, Chr(13), "")  ' CR
    NormalizeUoM = UCase$(Trim$(s))
End Function

Public Function IsCompatibleUoM(ByVal actualUoM As String, ByVal expectedUom As String) As Boolean
    Dim act As String, exp As String
    act = NormalizeUoM(actualUoM)
    exp = NormalizeUoM(expectedUom)
    
    If act = "" Or exp = "" Then
        IsCompatibleUoM = False
        Exit Function
    End If
    
    If act = exp Then
        IsCompatibleUoM = True
        Exit Function
    End If
    
    ' Equivalence group for base discrete counting units: EA, UN, PC, ST, PCE
    Dim isActCount As Boolean, isExpCount As Boolean
    isActCount = (act = "EA" Or act = "UN" Or act = "PC" Or act = "ST" Or act = "PCE")
    isExpCount = (exp = "EA" Or exp = "UN" Or exp = "PC" Or exp = "ST" Or exp = "PCE")
    
    If isActCount And isExpCount Then
        IsCompatibleUoM = True
        Exit Function
    End If
    
    IsCompatibleUoM = False
End Function

Public Function FormatSAPQty(ByVal qty As Double) As String
    If qty = Fix(qty) Then
        FormatSAPQty = Format$(qty, "0")
    Else
        Dim s As String
        s = Format$(qty, "0.000")
        Do While Right$(s, 1) = "0"
            s = Left$(s, Len(s) - 1)
        Loop
        If Right$(s, 1) = "." Or Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
        FormatSAPQty = s
    End If
End Function



