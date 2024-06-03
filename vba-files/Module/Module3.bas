Attribute VB_Name = "Module3"
Const FOLDER_PATH = "C:\BrityLockNLock\3. Download\P17_셀아웃_데이터_작성_자동화\01_EDI_raw_data"
Const HOMEPLUS = "홈플러스"
Const LOTTE = "롯데마트"
Const EMART = "이마트"
Const COUPANG = "쿠팡"
Const DEST_PATH = "C:\BrityLockNLock\4. Edit\P17_셀아웃_데이터\"
Const DEST_FILE = "주차별데이터_MM월NN주차_내부정리용.xlsx"
Const CONCLUDE = "C:\BrityLockNLock\3. Download\P17_셀아웃_데이터_작성_자동화\자재검증리스트\자재검증리스트.xlsx"

Sub MoveDailyDataToMainExcel(Unrefined As String)

'declare member
Dim rootWb As Workbook
Dim targetWb As Workbook
Dim rootws As Worksheet
Dim targetws As Worksheet
Dim lastrowT As Long
Dim lastrowR As Long
Dim store As String
Dim curDate As String
Dim semiAnnual As String
Dim quarter As String
Dim fileName As String
Dim storename As String
Dim koreandate As String

'assign value
fileName = ExtractFileName(Unrefined)
store = identityStore(fileName)
Set rootWb = Workbooks.Open(DEST_PATH + DEST_FILE)
Set targetWb = Workbooks.Open(Unrefined)
Set rootws = rootWb.Sheets(store)
Set targetws = targetWb.Sheets(1)
 semiAnnual = getSemiAnnum(fileName)
 quarter = getQuarter(fileName)

'execute function



Select Case store
        
    Case HOMEPLUS
        curDate = getdatestring(targetws.Cells(5, 1).value)
        storename = identityStore(fileName)
        lastrow = GetLastRow(targetWb, "Sheet1", "A")
        lastrowSelf = GetLastRow(rootWb, storename, "A") + 1
              
        rootws.Range("G" & lastrowSelf & ":G" & lastrowSelf - 3 + lastrow - 11).value = curDate
        rootws.Range("A" & lastrowSelf & ":A" & lastrowSelf - 3 + lastrow - 11).value = semiAnnual
        rootws.Range("B" & lastrowSelf & ":B" & lastrowSelf - 3 + lastrow - 11).value = quarter
        
        targetws.Range("B13:B" & lastrow - 1).Copy
        rootws.Range("H" & lastrowSelf & ":H" & lastrowSelf - 3 + lastrow - 11).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        targetws.Range("D13:D" & lastrow - 1).Copy
        rootws.Range("O" & lastrowSelf & ":O" & lastrowSelf - 3 + lastrow - 11).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        targetws.Range("E13:E" & lastrow - 1).Copy
        rootws.Range("L" & lastrowSelf & ":L" & lastrowSelf - 3 + lastrow - 11).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False


        targetws.Range("F13:F" & lastrow - 1).Copy
        rootws.Range("M" & lastrowSelf & ":M" & lastrowSelf - 3 + lastrow - 11).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

    Case EMART
        resultString = ExtractFileType(fileName)
        Select Case resultString
            Case "수량"
                lastrow = GetLastRow2(targetWb)
                
                For i = 1 To 7
                
                If i = 1 Then
                targetws.Range("A2:A" & lastrow).Copy
                rootws.Range("H3:H" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                targetws.Range("B2:B" & lastrow).Copy
                rootws.Range("O3:O" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                targetws.Range(targetws.Cells(2, 3), targetws.Cells(lastrow, 3)).Copy
                rootws.Range("L3:L" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                targetws.Cells(1, i + 2).Copy
                rootws.Range("G3:G" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                
                'convert Kr date to iso date
                koreandate = targetws.Cells(1, i + 2).value
                Romandate = ConvertKoreanDateToRoman(koreandate)
                rootws.Range("G3:G" & lastrow + 1).value = Romandate
                
                Else
                targetws.Range("A2:A" & lastrow).Copy
                rootws.Range("H" & (lastrow - 1) * (i - 1) + 3 & ":H" & (lastrow - 1) * i + 3).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                targetws.Range("B2:B" & lastrow).Copy
                rootws.Range("O" & (lastrow - 1) * (i - 1) + 3 & ":O" & (lastrow - 1) * i + 3).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                targetws.Range(targetws.Cells(2, i + 2), targetws.Cells(lastrow, i + 2)).Copy
                rootws.Range("L" & (lastrow - 1) * (i - 1) + 3 & ":L" & (lastrow - 1) * i + 3).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                targetws.Cells(1, i + 2).Copy
                rootws.Range("G" & (lastrow - 1) * (i - 1) + 3 & ":G" & (lastrow - 1) * i + 3).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                
                'convert Kr date to iso date
                koreandate = targetws.Cells(1, i + 2).value
                Romandate = ConvertKoreanDateToRoman(koreandate)
                rootws.Range("G" & (lastrow - 1) * (i - 1) + 3 & ":G" & (lastrow - 1) * i + 3).value = Romandate
                End If
                
                Next i
                
                lastRow2 = GetLastRow(rootWb, store, "G")
                rootws.Range("A3:A" & lastRow2).value = semiAnnual
                rootws.Range("B3:B" & lastRow2).value = quarter
                
             Case "금액"
                lastrow = GetLastRow(rootWb, store, "A")
                rootws.Range("M3:M" & lastrow).value = "=HLOOKUP(RC[-6],'[이마트_0414_금액.xlsx]기간별매출(상품별)_일별요약_금액(원)'!C3:C9,MATCH(RC[-5],'[이마트_0414_금액.xlsx]기간별매출(상품별)_일별요약_금액(원)'!C1,0), FALSE)"
            
            
         '   Case "쿠팡"

                
        End Select

    Case COUPANG
            sortAsc targetWb
            lastrow = targetws.Cells(targetws.Rows.count, "A").End(xlUp).Row
                    
            targetws.Range("A2:A" & lastrow).Copy
            rootws.Range("T3:T" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                   Application.CutCopyMode = False
                    
            targetws.Range("H2:H" & lastrow).Copy
            rootws.Range("O3:O" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                   Application.CutCopyMode = False
                    
            targetws.Range("I2:I" & lastrow).Copy
            rootws.Range("H3:H" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                   Application.CutCopyMode = False
                    
            targetws.Range("L2:L" & lastrow).Copy
            rootws.Range("P3:P" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                   Application.CutCopyMode = False
                    
            targetws.Range("M2:M" & lastrow).Copy
            rootws.Range("Q3:Q" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                   Application.CutCopyMode = False
                    
            targetws.Range("N2:N" & lastrow).Copy
            rootws.Range("R3:R" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                   Application.CutCopyMode = False
                    
            targetws.Range("O2:O" & lastrow).Copy
            rootws.Range("S3:S" & lastrow + 1).PasteSpecial Paste:=xlPasteValues
                   Application.CutCopyMode = False
                    
            rootws.Range("A3:A" & lastrow + 1).value = semiAnnual
            rootws.Range("B3:B" & lastrow + 1).value = quarter


End Select
rootWb.Save
targetWb.Close

End Sub

Sub SealProcess()
'declare member
Dim rootWb As Workbook
Dim targetWb As Workbook
Dim rootws As Worksheet
Dim targetws As Worksheet
Dim lastrow As Long

Dim store As String
Dim curDate As String
Dim semiAnnual As String
Dim quarter As String
Dim fileName As String
Dim storename As String
Dim koreandate As String

'assign value
fileName = ExtractFileName(Unrefined)
store = identityStore(fileName)
Set rootWb = Workbooks.Open(DEST_PATH + DEST_FILE)
Set targetWb = Workbooks.Open(CONCLUDE)
Set targetws = targetWb.Sheets(1)
 semiAnnual = getSemiAnnum(fileName)
 quarter = getQuarter(fileName)

Set rootws = rootWb.Sheets(HOMEPLUS)
rootws.Range("I3:I" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,2,0)"
rootws.Range("J3:J" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,3,0)"
rootws.Range("K3:K" & lastrow).value = "VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,4,0)"
rootws.Range("N3:N" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,5,0)"

Set rootws = rootWb.Sheets(EMART)
rootws.Range("I3:I" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,2,0)"
rootws.Range("J3:J" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,3,0)"
rootws.Range("K3:K" & lastrow).value = "VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,4,0)"
rootws.Range("N3:N" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,5,0)"

Set rootws = rootWb.Sheets(LOTTE)
rootws.Range("I3:I" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,2,0)"
rootws.Range("J3:J" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,3,0)"
rootws.Range("K3:K" & lastrow).value = "VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,4,0)"
rootws.Range("N3:N" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,5,0)"

Set rootws = rootWb.Sheets(COUPANG)
rootws.Range("I3:I" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,2,0)"
rootws.Range("J3:J" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,3,0)"
rootws.Range("K3:K" & lastrow).value = "VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,4,0)"
rootws.Range("N3:N" & lastrow).value = "=VLOOKUP(H3,[자재검증리스트.xlsx]자재코드'!A:E,5,0)"
targetWb.Close
rootws.Range("I3:I" & lastrow).value = "=DATE(LEFT(R[18]C[11],4),MID(R[18]C[11],5,2),RIGHT(R[18]C[11],2))"
rootws.Range("J3:J" & lastrow).value = "=(RC[7])"
rootws.Range("K3:K" & lastrow).value = "=(RC[1]*RC[8])"

rootWb.Save
rootWb.Close

End Sub

Function GetLastRow(wb As Workbook, sheetname As String, col As String) As Long

Dim lastrow As Long
Dim ws As Worksheet
Set ws = wb.Sheets(sheetname)

lastrow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
GetLastRow = lastrow

End Function

Function GetLastRow2(wb As Workbook) As Long

Dim lastrow As Long
Dim ws As Worksheet
Set ws = wb.Sheets(1)

lastrow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
GetLastRow2 = lastrow

End Function

Function identityStore(fileName As String) As String
Dim resultString As String
Position = InStr(fileName, "_")

    If Position > 0 Then
        resultString = Left(fileName, Position - 1)
    End If
    
 Select Case resultString
        Case HOMEPLUS
            resultString = HOMEPLUS
        Case COUPANG
            resultString = COUPANG
        Case EMART
            resultString = EMART
        Case LOTTE
            resultString = LOTTE
        Case Else
            resultString = fileName
    End Select
    identityStore = resultString
End Function

Function getdatestring(D As String) As String
Dim inputString As String
Dim outputstring As String
Dim post As Long

inputString = Trim(D)
post = InStr(inputString, "~")
If post > 0 Then
        resultString = Mid(inputString, post + 1)
        Else
        resultString = D
    End If
getdatestring = resultString
End Function

Function sortAsc(wb As Workbook)
Dim ws As Worksheet
Set ws = wb.Sheets(1)
    lastrow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=Range( _
        "A2:A5333"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ws.Sort
        .SetRange Range("A1:U" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Function

Function getSemiAnnum(fileName As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim resultString As String
    Dim checkstring As String
    Dim finalString As String

    startPos = InStr(fileName, "_")
    endPos = InStr(fileName, ".")
    resultString = Mid(fileName, startPos + 1, endPos - startPos - 1)
    checkstring = Left(resultString, 2)
    
    If (CInt(checkstring) > 7) And Trim(checkstring) <> "" Then
        finalString = "상반기"
    Else
        finalString = "하반기"
    End If
    getSemiAnnum = finalString
End Function

Function getQuarter(fileName As String) As String
Dim startPos As Long
    Dim endPos As Long
    Dim resultString As String
    Dim checkstring As String
    Dim finalString As String

    startPos = InStr(fileName, "_")
    endPos = InStr(fileName, ".")
    resultString = Mid(fileName, startPos + 1, endPos - startPos - 1)
    checkstring = Left(resultString, 2)


        If CInt(checkstring) < 4 And Trim(checkstring) <> "" Then
        finalString = "1Q"
        
        ElseIf CInt(checkstring) < 7 And Trim(checkstring) <> "" Then
        finalString = "2Q"
        
        ElseIf CInt(checkstring) < 10 And Trim(checkstring) <> "" Then
        finalString = "3Q"
        
        Else
        finalString = "4Q"
        
        End If
        
    getQuarter = finalString
End Function

Function ExtractFileName(fullPath As String) As String

    Dim fileName As String
    Dim lastBackslashPos As Long

    lastBackslashPos = InStrRev(fullPath, "\")

    fileName = Mid(fullPath, lastBackslashPos + 1)

    ExtractFileName = fileName
End Function

Function ExtractFileType(fullPath As String) As String

    Dim startPos As Integer
    Dim endPos As Integer
    Dim resultString As String

    startPos = InStr(fullPath, "_") + 1

    startPos = InStr(startPos, fullPath, "_") + 1

    endPos = InStr(startPos, fullPath, ".")

    resultString = Mid(fullPath, startPos, endPos - startPos)

    ExtractFileType = resultString
End Function

Function ConvertKoreanDateToRoman(dateString As String) As String
    Dim koreanMonths As Variant
    Dim monthString As String
    Dim dayString As String
    Dim month As String
    Dim day As String
    Dim currentYear As Integer
    Dim isoDate As String

    koreanMonths = Array("1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월")
    
    monthString = Split(dateString, " ")(0)
    dayString = Split(dateString, " ")(1)
    
    month = Format(Application.Match(monthString, koreanMonths, 0), "00")
    day = Format(Val(Replace(dayString, "일", "")), "00")
    
    currentYear = Year(Date)
    
    isoDate = currentYear & "-" & month & "-" & day
    
    ConvertKoreanDateToRoman = isoDate
End Function


