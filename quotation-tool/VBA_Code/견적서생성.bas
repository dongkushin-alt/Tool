Attribute VB_Name = "견적서생성"
'================================================
' 견적서 생성 및 저장 매크로
' 작성자: Quotation Tool
' 작성일: 2026-06-10
'================================================

Option Explicit

Const QUOTE_SHEET As String = "견적서_작성"
Const QUOTE_TEMP_SHEET As String = "견적서_템플릿"

'================================================
' 1. 견적서 생성 메인 함수
'================================================
Public Sub GenerateQuote()
    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim lastRow As Long
    Dim response As Integer
    Dim fileName As String
    Dim savePath As String
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets(QUOTE_SHEET)
    
    ' 데이터 유무 확인
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "견적 내용이 없습니다. 부품을 먼저 추가해주세요.", vbExclamation, "알림"
        Exit Sub
    End If
    
    ' 견적서 생성 확인
    response = MsgBox("견적서를 생성하시겠습니까?" & vbCrLf & _
                     "부품 수: " & (lastRow - 1) & "개", _
                     vbYesNo, "견적서 생성")
    
    If response = vbNo Then Exit Sub
    
    ' 새 워크북 생성
    Set newWb = Workbooks.Add
    
    ' 데이터 복사
    CopyQuoteData ws, newWb
    
    ' 견적서 정보 입력
    FillQuoteInfo newWb
    
    ' 파일 저장
    fileName = GenerateFileName()
    savePath = CreateQuotationFolder() & "\" & fileName
    
    newWb.SaveAs savePath, xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False
    
    MsgBox "견적서가 생성되었습니다." & vbCrLf & _
           "저장 위치: " & savePath, vbInformation, "완료"
    
    Exit Sub
ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical, "오류"
End Sub

'================================================
' 2. 견적 데이터 복사
'================================================
Private Sub CopyQuoteData(sourceWs As Worksheet, targetWb As Workbook)
    Dim targetWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set targetWs = targetWb.Sheets(1)
    
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 2).End(xlUp).Row
    
    ' 헤더 추가
    targetWs.Cells(1, 1).Value = "순번"
    targetWs.Cells(1, 2).Value = "부품번호"
    targetWs.Cells(1, 3).Value = "부품명"
    targetWs.Cells(1, 4).Value = "원가"
    targetWs.Cells(1, 5).Value = "카테고리"
    targetWs.Cells(1, 6).Value = "수량"
    targetWs.Cells(1, 7).Value = "판매가"
    targetWs.Cells(1, 8).Value = "이익"
    
    ' 헤더 서식
    With targetWs.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(192, 192, 192)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 데이터 복사 (B2부터)
    For i = 2 To lastRow
        targetWs.Cells(i, 1).Value = i - 1 ' 순번
        targetWs.Cells(i, 2).Value = sourceWs.Cells(i, 2).Value ' 부품번호
        targetWs.Cells(i, 3).Value = sourceWs.Cells(i, 3).Value ' 부품명
        targetWs.Cells(i, 4).Value = sourceWs.Cells(i, 4).Value ' 원가
        targetWs.Cells(i, 5).Value = sourceWs.Cells(i, 5).Value ' 카테고리
        targetWs.Cells(i, 6).Value = sourceWs.Cells(i, 6).Value ' 수량
        targetWs.Cells(i, 7).Value = sourceWs.Cells(i, 7).Value ' 판매가
        targetWs.Cells(i, 8).Value = sourceWs.Cells(i, 8).Value ' 이익
    Next i
    
    ' 숫자 서식 적용
    With targetWs.Range("D2:D" & lastRow)
        .NumberFormat = "#,##0"
    End With
    With targetWs.Range("G2:H" & lastRow)
        .NumberFormat = "#,##0"
    End With
    
    ' 합계 계산
    CalculateTotalInQuote targetWs, lastRow
    
    ' 열 너비 자동 조정
    targetWs.Columns("A:H").AutoFit
End Sub

'================================================
' 3. 견적서 정보 입력 (고객정보, 날짜 등)
'================================================
Private Sub FillQuoteInfo(targetWb As Workbook)
    Dim targetWs As Worksheet
    Dim quoteDate As String
    Dim quoteNumber As String
    Dim expiryDate As String
    
    Set targetWs = targetWb.Sheets(1)
    
    ' 견적번호 생성
    quoteNumber = "QT-" & Format(Now(), "yyyymmdd") & "-" & Format(Rnd() * 10000, "0000")
    
    ' 견적일자
    quoteDate = Format(Now(), "yyyy.mm.dd")
    
    ' 유효기간 (30일)
    expiryDate = Format(Now() + 30, "yyyy.mm.dd")
    
    ' 제목 행 추가
    targetWs.Rows(1).Insert
    targetWs.Cells(1, 1).Value = "== 견 적 서 =="
    targetWs.Cells(1, 1).Font.Bold = True
    targetWs.Cells(1, 1).Font.Size = 14
    
    targetWs.Rows(2).Insert
    targetWs.Cells(2, 1).Value = "견적번호: " & quoteNumber
    targetWs.Cells(2, 1).Font.Bold = True
    
    targetWs.Cells(3, 1).Value = "견적일자: " & quoteDate
    targetWs.Cells(4, 1).Value = "유효기간: " & expiryDate
    targetWs.Rows(5).Insert
End Sub

'================================================
' 4. 견적서 합계 계산
'================================================
Private Sub CalculateTotalInQuote(ws As Worksheet, lastDataRow As Long)
    Dim totalPrice As Double
    Dim totalCost As Double
    Dim totalProfit As Double
    Dim i As Long
    Dim summaryRow As Long
    
    summaryRow = lastDataRow + 2
    totalPrice = 0
    totalCost = 0
    totalProfit = 0
    
    ' 합계 계산
    For i = 2 To lastDataRow
        If ws.Cells(i, 7).Value <> "" Then
            totalPrice = totalPrice + ws.Cells(i, 7).Value
            totalCost = totalCost + (ws.Cells(i, 4).Value * ws.Cells(i, 6).Value)
            totalProfit = totalProfit + ws.Cells(i, 8).Value
        End If
    Next i
    
    ' 합계 행에 입력
    ws.Cells(summaryRow, 5).Value = "합계"
    ws.Cells(summaryRow, 6).Value = ""
    ws.Cells(summaryRow, 7).Value = totalPrice
    ws.Cells(summaryRow, 8).Value = totalProfit
    
    ' 합계 행 서식 (굵게, 배경색)
    With ws.Range(ws.Cells(summaryRow, 5), ws.Cells(summaryRow, 8))
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .NumberFormat = "#,##0"
    End With
End Sub

'================================================
' 5. 견적번호 생성
'================================================
Private Function GenerateFileName() As String
    Dim quoteNumber As String
    quoteNumber = "QT-" & Format(Now(), "yyyymmdd-hhmm") & ".xlsx"
    GenerateFileName = quoteNumber
End Function

'================================================
' 6. 견적서 저장 폴더 생성
'================================================
Public Function CreateQuotationFolder() As String
    Dim folderPath As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 현재 파일의 디렉토리 + Quotations 폴더
    folderPath = ThisWorkbook.Path & "\Quotations"
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    CreateQuotationFolder = folderPath
    Set fso = Nothing
End Function

'================================================
' 7. 이전 견적서 목록 보기
'================================================
Public Sub ShowQuotationList()
    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim listText As String
    Dim fileCount As Integer
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = ThisWorkbook.Path & "\Quotations"
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "견적서가 저장된 폴더가 없습니다.", vbInformation, "알림"
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(folderPath)
    fileCount = 0
    listText = "생성된 견적서 목록:" & vbCrLf & vbCrLf
    
    For Each file In folder.Files
        If Right(file.Name, 5) = ".xlsx" Then
            fileCount = fileCount + 1
            listText = listText & fileCount & ". " & file.Name & _
                      " (" & Format(file.DateLastModified, "yyyy-mm-dd hh:mm") & ")" & vbCrLf
        End If
    Next file
    
    If fileCount = 0 Then
        MsgBox "생성된 견적서가 없습니다.", vbInformation, "알림"
    Else
        MsgBox listText & vbCrLf & "저장 위치: " & folderPath, vbInformation, "견적서 목록"
    End If
    
    Set fso = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical, "오류"
End Sub

'================================================
' 8. 폴더 열기
'================================================
Public Sub OpenQuotationFolder()
    Dim folderPath As String
    Dim shell As Object
    
    On Error GoTo ErrorHandler
    
    ' 폴더 생성 (없는 경우)
    folderPath = CreateQuotationFolder()
    
    Set shell = CreateObject("Shell.Application")
    shell.Open folderPath
    
    Set shell = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "폴더를 열 수 없습니다.", vbCritical, "오류"
End Sub
