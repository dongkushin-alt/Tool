Attribute VB_Name = "메인_매크로"
'================================================
' 견적서 자동생성 도구 - 메인 매크로
' 작성자: Quotation Tool
' 작성일: 2026-06-10
'================================================

Option Explicit

' 상수 정의
Const DB_SHEET As String = "부품DB"
Const MARGIN_SHEET As String = "마진설정"
Const QUOTE_SHEET As String = "견적서_작성"
Const QUOTE_TEMP_SHEET As String = "견적서_템플릿"

'================================================
' 1. 부품 정보 조회 함수
'================================================
Public Function SearchPart(partNumber As String) As Variant()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim result(0 To 2) As Variant ' partName, cost, category
    
    Set ws = ThisWorkbook.Sheets(DB_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 부품번호 검색 (A:부품번호, B:부품명, C:카테고리, D:원가)
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = partNumber Then
            result(0) = ws.Cells(i, 2).Value ' 부품명
            result(1) = ws.Cells(i, 4).Value ' 원가
            result(2) = ws.Cells(i, 3).Value ' 카테고리
            SearchPart = result
            Exit Function
        End If
    Next i
    
    ' 검색 실패
    result(0) = "찾을 수 없음"
    result(1) = 0
    result(2) = ""
    SearchPart = result
End Function

'================================================
' 2. 판매가격 계산 함수
'================================================
Public Function CalculatePrice(cost As Double, category As String, quantity As Double) As Double
    Dim ws As Worksheet
    Dim marginRate As Double
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets(MARGIN_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 기본 마진율 (첫 행)
    marginRate = ws.Cells(2, 2).Value ' 전체 기본 마진율
    
    ' 카테고리별 마진율 검색 (A:카테고리, B:마진율)
    For i = 3 To lastRow
        If ws.Cells(i, 1).Value = category Then
            marginRate = ws.Cells(i, 2).Value
            Exit For
        End If
    Next i
    
    ' 판매가 = (원가 * (1 + 마진율)) * 수량
    CalculatePrice = cost * (1 + marginRate) * quantity
End Function

'================================================
' 3. 마진율 조회 함수
'================================================
Public Function GetMarginRate(category As String) As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim defaultMargin As Double
    
    Set ws = ThisWorkbook.Sheets(MARGIN_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 기본 마진율
    defaultMargin = ws.Cells(2, 2).Value
    
    ' 카테고리별 마진율 검색
    For i = 3 To lastRow
        If ws.Cells(i, 1).Value = category Then
            GetMarginRate = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    GetMarginRate = defaultMargin
End Function

'================================================
' 4. 입력값 검증 함수
'================================================
Public Function ValidateInput(partNumber As String, quantity As Double) As Boolean
    If partNumber = "" Then
        MsgBox "부품번호를 입력해주세요", vbExclamation, "입력 오류"
        ValidateInput = False
        Exit Function
    End If
    
    If quantity <= 0 Then
        MsgBox "수량은 0보다 커야 합니다", vbExclamation, "입력 오류"
        ValidateInput = False
        Exit Function
    End If
    
    ValidateInput = True
End Function

'================================================
' 5. 부품 조회 및 자동 입력 (셀 변경 시)
'================================================
Public Sub OnPartNumberChange()
    ' 이 매크로는 부품번호 입력 후 엔터를 누르면 실행됨
    ' 부품DB 시트의 C2:C1000 범위에 데이터 검증 규칙으로 등록
    
    Dim partNumber As String
    Dim searchResult() As Variant
    Dim ws As Worksheet
    Dim row As Long
    
    Set ws = ThisWorkbook.Sheets(QUOTE_SHEET)
    row = ActiveCell.Row
    
    If row < 2 Then Exit Sub
    
    partNumber = ws.Cells(row, 2).Value ' B열: 부품번호
    
    If partNumber <> "" Then
        searchResult = SearchPart(partNumber)
        
        ' 검색 성공 시
        If searchResult(0) <> "찾을 수 없음" Then
            ws.Cells(row, 3).Value = searchResult(0) ' 부품명
            ws.Cells(row, 4).Value = searchResult(1) ' 원가
            ws.Cells(row, 5).Value = searchResult(2) ' 카테고리
            
            MsgBox "부품 정보가 자동 조회되었습니다.", vbInformation, "성공"
        Else
            MsgBox "부품번호를 찾을 수 없습니다. 부품DB를 확인해주세요.", vbExclamation, "조회 실패"
            ws.Cells(row, 2).Value = ""
        End If
    End If
End Sub

'================================================
' 6. 판매가 자동 계산 (수량 입력 후)
'================================================
Public Sub CalculateSalesPrice()
    ' 이 매크로는 수량 입력 후 엔터를 누르면 실행됨
    
    Dim ws As Worksheet
    Dim row As Long
    Dim cost As Double
    Dim quantity As Double
    Dim category As String
    Dim salesPrice As Double
    
    Set ws = ThisWorkbook.Sheets(QUOTE_SHEET)
    row = ActiveCell.Row
    
    If row < 2 Then Exit Sub
    
    cost = ws.Cells(row, 4).Value ' D열: 원가
    quantity = ws.Cells(row, 6).Value ' F열: 수량
    category = ws.Cells(row, 5).Value ' E열: 카테고리
    
    If cost > 0 And quantity > 0 Then
        salesPrice = CalculatePrice(cost, category, quantity)
        ws.Cells(row, 7).Value = salesPrice ' G열: 판매가
        ws.Cells(row, 8).Value = salesPrice - (cost * quantity) ' H열: 이익
    End If
End Sub

'================================================
' 7. 전체 금액 계산
'================================================
Public Sub CalculateTotalPrice()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalPrice As Double
    Dim totalProfit As Double
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets(QUOTE_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "입력된 부품이 없습니다.", vbExclamation, "알림"
        Exit Sub
    End If
    
    totalPrice = 0
    totalProfit = 0
    For i = 2 To lastRow
        If ws.Cells(i, 7).Value <> "" Then
            totalPrice = totalPrice + ws.Cells(i, 7).Value
            totalProfit = totalProfit + ws.Cells(i, 8).Value
        End If
    Next i
    
    MsgBox "합계:" & vbCrLf & _
           "판매가: " & Format(totalPrice, "#,##0") & "원" & vbCrLf & _
           "이익: " & Format(totalProfit, "#,##0") & "원", _
           vbInformation, "견적 요약"
End Sub

'================================================
' 8. 견적서 초기화
'================================================
Public Sub ClearQuote()
    Dim ws As Worksheet
    Dim response As Integer
    
    Set ws = ThisWorkbook.Sheets(QUOTE_SHEET)
    
    response = MsgBox("견적서를 초기화하시겠습니까?", vbYesNo, "확인")
    If response = vbYes Then
        ws.Range("A2:H100").ClearContents
        MsgBox "견적서가 초기화되었습니다.", vbInformation, "완료"
    End If
End Sub

'================================================
' 9. 마진율 설정 가이드
'================================================
Public Sub ShowMarginGuide()
    MsgBox "마진설정 시트에서 마진율을 조정할 수 있습니다." & vbCrLf & vbCrLf & _
           "- 전체 기본 마진율: B2 셀" & vbCrLf & _
           "- 카테고리별 마진율: B3 이하 셀" & vbCrLf & vbCrLf & _
           "마진율은 소수점으로 입력하세요 (예: 0.2 = 20%)", _
           vbInformation, "마진율 설정"
End Sub
