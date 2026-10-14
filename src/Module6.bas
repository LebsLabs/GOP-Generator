Attribute VB_Name = "Module6"
Option Explicit

'=========================================================
' Refreshes the exchange rate row on the data sheet from
' the Frankfurter v2 API. Rates are expressed as the value
' of 1 unit of the currency in USD.
'=========================================================
Public Sub Refresh_Latest_Exchange_Rates()

    'Change this ONCE for each insurance workbook:
    Const DATA_SHEET As String = "ENSA"

    Const SHEET_PASSWORD As String = "DEMO_PASSWORD"
    Const CODE_ROW As Long = 2          'row holding currency codes
    Const RATE_ROW As Long = 3          'row receiving the rates
    Const FIRST_COL As Long = 13        'column M
    Const MAX_COL As Long = 40          'safety stop
    Const STAMP_LABEL_CELL As String = "M4"
    Const STAMP_CELL As String = "N4"   '"rates last refreshed"

    Dim ws As Worksheet
    Dim col As Long
    Dim currencyCode As String
    Dim rateValue As Double
    Dim failedCurrencies As String
    Dim successCount As Long
    Dim attemptCount As Long
    
    Dim savedCalculation As XlCalculation
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)

    savedCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Cursor = xlWait

    failedCurrencies = ""
    successCount = 0
    attemptCount = 0

    'Walk the header row until it runs out of currency codes,
    'so adding a currency needs no code change
    col = FIRST_COL
    Do While col <= MAX_COL
        currencyCode = Trim$(UCase$(CStr(ws.Cells(CODE_ROW, col).Value)))
        If currencyCode = "" Then Exit Do

        attemptCount = attemptCount + 1

        If currencyCode = "USD" Then
            rateValue = 1
        Else
            rateValue = GetExchangeRateToUSD(currencyCode)
        End If

        If rateValue > 0 Then
            'Value only - the sheet owns the number format
            ws.Cells(RATE_ROW, col).Value = rateValue
            successCount = successCount + 1
        Else
            failedCurrencies = failedCurrencies & currencyCode & vbCrLf
        End If

        col = col + 1
    Loop

    'Display the label and timestamp after a complete refresh
    'Create the label shown beside the timestamp
    With ws.Range(STAMP_LABEL_CELL)
        .Value = "Rates last refreshed:"
        .HorizontalAlignment = xlRight
    End With

    'Update timestamp only after a complete successful refresh
    If attemptCount > 0 And successCount = attemptCount Then
        With ws.Range(STAMP_CELL)
            .Value = Now
            .NumberFormat = "dd/mm/yy hh:mm"
            .Interior.Color = RGB(255, 255, 0)
            .HorizontalAlignment = xlCenter
        End With
    End If

    

    Application.Calculation = savedCalculation
    Application.Calculate
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    '=========================================================
    ' Report
    '=========================================================
    If attemptCount = 0 Then
        MsgBox "No currency codes were found in row " & CODE_ROW & _
               " of sheet '" & DATA_SHEET & "'.", _
               vbExclamation, "Exchange Rate Update"

    ElseIf successCount = 0 Then
        MsgBox "No exchange rates could be retrieved." & vbCrLf & vbCrLf & _
               "This usually means there is no internet connection or the " & _
               "rate service is unavailable. The existing rates have been " & _
               "left unchanged.", _
               vbCritical, "Exchange Rate Update Failed"

    ElseIf failedCurrencies = "" Then
        MsgBox "Latest exchange rates were downloaded successfully." & vbCrLf & vbCrLf & _
               successCount & " of " & attemptCount & " currencies updated.", _
               vbInformation, "Exchange Rates Updated"

    Else
        MsgBox successCount & " of " & attemptCount & _
               " currencies were updated." & vbCrLf & vbCrLf & _
               "These could not be retrieved and keep their previous rate:" & _
               vbCrLf & vbCrLf & failedCurrencies, _
               vbExclamation, "Exchange Rate Update"
    End If

    Exit Sub

ErrorHandler:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description

    On Error Resume Next
    
    If savedCalculation <> 0 Then
        Application.Calculation = savedCalculation
    End If
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "The exchange rates could not be updated." & vbCrLf & vbCrLf & _
           "Error " & savedErrorNumber & ":" & vbCrLf & _
           savedErrorDescription, _
           vbCritical, "Exchange Rate Error"
End Sub


'=========================================================
' Returns the value of 1 unit of currencyCode in USD.
' Returns 0 if the rate could not be retrieved.
'=========================================================
Private Function GetExchangeRateToUSD( _
    ByVal currencyCode As String) As Double

    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim ratePosition As Long
    Dim rateSubstring As String
    Dim commaPosition As Long
    Dim bracePosition As Long

    On Error GoTo FailedRequest

    currencyCode = Trim$(UCase$(currencyCode))

    'USD is always 1
    If currencyCode = "USD" Then
        GetExchangeRateToUSD = 1
        Exit Function
    End If

    'Frankfurter v2 - blends 50+ central banks and resolves
    'pegged currencies from the peg, so no manual remapping
    url = "https://api.frankfurter.dev/v2/rate/" & currencyCode & "/USD"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    'Resolve, connect, send, receive - stops Excel hanging
    'indefinitely on a slow or dead endpoint
    http.SetTimeouts 5000, 5000, 10000, 10000

    http.Open "GET", url, False
    http.Send

    If http.status <> 200 Then
        GoTo FailedRequest
    End If

    responseText = http.responseText

    'Example response:
    '{"date":"2026-08-13","base":"GBP","quote":"USD","rate":1.3492}
    ratePosition = InStr(1, responseText, """rate"":", vbTextCompare)
    If ratePosition = 0 Then
        GoTo FailedRequest
    End If

    rateSubstring = Mid$(responseText, ratePosition + Len("""rate"":"))

    commaPosition = InStr(rateSubstring, ",")
    bracePosition = InStr(rateSubstring, "}")

    If commaPosition > 0 Then
        rateSubstring = Left$(rateSubstring, commaPosition - 1)
    ElseIf bracePosition > 0 Then
        rateSubstring = Left$(rateSubstring, bracePosition - 1)
    End If

    rateSubstring = Trim$(rateSubstring)

    'Frankfurter JSON always uses a decimal point.
    'Val handles this safely even if Windows uses comma decimals.
    GetExchangeRateToUSD = val(rateSubstring)

    If GetExchangeRateToUSD <= 0 Then
        GoTo FailedRequest
    End If

    Exit Function

FailedRequest:
    GetExchangeRateToUSD = 0
End Function
