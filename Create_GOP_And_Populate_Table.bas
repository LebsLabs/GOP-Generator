Option Explicit

Public Sub Create_GOP_And_Populate_Table()

    Const FORM_SHEET As String = "FORM"
    Const TEMPLATE_NAME As String = "GOP1.docx"
    Const SHEET_PASSWORD As String = "Password123"

    'Confirmed local OneDrive folder
    Const FOLDER_PATH As String = _
        "C:\Users\...\

    Const wdExportFormatPDF As Long = 17
    Const wdDoNotSaveChanges As Long = 0

    Dim ws As Worksheet
    Dim wsInsurance As Worksheet

    Dim wordApp As Object
    Dim wordDoc As Object

    Dim templatePath As String
    Dim pdfPath As String
    Dim pdfFileName As String

    Dim documentDate As String
    Dim provider As String
    Dim clientName As String
    Dim caseRefNo As String
    Dim insuranceCompany As String
    Dim dateOfService As String
    Dim patientType As String
    Dim currencyCode As String
    Dim amount As String
    Dim coordinator As String

    Dim nextRow As Long
    Dim rowWasAdded As Boolean

    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(FORM_SHEET)

    'Read the FORM values
    documentDate = FormatFormDate(ws.Range("B2").Value)
    provider = Trim$(CStr(ws.Range("B3").Value))
    clientName = Trim$(CStr(ws.Range("B4").Value))
    caseRefNo = Trim$(CStr(ws.Range("B5").Value))
    insuranceCompany = Trim$(CStr(ws.Range("B6").Value))
    dateOfService = FormatFormDate(ws.Range("B7").Value)
    patientType = Trim$(CStr(ws.Range("B8").Value))
    currencyCode = Trim$(CStr(ws.Range("B9").Value))
    amount = FormatFormAmount(ws.Range("B10").Value)
    coordinator = Trim$(CStr(ws.Range("B11").Value))

    'Check that all fields are completed
    If documentDate = "" Or _
       provider = "" Or _
       clientName = "" Or _
       caseRefNo = "" Or _
       insuranceCompany = "" Or _
       dateOfService = "" Or _
       patientType = "" Or _
       currencyCode = "" Or _
       amount = "" Or _
       coordinator = "" Then

        MsgBox "Please complete all fields in cells B2:B11.", _
               vbExclamation, "Missing Information"
        Exit Sub

    End If

    'Select the correct insurance worksheet (validated up front,
    'before anything is created)
    Select Case UCase$(insuranceCompany)

        Case "InsuranceCompany1"
            Set wsInsurance = ThisWorkbook.Worksheets("InsuranceCompany1")

        Case "InsuranceCompany2"
            Set wsInsurance = ThisWorkbook.Worksheets("InsuranceCompany2")

        Case "InsuranceCompany3"
            Set wsInsurance = ThisWorkbook.Worksheets("InsuranceCompany3")

        Case Else

            MsgBox "Please select InsuranceCompany1, InsuranceCompany2 or InsuranceCompany3 in cell B6.", _
                   vbExclamation, "Insurance Company Missing"
            Exit Sub

    End Select
    
    
    '=========================================================
    ' Temporarily unprotect insurance sheet
    '=========================================================
    wsInsurance.Unprotect Password:=SHEET_PASSWORD

    'Check that the local folder exists
    If Dir(FOLDER_PATH, vbDirectory) = "" Then

        MsgBox "The local folder could not be found:" & vbCrLf & vbCrLf & _
               FOLDER_PATH, _
               vbCritical, "Folder Not Found"
        Exit Sub

    End If

    templatePath = FOLDER_PATH & "\" & TEMPLATE_NAME

    'Check that the Word template exists
    If Dir(templatePath) = "" Then

        MsgBox "The Word template could not be found:" & vbCrLf & vbCrLf & _
               templatePath, _
               vbCritical, "Template Not Found"
        Exit Sub

    End If

    'Create the PDF filename
    pdfFileName = "GOP - " & _
                  CleanFileName(caseRefNo) & " - " & _
                  CleanFileName(clientName) & ".pdf"

    pdfPath = FOLDER_PATH & "\" & pdfFileName

    'Ask before replacing an existing PDF
    If Dir(pdfPath) <> "" Then

        If MsgBox( _
            "A PDF with this name already exists:" & vbCrLf & vbCrLf & _
            pdfFileName & vbCrLf & vbCrLf & _
            "Do you want to replace it?", _
            vbQuestion + vbYesNo, _
            "PDF Already Exists") = vbNo Then

            Exit Sub

        End If

        Kill pdfPath

    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Cursor = xlWait

    '--- Step 1: populate the insurance table ---

    nextRow = wsInsurance.Cells(wsInsurance.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 9 Then nextRow = 9

    With wsInsurance

        .Range("A" & nextRow).Value = ws.Range("B5").Value
        .Range("B" & nextRow).Value = ws.Range("B3").Value
        .Range("C" & nextRow).Value = ws.Range("B2").Value
        .Range("D" & nextRow).Value = ws.Range("B7").Value
        .Range("E" & nextRow).Value = ws.Range("B9").Value
        .Range("F" & nextRow).Value = ws.Range("B10").Value

        'Store Coordinator in hidden column AG
        .Range("AG" & nextRow).Value = ws.Range("B11").Value

        .Range("C" & nextRow).NumberFormat = "dd/mm/yyyy"
        .Range("D" & nextRow).NumberFormat = "dd/mm/yyyy"
        .Range("F" & nextRow).NumberFormat = "#,##0.00"

    End With

    rowWasAdded = True

    '--- Step 2: create the PDF ---

    Set wordApp = CreateObject("Word.Application")

    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    Set wordDoc = wordApp.Documents.Open( _
        Filename:=templatePath, _
        ReadOnly:=True, _
        AddToRecentFiles:=False)

    ReplaceWordPlaceholder wordDoc, "<<DATE>>", documentDate
    ReplaceWordPlaceholder wordDoc, "<<PROVIDER>>", provider
    ReplaceWordPlaceholder wordDoc, "<<CLIENT>>", clientName
    ReplaceWordPlaceholder wordDoc, "<<CASE REF NO>>", caseRefNo
    ReplaceWordPlaceholder wordDoc, "<<INSURANCE CO>>", insuranceCompany
    ReplaceWordPlaceholder wordDoc, "<<DATE OF SERVICE>>", dateOfService
    ReplaceWordPlaceholder wordDoc, "<<TYPE PATIENT>>", patientType
    ReplaceWordPlaceholder wordDoc, "<<CURRENCYCODE>>", currencyCode
    ReplaceWordPlaceholder wordDoc, "<<SUM>>", amount
    ReplaceWordPlaceholder wordDoc, "<<COORDINATOR>>", coordinator

    wordDoc.ExportAsFixedFormat _
        OutputFileName:=pdfPath, _
        ExportFormat:=wdExportFormatPDF

    wordDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set wordDoc = Nothing

    wordApp.Quit
    Set wordApp = Nothing

    If Dir(pdfPath) = "" Then

        Err.Raise vbObjectError + 1000, _
                  "Create_GOP_And_Populate_Table", _
                  "The PDF export finished, but the PDF file could not be found."

    End If
    
    '=========================================================
    ' Protect insurance sheet again
    '=========================================================
    wsInsurance.Protect Password:=SHEET_PASSWORD
    

    '--- Step 3: clear the FORM only after both steps succeeded ---

    ws.Range("B2:B11").ClearContents
    ws.Range("B2").Select

    Application.Cursor = xlDefault
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "The GOP was created and logged successfully." & vbCrLf & vbCrLf & _
           "PDF file:" & vbCrLf & pdfFileName & vbCrLf & vbCrLf & _
           "Logged to """ & insuranceCompany & """ sheet, row " & nextRow & ".", _
           vbInformation, "GOP Created"

    Exit Sub

ErrorHandler:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=wdDoNotSaveChanges
    End If

    If Not wordApp Is Nothing Then
        wordApp.Quit
    End If

    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    

    '=========================================================
    ' CLEAN UP INSURANCE TABLE AND REPROTECT SHEET
    '=========================================================
    If Not wsInsurance Is Nothing Then

        wsInsurance.Unprotect Password:=SHEET_PASSWORD

        If rowWasAdded Then
            wsInsurance.Rows(nextRow).ClearContents
        End If

        wsInsurance.Protect Password:=SHEET_PASSWORD

    End If
    
    

    Application.Cursor = xlDefault
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    On Error GoTo 0

    MsgBox "The GOP could not be created." & vbCrLf & vbCrLf & _
           "Error " & errorNumber & ":" & vbCrLf & _
           errorDescription & vbCrLf & vbCrLf & _
           "The FORM data has not been cleared.", _
           vbCritical, "GOP Creation Error"

End Sub


Private Sub ReplaceWordPlaceholder( _
    ByVal wordDoc As Object, _
    ByVal placeholder As String, _
    ByVal replacementText As String)

    Dim storyRange As Object
    Dim currentRange As Object

    For Each storyRange In wordDoc.StoryRanges

        Set currentRange = storyRange

        Do While Not currentRange Is Nothing

            With currentRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting

                .Text = placeholder
                .Replacement.Text = replacementText

                .Forward = True
                .Wrap = 1
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False

                .Execute Replace:=2
            End With

            Set currentRange = currentRange.NextStoryRange

        Loop

    Next storyRange

End Sub


Private Function FormatFormDate(ByVal cellValue As Variant) As String

    If IsDate(cellValue) Then
        FormatFormDate = Format$(CDate(cellValue), "dd/mm/yyyy")
    Else
        FormatFormDate = Trim$(CStr(cellValue))
    End If

End Function


Private Function FormatFormAmount(ByVal cellValue As Variant) As String

    If IsNumeric(cellValue) Then
        FormatFormAmount = Format$(CDbl(cellValue), "#,##0.00")
    Else
        FormatFormAmount = Trim$(CStr(cellValue))
    End If

End Function


Private Function CleanFileName(ByVal textValue As String) As String

    Dim invalidCharacters As Variant
    Dim item As Variant

    invalidCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    CleanFileName = Trim$(textValue)

    For Each item In invalidCharacters
        CleanFileName = Replace(CleanFileName, CStr(item), "-")
    Next item

End Function


Public Sub Check_GOP_Path()

    Const FOLDER_PATH As String = _
        "C:\Users\...\

    If Dir(FOLDER_PATH, vbDirectory) <> "" Then

        MsgBox "The GOP folder was found successfully:" & vbCrLf & vbCrLf & _
               FOLDER_PATH, _
               vbInformation, "Path Confirmed"

    Else

        MsgBox "The GOP folder could not be found:" & vbCrLf & vbCrLf & _
               FOLDER_PATH, _
               vbCritical, "Path Not Found"

    End If

End Sub