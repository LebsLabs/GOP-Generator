Attribute VB_Name = "Module1"
Option Explicit
'=========================================================
' GOP GENERATOR - driven by the hidden MAP sheet.
'
' v3: approval by single-use code.
'  - Amounts over the thresholds no longer accept a Yes/No click.
'    The user must enter a code issued by the approver.
'  - Codes live on the hidden APPROVALS sheet:
'      A Code No | B Code | C Approver | D Max Level |
'      E Status  | F Used On | G Case Ref | H Amount USD | I Entered By
'  - Level 1 (Approver A / Approver B) covers USD 2,000-5,000.
'    Level 2 (Senior Approver) is required above USD 5,000.
'    is also accepted for a Level 1 amount.
'  - A code is only consumed once the PDF exists and the row is
'    logged, so a failed run does not waste a code.
'=========================================================
Private Const FORM_SHEET     As String = "FORM"
Private Const DATA_SHEET     As String = "ENSA"   'change once per insurance workbook
Private Const MAP_SHEET      As String = "MAP"
Private Const APPROVAL_SHEET As String = "APPROVALS"
Private Const SHEET_PASSWORD As String = "DEMO_PASSWORD" 
'PDFs are written here, inside the folder holding this workbook
Private Const PDF_SUBFOLDER  As String = "GOPs"
Private Const FIRST_DATA_ROW As Long = 9
Private Const LAST_DATA_ROW  As Long = 2000
'approval code table
Private Const APPR_FIRST_ROW As Long = 2
Private Const APPR_ATTEMPTS  As Long = 3
Private Const APPR_LOW_STOCK As Long = 5
'FX rate table on the data sheet
Private Const FX_CODES As String = "L2:T2"
Private Const FX_RATES As String = "L3:T3"
'Approval thresholds, in USD
Private Const APPROVAL_LEVEL1 As Double = 2000
Private Const APPROVAL_LEVEL2 As Double = 5000
Private Const wdExportFormatPDF  As Long = 17
Private Const wdDoNotSaveChanges As Long = 0
Private Type FieldDef
    Caption     As String
    nameRef     As String
    EnsaCol     As String
    placeholder As String
    Required    As Boolean
    FieldType   As String
    NumFmt      As String
    RawValue    As Variant
    TextValue   As String
End Type
'=========================================================
' MAIN
'=========================================================
Public Sub Create_GOP_And_Populate_Table()
    Dim ws As Worksheet, wsData As Worksheet, wsMap As Worksheet
    Dim f() As FieldDef
    Dim n As Long, i As Long
    Dim problems As String
    Dim templateName As String, caseRef As String, providerName As String
    Dim baseFolderPath As String, templatePath As String, pdfFolderPath As String
    Dim pdfPath As String, pdfFileName As String
    Dim rate As Double, amountUsd As Double
    Dim m As Variant
    Dim nextRow As Long
    Dim rowWasAdded As Boolean
    Dim wordApp As Object, wordDoc As Object
    Dim errNum As Long, errDesc As String
    Dim apprRow As Long, apprName As String, apprCode As String
    
    'added additionally for PDF store path
    Dim tempPdfPath As String
    Dim exportErr As Long
    
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Worksheets(FORM_SHEET)
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    Set wsMap = ThisWorkbook.Worksheets(MAP_SHEET)
    '--- 1. load the mapping table -------------------------
    n = LoadFieldMap(wsMap, f)
    If n = 0 Then
        MsgBox "No fields found on the " & MAP_SHEET & " sheet.", _
               vbCritical, "Mapping Table Empty"
        Exit Sub
    End If
    '--- 2. read the form via named ranges -----------------
    For i = 1 To n
        If Not ReadField(f(i)) Then
            MsgBox "The defined name '" & f(i).nameRef & "' (field " & i & _
                   " on " & MAP_SHEET & ") does not exist in this workbook.", _
                   vbCritical, "Field Name Not Found"
            Exit Sub
        End If
    Next i
    '--- 3. tidy the text on the way in --------------------
    SetFieldText f, n, "frm_CaseRef", UCase$(FieldText(f, n, "frm_CaseRef"))
    SetFieldText f, n, "frm_Provider", ProperName(FieldText(f, n, "frm_Provider"))
    SetFieldText f, n, "frm_Patient", ProperName(FieldText(f, n, "frm_Patient"))
    '--- 4. completeness and type checks -------------------
    For i = 1 To n
        If f(i).Required And Len(f(i).TextValue) = 0 Then
            problems = problems & vbCrLf & "   - " & f(i).Caption & " is empty"
        ElseIf Len(f(i).TextValue) > 0 Then
            Select Case UCase$(f(i).FieldType)
                Case "DATE"
                    If Not IsDate(f(i).RawValue) Then
                        problems = problems & vbCrLf & "   - " & f(i).Caption & _
                                   " is not a valid date"
                    End If
                Case "NUMBER"
                    If Not IsNumeric(f(i).RawValue) Then
                        problems = problems & vbCrLf & "   - " & f(i).Caption & _
                                   " must be a number"
                    ElseIf CDbl(f(i).RawValue) <= 0 Then
                        problems = problems & vbCrLf & "   - " & f(i).Caption & _
                                   " must be greater than zero"
                    End If
            End Select
        End If
    Next i
    If Len(problems) > 0 Then
        MsgBox "Please correct the following before generating a GOP:" & _
               vbCrLf & problems, vbExclamation, "Missing Or Invalid Information"
        Exit Sub
    End If
    templateName = FieldText(f, n, "frm_Template")
    caseRef = FieldText(f, n, "frm_CaseRef")
    providerName = FieldText(f, n, "frm_Provider")
    '--- 5. currency must exist in the rate table ----------
    m = Application.Match(FieldText(f, n, "frm_Currency"), _
                          wsData.Range(FX_CODES), 0)
    If IsError(m) Then
        MsgBox "The currency '" & FieldText(f, n, "frm_Currency") & _
               "' is not in the rate table on " & DATA_SHEET & " " & FX_CODES & _
               "." & vbCrLf & vbCrLf & _
               "Add it there first, otherwise the USD conversion will not work.", _
               vbExclamation, "Currency Not In Rate Table"
        Exit Sub
    End If
    rate = CDbl(wsData.Range(FX_RATES).Cells(1, CLng(m)).Value)
    amountUsd = CDbl(FieldRaw(f, n, "frm_Amount")) * rate
    '--- 6. approval by single-use code --------------------
    If amountUsd > APPROVAL_LEVEL1 Then
        If Not GetApproval(amountUsd, caseRef, apprRow, apprName, apprCode) Then
            Exit Sub                       'user cancelled or code not accepted
        End If
    End If
    '--- 7. duplicate case reference -----------------------
    If Application.CountIf(wsData.Range("A" & FIRST_DATA_ROW & ":A" & LAST_DATA_ROW), _
                           caseRef) > 0 Then
        If MsgBox("Case reference " & caseRef & " is already in the " & DATA_SHEET & _
                  " register." & vbCrLf & vbCrLf & "Log it again anyway?", _
                  vbQuestion + vbYesNo, "Duplicate Case Reference") = vbNo Then
            Exit Sub
        End If
    End If
    '--- 8. date sanity ------------------------------------
    If IsDate(FieldRaw(f, n, "frm_DOS")) And IsDate(FieldRaw(f, n, "frm_DateGOP")) Then
        If CDate(FieldRaw(f, n, "frm_DOS")) > CDate(FieldRaw(f, n, "frm_DateGOP")) Then
            If MsgBox("The Date of Service is later than the Date of GOP." & _
                      vbCrLf & vbCrLf & "Continue?", vbQuestion + vbYesNo, _
                      "Check The Dates") = vbNo Then Exit Sub
        End If
    End If
    If IsDate(FieldRaw(f, n, "frm_DOB")) Then
        If CDate(FieldRaw(f, n, "frm_DOB")) >= Date Then
            MsgBox "The Date of Birth is today or in the future.", _
                   vbExclamation, "Check The Date Of Birth"
            Exit Sub
        End If
    End If
    '--- 9. locate the template ----------------------------
    baseFolderPath = BaseFolder()
    templatePath = baseFolderPath & "\" & templateName
    If LCase$(Right$(templatePath, 5)) <> ".docx" Then
        templatePath = templatePath & ".docx"
    End If
    If Not FileThere(templatePath) Then
        MsgBox "The Word template could not be found:" & vbCrLf & vbCrLf & _
               templatePath & vbCrLf & vbCrLf & _
               "Check the name in the 'Select the GOP to be filled out' list " & _
               "and make sure the document sits in that folder.", _
               vbCritical, "Template Not Found"
        Exit Sub
    End If
    '--- 10. PDF folder, name and overwrite check ----------
    pdfFolderPath = baseFolderPath & "\" & PDF_SUBFOLDER
    If Not FolderThere(pdfFolderPath) Then
        MakeFolder pdfFolderPath
        If Not FolderThere(pdfFolderPath) Then
            pdfFolderPath = baseFolderPath        'fall back to the main folder
        End If
    End If
    pdfFileName = "GOP - " & CleanFileName(caseRef) & " - " & _
                  CleanFileName(providerName) & ".pdf"
    pdfPath = pdfFolderPath & "\" & pdfFileName
    If FileThere(pdfPath) Then
        If MsgBox("A PDF with this name already exists:" & vbCrLf & vbCrLf & _
                  pdfFileName & vbCrLf & vbCrLf & "Replace it?", _
                  vbQuestion + vbYesNo, "PDF Already Exists") = vbNo Then Exit Sub
        If Not DeleteFileQuiet(pdfPath) Then
            MsgBox "The existing PDF could not be replaced. Close it if it is " & _
                   "open, then try again:" & vbCrLf & vbCrLf & pdfFileName, _
                   vbExclamation, "PDF In Use"
            Exit Sub
        End If
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Cursor = xlWait
   
    '--- 11. write the register row ------------------------
    nextRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < FIRST_DATA_ROW Then nextRow = FIRST_DATA_ROW
    If nextRow > LAST_DATA_ROW Then
        
        RestoreExcel
        MsgBox "The register is full to row " & LAST_DATA_ROW & _
               ". Extend the formatting and formulas before adding more cases.", _
               vbExclamation, "Register Full"
        Exit Sub
    End If
    For i = 1 To n
        If Len(f(i).EnsaCol) > 0 Then
            With wsData.Range(f(i).EnsaCol & nextRow)
                Select Case UCase$(f(i).FieldType)
                    Case "DATE", "NUMBER": .Value = f(i).RawValue
                    Case Else:             .Value = f(i).TextValue
                End Select
                If Len(f(i).NumFmt) > 0 Then .NumberFormat = f(i).NumFmt
            End With
        End If
    Next i
    rowWasAdded = True
    
    
    '--- 12. build the PDF ---------------------------------
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    On Error Resume Next
    wordApp.AutomationSecurity = 1          'msoAutomationSecurityLow, no macro prompt
    On Error GoTo ErrorHandler
    'A NEW document based on the template. Opening the template itself
    'ReadOnly makes Word refuse the find-and-replace on some builds.
    Set wordDoc = wordApp.Documents.Add(Template:=templatePath, _
                                        NewTemplate:=False, _
                                        Visible:=False)
    For i = 1 To n
        If Len(f(i).placeholder) > 0 Then
            ReplaceWordPlaceholder wordDoc, f(i).placeholder, f(i).TextValue
        End If
    Next i
    'Write the PDF to a local temp folder first. Word often cannot export
    'straight into a synced SharePoint / OneDrive folder - that is the
    'usual source of "Error 4198: Command failed".
    tempPdfPath = TempFolder() & "\" & pdfFileName
    DeleteFileQuiet tempPdfPath
    exportErr = 0
    On Error Resume Next
    wordDoc.ExportAsFixedFormat OutputFileName:=tempPdfPath, _
                                ExportFormat:=wdExportFormatPDF
    exportErr = Err.Number
    Err.Clear
    If Not FileThere(tempPdfPath) Then
        'fallback route to PDF
        wordDoc.SaveAs2 fileName:=tempPdfPath, FileFormat:=wdExportFormatPDF
        exportErr = Err.Number
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    If Not FileThere(tempPdfPath) Then
        Err.Raise vbObjectError + 1001, "Create_GOP_And_Populate_Table", _
            "Word could not create the PDF (Word error " & exportErr & ")." & _
            vbCrLf & "Check that the template opens normally and that no " & _
            "other copy of Word is stuck in Task Manager."
    End If
    wordDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set wordDoc = Nothing
    wordApp.Quit
    Set wordApp = Nothing
    'Now move it into the GOPs folder. Excel copies into synced folders
    'reliably, which Word does not.
    If Not CopyFileQuiet(tempPdfPath, pdfPath) Then
        Err.Raise vbObjectError + 1002, "Create_GOP_And_Populate_Table", _
            "The PDF was created but could not be copied into:" & vbCrLf & _
            pdfFolderPath & vbCrLf & vbCrLf & _
            "Your copy is here and has not been lost:" & vbCrLf & tempPdfPath & _
            vbCrLf & vbCrLf & "Check that the GOPs folder is fully synced " & _
            "(green tick, not a pending arrow)."
    End If
    DeleteFileQuiet tempPdfPath
    If Not FileThere(pdfPath) Then
        Err.Raise vbObjectError + 1000, "Create_GOP_And_Populate_Table", _
                  "The export finished but the PDF could not be found."
    End If
    
    '--- 12b. burn the approval code ----------------------
    If apprRow > 0 Then
        ConsumeApprovalCode apprRow, caseRef, amountUsd
    End If
    
    '--- 13. clear the form, keep the template selection ---
    For i = 1 To n
        If f(i).nameRef <> "frm_Template" Then
            ThisWorkbook.Names(f(i).nameRef).RefersToRange.ClearContents
        End If
    Next i
    RestoreExcel
    On Error Resume Next
    ws.Activate
    ThisWorkbook.Names("frm_DateGOP").RefersToRange.Select
    On Error GoTo 0
    MsgBox "The GOP was created and logged." & vbCrLf & vbCrLf & _
           "PDF:  " & pdfFileName & vbCrLf & _
           "Saved in:  " & pdfFolderPath & vbCrLf & _
           "Value:  " & Format$(amountUsd, "#,##0.00") & " USD" & vbCrLf & _
           "Logged to row " & nextRow & " of " & DATA_SHEET & "." & _
           IIf(apprRow > 0, vbCrLf & "Approved by:  " & apprName & _
               "  (code " & apprCode & ", now used)", ""), _
           vbInformation, "GOP Created"
    Exit Sub
ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=wdDoNotSaveChanges
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    If Not wsData Is Nothing Then
        
        If rowWasAdded Then
            For i = 1 To n
                If Len(f(i).EnsaCol) > 0 Then
                    wsData.Range(f(i).EnsaCol & nextRow).ClearContents
                End If
            Next i
        End If
        
    End If
    RestoreExcel
    On Error GoTo 0
    MsgBox "The GOP could not be created." & vbCrLf & vbCrLf & _
           "Error " & errNum & ":" & vbCrLf & errDesc & vbCrLf & vbCrLf & _
           "The form has not been cleared. " & _
           "The approval code has NOT been used and can be entered again.", _
           vbCritical, "GOP Creation Error"
End Sub
Private Sub RestoreExcel()
    On Error Resume Next
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
End Sub
'=========================================================
' APPROVAL CODES
'=========================================================
' Asks for a code, validates it, and hands back the row that
' will be burned once the GOP is safely created.
Private Function GetApproval(amountUsd As Double, caseRef As String, _
                             ByRef apprRow As Long, ByRef apprName As String, _
                             ByRef apprCode As String) As Boolean
    Dim requiredLevel As Long, whoCanApprove As String
    Dim attempt As Long, entry As String, problem As String
    Dim remaining As Long
    apprRow = 0: apprName = "": apprCode = ""
    If amountUsd > APPROVAL_LEVEL2 Then
        requiredLevel = 2
        whoCanApprove = "Sean"
    Else
        requiredLevel = 1
        whoCanApprove = "Insaf, Lukanyo or Sean"
    End If
    For attempt = 1 To APPR_ATTEMPTS
        entry = InputBox( _
            "This GOP is " & Format$(amountUsd, "#,##0.00") & " USD." & vbCrLf & _
            "Case reference: " & caseRef & vbCrLf & vbCrLf & _
            "It needs approval from " & whoCanApprove & "." & vbCrLf & vbCrLf & _
            "Ask the approver for one approval code from their list and type " & _
            "that number below." & vbCrLf & _
            "Each code works once only." & vbCrLf & vbCrLf & _
            "Attempt " & attempt & " of " & APPR_ATTEMPTS & ".", _
            "Approval Code Required")
        If StrPtr(entry) = 0 Then GetApproval = False: Exit Function   'Cancel
        entry = CleanCodeEntry(entry)
        If Len(entry) = 0 Then
            problem = "No code was entered."
        Else
            problem = CheckApprovalCode(entry, requiredLevel, apprRow, apprName)
        End If
        If Len(problem) = 0 Then
            apprCode = entry
            GetApproval = True
            remaining = CodesRemaining(apprName)
            If remaining <= APPR_LOW_STOCK Then
                MsgBox apprName & " has only " & remaining & " unused approval " & _
                       "code(s) left. Issue a new batch soon.", _
                       vbInformation, "Approval Codes Running Low"
            End If
            Exit Function
        End If
        If attempt < APPR_ATTEMPTS Then
            MsgBox problem & vbCrLf & vbCrLf & "Please try again.", _
                   vbExclamation, "Approval Code Not Accepted"
        Else
            MsgBox problem & vbCrLf & vbCrLf & _
                   "No more attempts. The GOP has not been created and nothing " & _
                   "has been logged.", vbCritical, "Approval Failed"
        End If
    Next attempt
    apprRow = 0: apprName = ""
    GetApproval = False
End Function
' Returns "" when the code is good, otherwise the reason it was refused.
Private Function CheckApprovalCode(entry As String, requiredLevel As Long, _
                                   ByRef apprRow As Long, _
                                   ByRef apprName As String) As String
    Dim wsA As Worksheet, lastRow As Long, r As Long
    Dim cellCode As String, status As String, lvl As Long
    apprRow = 0: apprName = ""
    Set wsA = ApprovalSheet()
    If wsA Is Nothing Then
        CheckApprovalCode = "The " & APPROVAL_SHEET & " sheet is missing, so the " & _
                            "code cannot be checked."
        Exit Function
    End If
    If Not IsNumeric(entry) Then
        CheckApprovalCode = "'" & entry & "' is not a number. An approval code " & _
                            "is a number from the approver's list."
        Exit Function
    End If
    lastRow = wsA.Cells(wsA.Rows.Count, "B").End(xlUp).Row
    For r = APPR_FIRST_ROW To lastRow
        cellCode = CleanCodeEntry(CStr(wsA.Cells(r, "B").Value))
        If Len(cellCode) > 0 Then
            If StrComp(cellCode, entry, vbTextCompare) = 0 Then
                apprName = Trim$(CStr(wsA.Cells(r, "C").Value))
                lvl = val(CStr(wsA.Cells(r, "D").Value))
                status = UCase$(Trim$(CStr(wsA.Cells(r, "E").Value)))
                If status <> "AVAILABLE" Then
                    CheckApprovalCode = "That code has already been used" & _
                        IIf(Len(CStr(wsA.Cells(r, "G").Value)) > 0, _
                            " for case " & CStr(wsA.Cells(r, "G").Value), "") & _
                        IIf(IsDate(wsA.Cells(r, "F").Value), _
                            " on " & Format$(wsA.Cells(r, "F").Value, _
                            "dd/mm/yyyy hh:mm"), "") & "." & vbCrLf & _
                        "Ask " & apprName & " for a fresh code."
                    apprRow = 0: apprName = ""
                    Exit Function
                End If
                If lvl < requiredLevel Then
                    CheckApprovalCode = "That code belongs to " & apprName & _
                        ", who can approve up to " & _
                        Format$(APPROVAL_LEVEL2, "#,##0") & " USD." & vbCrLf & _
                        "This amount needs a code from Sean."
                    apprRow = 0: apprName = ""
                    Exit Function
                End If
                apprRow = r
                CheckApprovalCode = ""
                Exit Function
            End If
        End If
    Next r
    CheckApprovalCode = "Code " & entry & " is not on the approval list."
End Function
' Marks the code used and records what it was used for.
Private Sub ConsumeApprovalCode(apprRow As Long, caseRef As String, _
                                amountUsd As Double)
    Dim wsA As Worksheet
    Set wsA = ApprovalSheet()
    If wsA Is Nothing Or apprRow < APPR_FIRST_ROW Then Exit Sub
    On Error Resume Next
    wsA.Unprotect Password:=SHEET_PASSWORD
    wsA.Cells(apprRow, "E").Value = "Used"
    wsA.Cells(apprRow, "F").Value = Now
    wsA.Cells(apprRow, "F").NumberFormat = "dd/mm/yyyy hh:mm"
    wsA.Cells(apprRow, "G").Value = caseRef
    wsA.Cells(apprRow, "H").Value = amountUsd
    wsA.Cells(apprRow, "H").NumberFormat = "#,##0.00"
    wsA.Cells(apprRow, "I").Value = Application.UserName
    wsA.Protect Password:=SHEET_PASSWORD
    On Error GoTo 0
End Sub
Private Function CodesRemaining(apprName As String) As Long
    Dim wsA As Worksheet, lastRow As Long, r As Long
    Set wsA = ApprovalSheet()
    If wsA Is Nothing Then Exit Function
    lastRow = wsA.Cells(wsA.Rows.Count, "B").End(xlUp).Row
    For r = APPR_FIRST_ROW To lastRow
        If StrComp(Trim$(CStr(wsA.Cells(r, "C").Value)), apprName, vbTextCompare) = 0 Then
            If UCase$(Trim$(CStr(wsA.Cells(r, "E").Value))) = "AVAILABLE" Then
                CodesRemaining = CodesRemaining + 1
            End If
        End If
    Next r
End Function
Private Function ApprovalSheet() As Worksheet
    On Error Resume Next
    Set ApprovalSheet = ThisWorkbook.Worksheets(APPROVAL_SHEET)
    On Error GoTo 0
End Function
' Strips spaces, dashes and any stray non-digits from a typed code.
Private Function CleanCodeEntry(s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then CleanCodeEntry = CleanCodeEntry & ch
    Next i
End Function
'=========================================================
' MAPPING TABLE
'=========================================================
Private Function LoadFieldMap(wsMap As Worksheet, f() As FieldDef) As Long
    Dim lastRow As Long, r As Long, k As Long
    lastRow = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then LoadFieldMap = 0: Exit Function
    ReDim f(1 To lastRow - 1)
    For r = 2 To lastRow
        'stop at the first blank row - anything below is not a field
        If Len(Trim$(CStr(wsMap.Cells(r, "A").Value))) = 0 Then Exit For
        'a field must have both a caption and a defined name
        If Len(Trim$(CStr(wsMap.Cells(r, "B").Value))) > 0 Then
            k = k + 1
            f(k).Caption = Trim$(CStr(wsMap.Cells(r, "A").Value))
            f(k).nameRef = Trim$(CStr(wsMap.Cells(r, "B").Value))
            f(k).EnsaCol = Trim$(CStr(wsMap.Cells(r, "C").Value))
            f(k).placeholder = Trim$(CStr(wsMap.Cells(r, "D").Value))
            f(k).Required = (UCase$(Trim$(CStr(wsMap.Cells(r, "E").Value))) = "YES")
            f(k).FieldType = Trim$(CStr(wsMap.Cells(r, "F").Value))
            f(k).NumFmt = Trim$(CStr(wsMap.Cells(r, "G").Value))
        End If
    Next r
    LoadFieldMap = k
End Function
Private Function ReadField(fd As FieldDef) As Boolean
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names(fd.nameRef).RefersToRange
    On Error GoTo 0
    If r Is Nothing Then ReadField = False: Exit Function
    fd.RawValue = r.Value
    Select Case UCase$(fd.FieldType)
        Case "DATE":   fd.TextValue = FormatFormDate(r.Value)
        Case "NUMBER": fd.TextValue = FormatFormAmount(r.Value)
        Case Else:     fd.TextValue = Application.Trim(CStr(r.Value))
    End Select
    ReadField = True
End Function
Private Function FieldIndex(f() As FieldDef, n As Long, nameRef As String) As Long
    Dim i As Long
    For i = 1 To n
        If StrComp(f(i).nameRef, nameRef, vbTextCompare) = 0 Then
            FieldIndex = i: Exit Function
        End If
    Next i
End Function
Private Function FieldText(f() As FieldDef, n As Long, nameRef As String) As String
    Dim i As Long: i = FieldIndex(f, n, nameRef)
    If i > 0 Then FieldText = f(i).TextValue
End Function
Private Function FieldRaw(f() As FieldDef, n As Long, nameRef As String) As Variant
    Dim i As Long: i = FieldIndex(f, n, nameRef)
    If i > 0 Then FieldRaw = f(i).RawValue
End Function
Private Sub SetFieldText(f() As FieldDef, n As Long, nameRef As String, v As String)
    Dim i As Long: i = FieldIndex(f, n, nameRef)
    If i > 0 Then f(i).TextValue = v
End Sub
'=========================================================
' FILE SYSTEM  (FileSystemObject only - never Dir)
'=========================================================
Private Function FSO() As Object
    Static o As Object
    If o Is Nothing Then Set o = CreateObject("Scripting.FileSystemObject")
    Set FSO = o
End Function
Private Function FileThere(path As String) As Boolean
    On Error Resume Next
    If Len(path) > 0 Then FileThere = FSO.FileExists(path)
    On Error GoTo 0
End Function
Private Function FolderThere(path As String) As Boolean
    On Error Resume Next
    If Len(path) > 0 Then FolderThere = FSO.FolderExists(path)
    On Error GoTo 0
End Function
Private Sub MakeFolder(path As String)
    On Error Resume Next
    FSO.CreateFolder path
    On Error GoTo 0
End Sub
Private Function DeleteFileQuiet(path As String) As Boolean
    On Error Resume Next
    FSO.DeleteFile path, True
    DeleteFileQuiet = Not FSO.FileExists(path)
    On Error GoTo 0
End Function
Private Function TempFolder() As String
    Dim p As String
    On Error Resume Next
    p = Environ$("TEMP")
    If Len(p) = 0 Then p = Environ$("TMP")
    If Len(p) = 0 Then p = FSO.GetSpecialFolder(2).path   'TemporaryFolder
    On Error GoTo 0
    TempFolder = StripSlash(p)
End Function
Private Function CopyFileQuiet(sourcePath As String, destPath As String) As Boolean
    On Error Resume Next
    FSO.CopyFile sourcePath, destPath, True
    CopyFileQuiet = FSO.FileExists(destPath)
    On Error GoTo 0
End Function


'=========================================================
' FOLDER RESOLUTION
'=========================================================
Private Function BaseFolder() As String
    Dim p As String, cached As String, found As String
    Dim fd As Object
    '1. the workbook is open from a real Windows folder
    p = ThisWorkbook.path
    If Len(p) > 0 And LCase$(Left$(p, 4)) <> "http" Then
        BaseFolder = StripSlash(p)
        Exit Function
    End If
    '2. a folder chosen earlier and still valid on this machine
    cached = ""
    On Error Resume Next
    cached = Trim$(CStr(ThisWorkbook.Names("cfg_Folder").RefersToRange.Value))
    On Error GoTo 0
    If FolderHoldsThisWorkbook(cached) Then
        BaseFolder = StripSlash(cached)
        Exit Function
    End If
    '3. find the synced / shortcut copy of this library
    found = FindSyncedFolder()
    If Len(found) > 0 Then
        SaveFolderChoice found
        BaseFolder = StripSlash(found)
        Exit Function
    End If
    '4. ask once, then remember it on MAP
    Set fd = Application.FileDialog(4)      'msoFileDialogFolderPicker
    fd.Title = "Select the folder holding the GOP templates"
    fd.InitialFileName = Environ$("USERPROFILE") & "\"
    If fd.Show <> -1 Then
        Err.Raise vbObjectError + 2000, "BaseFolder", _
            "No folder was selected, so the GOP could not be saved."
    End If
    found = fd.SelectedItems(1)
    If Not FolderHoldsThisWorkbook(found) Then
        If MsgBox("That folder does not contain " & ThisWorkbook.Name & "." & _
                  vbCrLf & vbCrLf & "Use it anyway?", vbQuestion + vbYesNo, _
                  "Check The Folder") = vbNo Then
            Err.Raise vbObjectError + 2001, "BaseFolder", _
                "Folder selection cancelled."
        End If
    End If
    SaveFolderChoice found
    BaseFolder = StripSlash(found)
End Function
Private Sub SaveFolderChoice(path As String)
    On Error Resume Next
    ThisWorkbook.Names("cfg_Folder").RefersToRange.Value = StripSlash(path)
    On Error GoTo 0
End Sub
' A folder is the right one if this workbook sits in it.
Private Function FolderHoldsThisWorkbook(folderPath As String) As Boolean
    If Len(folderPath) = 0 Then Exit Function
    FolderHoldsThisWorkbook = _
        FileThere(StripSlash(folderPath) & "\" & ThisWorkbook.Name)
End Function
Private Function FindSyncedFolder() As String
    Dim roots As Collection, i As Long
    Dim fld As Object, sub1 As Object
    Dim reg As Object, keys As Variant
    Dim mp As String, root As String
    Set roots = New Collection
    On Error Resume Next
    If Len(Environ$("OneDriveCommercial")) > 0 Then roots.Add Environ$("OneDriveCommercial")
    If Len(Environ$("OneDrive")) > 0 Then roots.Add Environ$("OneDrive")
    If Len(Environ$("OneDriveConsumer")) > 0 Then roots.Add Environ$("OneDriveConsumer")
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    If Not reg Is Nothing Then
        reg.EnumKey &H80000001, "Software\SyncEngines\Providers\OneDrive", keys
        If IsArray(keys) Then
            For i = LBound(keys) To UBound(keys)
                mp = ""
                reg.GetStringValue &H80000001, _
                    "Software\SyncEngines\Providers\OneDrive\" & keys(i), _
                    "MountPoint", mp
                If Len(mp) > 0 Then roots.Add mp
            Next i
        End If
    End If
    On Error GoTo 0
    For i = 1 To roots.Count
        root = StripSlash(CStr(roots(i)))
        If FolderHoldsThisWorkbook(root) Then
            FindSyncedFolder = root
            Exit Function
        End If
        Set fld = Nothing
        On Error Resume Next
        If FSO.FolderExists(root) Then Set fld = FSO.GetFolder(root)
        On Error GoTo 0
        If Not fld Is Nothing Then
            On Error Resume Next
            For Each sub1 In fld.SubFolders
                If FolderHoldsThisWorkbook(sub1.path) Then
                    FindSyncedFolder = sub1.path
                    On Error GoTo 0
                    Exit Function
                End If
            Next sub1
            On Error GoTo 0
        End If
    Next i
End Function
'=========================================================
' HELPERS
'=========================================================
Private Function StripSlash(s As String) As String
    StripSlash = s
    Do While Len(StripSlash) > 0 And Right$(StripSlash, 1) = "\"
        StripSlash = Left$(StripSlash, Len(StripSlash) - 1)
    Loop
End Function
Private Function FormatFormDate(val As Variant) As String
    If IsDate(val) Then
        FormatFormDate = Format$(val, "dd/mm/yyyy")
    Else
        FormatFormDate = Application.Trim(CStr(val))
    End If
End Function
Private Function FormatFormAmount(val As Variant) As String
    If IsEmpty(val) Then Exit Function
    If IsNumeric(val) Then
        FormatFormAmount = Format$(CDbl(val), "#,##0.00")
    Else
        FormatFormAmount = Application.Trim(CStr(val))
    End If
End Function
Private Function ProperName(s As String) As String
    If Len(s) = 0 Then Exit Function
    ProperName = Application.Trim(Application.Proper(s))
End Function
Private Function CleanFileName(fileName As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    CleanFileName = fileName
    For i = LBound(bad) To UBound(bad)
        CleanFileName = Replace(CleanFileName, bad(i), "")
    Next i
    CleanFileName = Application.Trim(CleanFileName)
End Function
' Replaces the placeholder in the body and in every header and footer.
Private Sub ReplaceWordPlaceholder(doc As Object, placeholder As String, _
                                   replaceWithText As String)
    Dim sec As Object, hf As Object
    FindReplaceIn doc.Content, placeholder, replaceWithText
    On Error Resume Next
    For Each sec In doc.Sections
        For Each hf In sec.Headers
            FindReplaceIn hf.Range, placeholder, replaceWithText
        Next hf
        For Each hf In sec.Footers
            FindReplaceIn hf.Range, placeholder, replaceWithText
        Next hf
    Next sec
    On Error GoTo 0
End Sub
Private Sub FindReplaceIn(rng As Object, placeholder As String, _
                          replaceWithText As String)
    On Error Resume Next
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = placeholder
        .Replacement.Text = replaceWithText
        .Forward = True
        .Wrap = 1                    'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=2          'wdReplaceAll
    End With
    On Error GoTo 0
End Sub
'=========================================================
' RUN ONCE: unlock the manual-entry columns, then protect.
'=========================================================

Public Sub Prepare_ENSA_Protection()
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    If wsData.ProtectContents Then wsData.Unprotect Password:=SHEET_PASSWORD
    wsData.Cells.Locked = True
    wsData.Range("E2").Locked = False              'alert recipient
    wsData.Range("W2").Locked = False              'total fund
    wsData.Range("L2:T3").Locked = False           'currency codes + rates
    wsData.Range("T4").Locked = False              'rates-refreshed stamp
    wsData.Range("A" & FIRST_DATA_ROW & ":G" & LAST_DATA_ROW).Locked = False
        wsData.Range("I" & FIRST_DATA_ROW & ":J" & LAST_DATA_ROW).Locked = False
        wsData.Range("M" & FIRST_DATA_ROW & ":M" & LAST_DATA_ROW).Locked = False
        wsData.Range("O" & FIRST_DATA_ROW & ":U" & LAST_DATA_ROW).Locked = False
        wsData.Range("W" & FIRST_DATA_ROW & ":AC" & LAST_DATA_ROW).Locked = False
        wsData.Range("AE" & FIRST_DATA_ROW & ":AL" & LAST_DATA_ROW).Locked = False
        wsData.Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True, _
                   AllowFiltering:=True, AllowSorting:=True, _
                   AllowInsertingRows:=True, AllowDeletingRows:=True, _
                   AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                   AllowFormattingRows:=True
    MsgBox "Manual entry columns unlocked. Formula columns stay protected.", _
           vbInformation, "ENSA Protection Set"
End Sub
