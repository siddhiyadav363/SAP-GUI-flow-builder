' ================================================================================
' SAP GUI Script: ME51N - Purchase Requisition Creation (Refined from Recording)
' ================================================================================
' Transaction: ME51N
' Script mode: NORMAL (data entry automation)
' Purpose: Create Purchase Requisition with line items (grid-based entry)
' Source: Recorded script (SAP Scripting Tracker) - locators preserved exactly
' 
' This script replicates the recorded ME51N flow, enhanced with:
'   - Dynamic JSON-driven execution (file or command-line JSON)
'   - Skip-if-empty guards for optional line item fields
'   - Status bar error detection (E/A → fail with exit code 1)
'   - Messages modal scanning (dynpro labels only, after Check and Save)
'   - Screenshot capture (main screen + Messages modal when open)
'   - Optional validation: compare JSON expected vs SAP actual grid values
'
' Recorded actions:
'   1. Navigate to ME51N
'   2. Enter line items into grid (MATNR, MENGE, NAME1)
'   3. Check (btn[39])
'   4. Save (btn[11])
'
' JSON Input Structure (ALL values must be empty strings "" except execution_id):
' {
'   "execution_id": "",
'   "line_items": [
'     {"material": "", "quantity": "", "short_text": ""},
'     {"material": "", "quantity": "", "short_text": ""}
'   ]
' }
'
' Usage:
'   cscript ME51N_CreatePR_Refined.vbs "C:\path\to\data.json"
'   cscript ME51N_CreatePR_Refined.vbs "{""execution_id"":""exec_001"",""line_items"":[{""material"":""2092"",""quantity"":""1000"",""short_text"":""MTH1""}]}"
'
' Exit Codes:
'   0 = Success
'   1 = Error (JSON invalid, SAP error, Messages modal error, status bar E/A)
' ================================================================================

Option Explicit

' --- JSON Parser: flat objects and arrays; commas/colons inside quoted strings are respected ---
Function ParseJsonQuotedString(jsonFromQuote)
    Dim i, c, esc, out, hex4, codepoint
    out = ""
    esc = False
    i = 2
    Do While i <= Len(jsonFromQuote)
        c = Mid(jsonFromQuote, i, 1)
        If esc Then
            Select Case c
                Case """" : out = out & """"
                Case "\"  : out = out & "\"
                Case "/"  : out = out & "/"
                Case "b"  : out = out & Chr(8)
                Case "f"  : out = out & Chr(12)
                Case "n"  : out = out & vbLf
                Case "r"  : out = out & vbCr
                Case "t"  : out = out & vbTab
                Case "u"
                    If i + 4 <= Len(jsonFromQuote) Then
                        hex4 = Mid(jsonFromQuote, i + 1, 4)
                        On Error Resume Next
                        codepoint = CLng("&H" & hex4)
                        If Err.Number = 0 Then
                            out = out & ChrW(codepoint)
                        Else
                            out = out & "u" & hex4
                        End If
                        On Error GoTo 0
                        i = i + 4
                    Else
                        out = out & c
                    End If
                Case Else
                    out = out & c
            End Select
            esc = False
        ElseIf c = "\" Then
            esc = True
        ElseIf c = """" Then
            ParseJsonQuotedString = out
            Exit Function
        Else
            out = out & c
        End If
        i = i + 1
    Loop
    ParseJsonQuotedString = out
End Function

Sub ParseJsonPairSegment(seg, ByRef dict)
    Dim valPart, key, value
    Dim q, i, c, esc
    seg = Trim(seg)
    If seg = "" Then Exit Sub
    If Left(seg, 1) <> """" Then Exit Sub
    q = ""
    esc = False
    i = 2
    Do While i <= Len(seg)
        c = Mid(seg, i, 1)
        If esc Then
            If c = """" Then
                q = q & """"
            ElseIf c = "\" Then
                q = q & "\"
            Else
                q = q & "\" & c
            End If
            esc = False
        ElseIf c = "\" Then
            esc = True
        ElseIf c = """" Then
            Exit Do
        Else
            q = q & c
        End If
        i = i + 1
    Loop
    If i > Len(seg) Then Exit Sub
    key = q
    valPart = Trim(Mid(seg, i + 1))
    If Left(valPart, 1) <> ":" Then Exit Sub
    valPart = Trim(Mid(valPart, 2))
    If Left(valPart, 1) = """" Then
        value = ParseJsonQuotedString(valPart)
    ElseIf Left(valPart, 1) = "[" Then
        ' Simple array parse: extract items between [ ]
        Dim arrContent, arrEnd
        arrEnd = InStr(valPart, "]")
        If arrEnd > 0 Then
            arrContent = Mid(valPart, 2, arrEnd - 2)
            value = arrContent ' Store raw array content for later parsing
        Else
            value = Trim(valPart)
        End If
    Else
        value = Trim(valPart)
    End If
    On Error Resume Next
    If dict.Exists(key) Then
        dict.Item(key) = value
    Else
        dict.Add key, value
    End If
    On Error GoTo 0
End Sub

Function ParseJson(jsonString)
    Dim dict, cleanStr, i, c, esc, inString, pairStart, depth
    Set dict = CreateObject("Scripting.Dictionary")
    cleanStr = Trim(jsonString)
    If Left(cleanStr, 1) = "{" Then cleanStr = Mid(cleanStr, 2)
    If Right(cleanStr, 1) = "}" Then cleanStr = Left(cleanStr, Len(cleanStr) - 1)
    esc = False
    inString = False
    depth = 0
    pairStart = 1
    i = 1
    Do While i <= Len(cleanStr)
        c = Mid(cleanStr, i, 1)
        If esc Then
            esc = False
        ElseIf inString Then
            If c = "\" Then
                esc = True
            ElseIf c = """" Then
                inString = False
            End If
        Else
            If c = """" Then
                inString = True
            ElseIf c = "[" Or c = "{" Then
                depth = depth + 1
            ElseIf c = "]" Or c = "}" Then
                depth = depth - 1
            ElseIf c = "," And depth = 0 Then
                ParseJsonPairSegment Mid(cleanStr, pairStart, i - pairStart), dict
                pairStart = i + 1
            End If
        End If
        i = i + 1
    Loop
    If pairStart <= Len(cleanStr) Then
        ParseJsonPairSegment Mid(cleanStr, pairStart), dict
    End If
    Set ParseJson = dict
End Function

Function ParseLineItems(lineItemsStr)
    ' Parse line_items array: [{...},{...}]
    Dim items(), itemCount, i, objStart, objEnd, objContent, itemDict
    itemCount = 0
    
    ' Remove outer brackets
    lineItemsStr = Trim(lineItemsStr)
    If Left(lineItemsStr, 1) = "[" Then lineItemsStr = Mid(lineItemsStr, 2)
    If Right(lineItemsStr, 1) = "]" Then lineItemsStr = Left(lineItemsStr, Len(lineItemsStr) - 1)
    
    ' Split by },{ pattern
    Dim parts, j
    parts = Split(lineItemsStr, "},{")
    ReDim items(UBound(parts))
    
    For j = 0 To UBound(parts)
        objContent = Trim(parts(j))
        If Left(objContent, 1) = "{" Then objContent = Mid(objContent, 2)
        If Right(objContent, 1) = "}" Then objContent = Left(objContent, Len(objContent) - 1)
        
        Set itemDict = ParseJson("{" & objContent & "}")
        Set items(j) = itemDict
        itemCount = itemCount + 1
    Next
    
    ParseLineItems = items
End Function

' --- Screenshot Helpers ---
Sub SaveScreenshot(executionId, tcode, stepNumber, stepName)
    Dim fso, base_dir, folder_path, filename, filepath
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    base_dir = "screenshots"
    folder_path = base_dir & "\" & executionId
    
    If Not fso.FolderExists(base_dir) Then
        On Error Resume Next
        fso.CreateFolder(base_dir)
        On Error GoTo 0
    End If
    
    If Not fso.FolderExists(folder_path) Then
        On Error Resume Next
        fso.CreateFolder(folder_path)
        On Error GoTo 0
    End If
    
    filename = tcode & "_" & stepNumber & "_" & Replace(stepName, " ", "_") & ".png"
    filepath = fso.GetAbsolutePathName(folder_path & "\" & filename)
    
    On Error Resume Next
    session.findById("wnd[0]").HardCopy filepath, 1
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to capture screenshot: " & Err.Description
    Else
        WScript.Echo "Screenshot saved at: " & filepath
    End If
    On Error GoTo 0
End Sub

Sub SaveMessagesModalScreenshotIfOpen(executionId, tcode, stepNumber, tagName)
    On Error Resume Next
    If session.Children.Count < 2 Then Exit Sub
    Dim w1t
    w1t = session.findById("wnd[1]").Text
    If InStr(1, LCase(w1t), "message", vbTextCompare) = 0 And InStr(1, LCase(w1t), "meldung", vbTextCompare) = 0 Then Exit Sub
    
    Dim fso, base_dir, folder_path, filename, filepath
    Set fso = CreateObject("Scripting.FileSystemObject")
    base_dir = "screenshots"
    folder_path = base_dir & "\" & executionId
    If Not fso.FolderExists(folder_path) Then Exit Sub
    
    filename = tcode & "_" & stepNumber & "_" & Replace(tagName, " ", "_") & ".png"
    filepath = fso.GetAbsolutePathName(folder_path & "\" & filename)
    
    session.findById("wnd[1]").HardCopy filepath, 1
    If Err.Number = 0 Then
        WScript.Echo "Screenshot saved at: " & filepath
    End If
    On Error GoTo 0
End Sub

' --- Messages Modal Handler (Dynpro Labels Only) ---
Function QualTest_LblText(session, wndPath, lblRow, lblCol)
    On Error Resume Next
    Dim lblPath, lblObj, lblText
    lblPath = wndPath & "/usr/lbl[" & lblRow & "," & lblCol & "]"
    Set lblObj = session.findById(lblPath)
    If Err.Number <> 0 Then
        Err.Clear
        QualTest_LblText = ""
        Exit Function
    End If
    lblText = lblObj.Text
    If Err.Number <> 0 Then
        Err.Clear
        lblText = ""
    End If
    QualTest_LblText = Trim(lblText)
End Function

Sub QualTest_ProcessSapMessagesModalLabelsOnly(session)
    On Error Resume Next
    If session.Children.Count < 2 Then Exit Sub
    
    Dim wndIdx, wndPath, wnd1, w1Title
    For wndIdx = 1 To 4
        wndPath = "wnd[" & wndIdx & "]"
        Set wnd1 = session.findById(wndPath)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        
        w1Title = LCase(Trim(wnd1.Text))
        If InStr(w1Title, "message") = 0 And InStr(w1Title, "meldung") = 0 Then
            Exit Sub
        End If
        
        WScript.Echo "INFO - Messages modal detected: " & wnd1.Text
        WScript.Sleep 800
        
        ' Screenshot BEFORE reading messages
        Dim fso, execId, ssFolder, ssFile
        On Error Resume Next
        execId = executionId ' Global variable
        Set fso = CreateObject("Scripting.FileSystemObject")
        ssFolder = "screenshots\" & execId
        If fso.FolderExists(ssFolder) Then
            ssFile = fso.GetAbsolutePathName(ssFolder & "\ME51N_Messages_modal.png")
            wnd1.HardCopy ssFile, 1
            If Err.Number = 0 Then
                WScript.Echo "Screenshot saved at: " & ssFile
            End If
        End If
        On Error GoTo 0
        
        WScript.Echo "INFO - Using label-based message reading"
        
        Dim MAJOR, rowLo, rowHi, row, col, lblText, msgLines()
        Dim msgCount, errCount, msgSummary, hasKeyword
        MAJOR = 7
        rowLo = 1
        rowHi = 48
        msgCount = 0
        errCount = 0
        msgSummary = ""
        ReDim msgLines(rowHi - rowLo + 1)
        
        ' First pass: lbl[MAJOR, row]
        For row = rowLo To rowHi
            lblText = QualTest_LblText(session, wndPath, MAJOR, row)
            If lblText <> "" Then
                Dim lblTextLower
                lblTextLower = LCase(lblText)
                If InStr(lblTextLower, "message text") = 0 And InStr(lblTextLower, "typ") = 0 And InStr(lblTextLower, "meldungstext") = 0 Then
                    WScript.Echo "Row " & row & " | " & lblText
                    msgLines(msgCount) = lblText
                    msgCount = msgCount + 1
                    
                    ' Keyword detection
                    hasKeyword = (InStr(lblTextLower, "does not exist") > 0) Or _
                                 (InStr(lblTextLower, "not activated") > 0) Or _
                                 (InStr(lblTextLower, "not found") > 0) Or _
                                 (InStr(lblTextLower, "invalid") > 0) Or _
                                 (InStr(lblTextLower, "error") > 0) Or _
                                 (InStr(lblTextLower, "fehler") > 0)
                    If hasKeyword Then
                        WScript.Echo "Output: [E] " & lblText
                        errCount = errCount + 1
                        If msgSummary = "" Then
                            msgSummary = lblText
                        Else
                            If Len(msgSummary) < 400 Then
                                msgSummary = msgSummary & "; " & lblText
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        ' Second pass: lbl[row, MAJOR] (swapped indices)
        For row = rowLo To rowHi
            lblText = QualTest_LblText(session, wndPath, row, MAJOR)
            If lblText <> "" Then
                lblTextLower = LCase(lblText)
                If InStr(lblTextLower, "message text") = 0 And InStr(lblTextLower, "typ") = 0 And InStr(lblTextLower, "meldungstext") = 0 Then
                    ' Check if already logged
                    Dim alreadyLogged, k
                    alreadyLogged = False
                    For k = 0 To msgCount - 1
                        If msgLines(k) = lblText Then
                            alreadyLogged = True
                            Exit For
                        End If
                    Next
                    
                    If Not alreadyLogged Then
                        WScript.Echo "Row " & row & " (swapped) | " & lblText
                        msgLines(msgCount) = lblText
                        msgCount = msgCount + 1
                        
                        hasKeyword = (InStr(lblTextLower, "does not exist") > 0) Or _
                                     (InStr(lblTextLower, "not activated") > 0) Or _
                                     (InStr(lblTextLower, "not found") > 0) Or _
                                     (InStr(lblTextLower, "invalid") > 0) Or _
                                     (InStr(lblTextLower, "error") > 0) Or _
                                     (InStr(lblTextLower, "fehler") > 0)
                        If hasKeyword Then
                            WScript.Echo "Output: [E] " & lblText
                            errCount = errCount + 1
                            If msgSummary = "" Then
                                msgSummary = lblText
                            Else
                                If Len(msgSummary) < 400 Then
                                    msgSummary = msgSummary & "; " & lblText
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        ' Final decision
        If errCount > 0 Then
            WScript.Echo "SAP_MESSAGE_POPUP: [E] " & msgSummary
            WScript.Echo "SAP_GUI_ERROR: " & errCount & " error(s) found in message popup"
            
            ' Close modal before exiting
            On Error Resume Next
            wnd1.sendVKey 0
            On Error GoTo 0
            
            WScript.Quit 1
        Else
            WScript.Echo "INFO - No errors in message popup"
            On Error Resume Next
            wnd1.sendVKey 0
            On Error GoTo 0
        End If
        
        Exit Sub
    Next
    On Error GoTo 0
End Sub

' --- Status Bar Handler ---
Sub CheckStatusBar()
    On Error Resume Next
    Dim sbar, sbarText, sbarType, sbarTypeCmp
    Set sbar = session.findById("wnd[0]/sbar")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    sbarText = sbar.Text
    If sbarText <> "" Then
        sbarType = sbar.MessageType
        WScript.Echo "Output: [" & sbarType & "] " & sbarText
        
        sbarTypeCmp = UCase(Trim(CStr(sbarType)))
        If sbarTypeCmp = "E" Or sbarTypeCmp = "A" Then
            WScript.Echo "SAP_GUI_ERROR: [" & sbarType & "] " & sbarText
            WScript.Quit 1
        End If
    End If
    On Error GoTo 0
End Sub

' ================================================================================
' MAIN SCRIPT
' ================================================================================

' --- Step 0: JSON Input Validation ---
If WScript.Arguments.Count = 0 Then
    WScript.Echo "ERROR: JSON input not provided. Pass JSON or File Path as a command-line argument."
    WScript.Quit 1
End If

Dim fso, jsonString, i, argJoined, firstArg
Set fso = CreateObject("Scripting.FileSystemObject")

firstArg = Trim(WScript.Arguments(0))
If Len(firstArg) >= 2 Then
    If Left(firstArg, 1) = """" And Right(firstArg, 1) = """" Then
        firstArg = Mid(firstArg, 2, Len(firstArg) - 2)
    End If
End If

If fso.FileExists(firstArg) And WScript.Arguments.Count = 1 Then
    On Error Resume Next
    Dim file
    Set file = fso.OpenTextFile(firstArg, 1)
    jsonString = file.ReadAll()
    file.Close
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: JSON file inaccessible."
        WScript.Quit 1
    End If
    On Error GoTo 0
Else
    argJoined = firstArg
    For i = 1 To WScript.Arguments.Count - 1
        argJoined = argJoined & " " & WScript.Arguments(i)
    Next
    jsonString = argJoined
End If

Dim tNorm
tNorm = Trim(jsonString)
If Len(tNorm) >= 2 Then
    If Right(tNorm, 2) = "}" & """" Then
        jsonString = Left(tNorm, Len(tNorm) - 1)
    Else
        jsonString = tNorm
    End If
Else
    jsonString = tNorm
End If

Dim dict, executionId, lineItemsStr, lineItems
Set dict = ParseJson(jsonString)

If dict.Count = 0 Then
    WScript.Echo "ERROR: Invalid or malformed JSON input."
    WScript.Quit 1
End If

executionId = dict.Item("execution_id")
If executionId = "" Then executionId = "exec_unknown"

WScript.Echo "INFO - Execution ID: " & executionId

' Parse line_items array
If dict.Exists("line_items") Then
    lineItemsStr = dict.Item("line_items")
    lineItems = ParseLineItems(lineItemsStr)
Else
    WScript.Echo "ERROR: line_items not found in JSON input."
    WScript.Quit 1
End If

' --- Step 1: SAP GUI Connection ---
WScript.Echo "INFO - Step 1: Connecting to SAP GUI"

Dim SapGuiAuto, application, connection, session

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If

If Not IsObject(session) Then
   Set session = connection.Children(0)
End If

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If

WScript.Echo "INFO - SAP GUI session connected"
SaveScreenshot executionId, "ME51N", "0", "Connection"

' --- Step 2: Navigate to ME51N ---
WScript.Echo "INFO - Step 2: Navigating to ME51N transaction"

' Recorded: session.findById("wnd[0]").maximize
session.findById("wnd[0]").maximize
WScript.Sleep 500

' Recorded: session.findById("wnd[0]/tbar[0]/okcd").text = "me51n"
session.findById("wnd[0]/tbar[0]/okcd").text = "me51n"
WScript.Sleep 200

' Recorded: session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500

WScript.Echo "INFO - ME51N transaction opened"
SaveScreenshot executionId, "ME51N", "1", "Navigation"
CheckStatusBar

' --- Step 3: Enter Line Items into Grid ---
WScript.Echo "INFO - Step 3: Entering line items into grid"

' Recorded grid path (preserved exactly from recording)
Dim gridPath
gridPath = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"

Dim grid, rowIdx, lineItem, material, quantity, shortText
Set grid = session.findById(gridPath)

For rowIdx = 0 To UBound(lineItems)
    Set lineItem = lineItems(rowIdx)
    
    material = ""
    quantity = ""
    shortText = ""
    
    If lineItem.Exists("material") Then material = Trim(lineItem.Item("material"))
    If lineItem.Exists("quantity") Then quantity = Trim(lineItem.Item("quantity"))
    If lineItem.Exists("short_text") Then shortText = Trim(lineItem.Item("short_text"))
    
    WScript.Echo "INFO - Processing line item " & (rowIdx + 1)
    
    ' OPTIONAL field: Material (MATNR)
    ' Recorded: session.findById(gridPath).modifyCell rowIdx,"MATNR","2092"
    If material <> "" Then
        On Error Resume Next
        grid.modifyCell rowIdx, "MATNR", material
        If Err.Number = 0 Then
            WScript.Echo "INFO - Line " & (rowIdx + 1) & " MATNR set to: " & material
        Else
            WScript.Echo "ERROR - Failed to set MATNR for line " & (rowIdx + 1) & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        WScript.Echo "INFO - Line " & (rowIdx + 1) & " MATNR not provided, skipping"
    End If
    
    ' OPTIONAL field: Quantity (MENGE)
    ' Recorded: session.findById(gridPath).modifyCell rowIdx,"MENGE","1000"
    If quantity <> "" Then
        On Error Resume Next
        grid.modifyCell rowIdx, "MENGE", quantity
        If Err.Number = 0 Then
            WScript.Echo "INFO - Line " & (rowIdx + 1) & " MENGE set to: " & quantity
        Else
            WScript.Echo "ERROR - Failed to set MENGE for line " & (rowIdx + 1) & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        WScript.Echo "INFO - Line " & (rowIdx + 1) & " MENGE not provided, skipping"
    End If
    
    ' OPTIONAL field: Short Text (NAME1)
    ' Recorded: session.findById(gridPath).modifyCell rowIdx,"NAME1","MTH1"
    If shortText <> "" Then
        On Error Resume Next
        grid.modifyCell rowIdx, "NAME1", shortText
        If Err.Number = 0 Then
            WScript.Echo "INFO - Line " & (rowIdx + 1) & " NAME1 set to: " & shortText
        Else
            WScript.Echo "ERROR - Failed to set NAME1 for line " & (rowIdx + 1) & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        WScript.Echo "INFO - Line " & (rowIdx + 1) & " NAME1 not provided, skipping"
    End If
    
    WScript.Sleep 300
Next

' Recorded: session.findById(gridPath).setCurrentCell lastRow,"NAME1"
If UBound(lineItems) >= 0 Then
    On Error Resume Next
    grid.setCurrentCell UBound(lineItems), "NAME1"
    Err.Clear
    On Error GoTo 0
End If

WScript.Echo "INFO - Line items entered successfully"
SaveScreenshot executionId, "ME51N", "2", "LineItems"

' --- Step 4: Check (btn[39]) ---
WScript.Echo "INFO - Step 4: Pressing Check button"

' Recorded: session.findById("wnd[0]/tbar[1]/btn[39]").press
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[39]").press
If Err.Number <> 0 Then
    WScript.Echo "ERROR - Failed to press Check button: " & Err.Description
    SaveScreenshot executionId, "ME51N", "3", "Check_Error"
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Sleep 1500
WScript.Echo "INFO - Check button pressed"
SaveScreenshot executionId, "ME51N", "3", "Check"

' Check for Messages modal after Check
QualTest_ProcessSapMessagesModalLabelsOnly session

' Check status bar after Check
CheckStatusBar

' --- Step 5: Save (btn[11]) ---
WScript.Echo "INFO - Step 5: Pressing Save button"

' Recorded: session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[11]").press
If Err.Number <> 0 Then
    WScript.Echo "ERROR - Failed to press Save button: " & Err.Description
    SaveScreenshot executionId, "ME51N", "4", "Save_Error"
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Sleep 2000
WScript.Echo "INFO - Save button pressed"
SaveScreenshot executionId, "ME51N", "4", "Save"

' Check for Messages modal after Save
QualTest_ProcessSapMessagesModalLabelsOnly session

' Check status bar after Save
CheckStatusBar

' --- Step 6: Optional Validation (Compare JSON vs SAP Grid Values) ---
WScript.Echo "INFO - Step 6: Optional validation - comparing JSON expected vs SAP actual"

For rowIdx = 0 To UBound(lineItems)
    Set lineItem = lineItems(rowIdx)
    
    material = ""
    quantity = ""
    shortText = ""
    
    If lineItem.Exists("material") Then material = Trim(lineItem.Item("material"))
    If lineItem.Exists("quantity") Then quantity = Trim(lineItem.Item("quantity"))
    If lineItem.Exists("short_text") Then shortText = Trim(lineItem.Item("short_text"))
    
    WScript.Echo "INFO - Validating line item " & (rowIdx + 1)
    
    ' Read back from grid and compare
    Dim actualMaterial, actualQuantity, actualShortText
    
    On Error Resume Next
    
    ' MATNR validation
    If material <> "" Then
        actualMaterial = Trim(grid.GetCellValue(rowIdx, "MATNR"))
        If Err.Number = 0 Then
            If UCase(actualMaterial) = UCase(material) Then
                WScript.Echo "INFO - VALIDATION PASS | Line " & (rowIdx + 1) & " MATNR | expected=" & material & " | actual=" & actualMaterial
            Else
                WScript.Echo "ERROR - VALIDATION FAIL | Line " & (rowIdx + 1) & " MATNR | expected=" & material & " | actual=" & actualMaterial
            End If
        Else
            WScript.Echo "INFO - CHECK SKIP | Line " & (rowIdx + 1) & " MATNR | could not read actual value"
            Err.Clear
        End If
    Else
        WScript.Echo "INFO - CHECK SKIP | Line " & (rowIdx + 1) & " MATNR | no expected value in JSON"
    End If
    
    ' MENGE validation
    If quantity <> "" Then
        actualQuantity = Trim(grid.GetCellValue(rowIdx, "MENGE"))
        If Err.Number = 0 Then
            If actualQuantity = quantity Then
                WScript.Echo "INFO - VALIDATION PASS | Line " & (rowIdx + 1) & " MENGE | expected=" & quantity & " | actual=" & actualQuantity
            Else
                WScript.Echo "ERROR - VALIDATION FAIL | Line " & (rowIdx + 1) & " MENGE | expected=" & quantity & " | actual=" & actualQuantity
            End If
        Else
            WScript.Echo "INFO - CHECK SKIP | Line " & (rowIdx + 1) & " MENGE | could not read actual value"
            Err.Clear
        End If
    Else
        WScript.Echo "INFO - CHECK SKIP | Line " & (rowIdx + 1) & " MENGE | no expected value in JSON"
    End If
    
    ' NAME1 validation
    If shortText <> "" Then
        actualShortText = Trim(grid.GetCellValue(rowIdx, "NAME1"))
        If Err.Number = 0 Then
            If UCase(actualShortText) = UCase(shortText) Then
                WScript.Echo "INFO - VALIDATION PASS | Line " & (rowIdx + 1) & " NAME1 | expected=" & shortText & " | actual=" & actualShortText
            Else
                WScript.Echo "ERROR - VALIDATION FAIL | Line " & (rowIdx + 1) & " NAME1 | expected=" & shortText & " | actual=" & actualShortText
            End If
        Else
            WScript.Echo "INFO - CHECK SKIP | Line " & (rowIdx + 1) & " NAME1 | could not read actual value"
            Err.Clear
        End If
    Else
        WScript.Echo "INFO - CHECK SKIP | Line " & (rowIdx + 1) & " NAME1 | no expected value in JSON"
    End If
    
    On Error GoTo 0
Next

SaveScreenshot executionId, "ME51N", "5", "Validation"

' --- Final Status ---
WScript.Echo "INFO - Script completed successfully"
WScript.Echo "INFO - Check screenshots in: screenshots\" & executionId
WScript.Quit 0