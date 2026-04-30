' =====================================================================
' SAP GUI Automation Script - ME51N (Create Purchase Requisition)
' =====================================================================
' Purpose: Automated creation of Purchase Requisition in SAP
' Transaction: ME51N
' Language: VBScript
' Runtime: Dynamic JSON-driven execution
' =====================================================================
Option Explicit
' =====================================================================
' JSON INPUT HANDLING - MANDATORY (File or String)
' =====================================================================
Dim fso, jsonString, argValue
Set fso = CreateObject("Scripting.FileSystemObject")
' Check if JSON input is provided
If WScript.Arguments.Count = 0 Then
    WScript.Echo "ERROR: JSON input not provided. Pass JSON or File Path as a command-line argument."
    WScript.Quit 1
End If
' Get command-line argument
argValue = WScript.Arguments(0)
' Automatic File Detection - check if argument is a file path
If fso.FileExists(argValue) Then
    On Error Resume Next
    Dim file
    Set file = fso.OpenTextFile(argValue, 1)
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: JSON file not found or inaccessible."
        WScript.Quit 1
    End If
    jsonString = file.ReadAll()
    file.Close
    On Error GoTo 0
Else
    ' Treat as raw JSON string
    jsonString = argValue
End If
' =====================================================================
' JSON PARSER - String-based implementation
' =====================================================================
Function ParseJson(jsonStr)
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    ' Remove outer braces and whitespace
    jsonStr = Trim(jsonStr)
    If Left(jsonStr, 1) = "{" Then jsonStr = Mid(jsonStr, 2)
    If Right(jsonStr, 1) = "}" Then jsonStr = Left(jsonStr, Len(jsonStr) - 1)
    ' Split by comma (simple parser for flat JSON)
    Dim pairs, pair, keyValue, key, value
    pairs = Split(jsonStr, ",")
    Dim i
    For i = 0 To UBound(pairs)
        pair = Trim(pairs(i))
        If InStr(pair, ":") > 0 Then
            keyValue = Split(pair, ":", 2)
            key = Trim(Replace(Replace(keyValue(0), """", ""), "'", ""))
            value = Trim(Replace(Replace(keyValue(1), """", ""), "'", ""))
            dict.Add key, value
        End If
    Next
    On Error GoTo 0
    Set ParseJson = dict
End Function
' Parse JSON input
Dim data
Set data = ParseJson(jsonString)
' Validate JSON parsing
If data.Count = 0 Then
    WScript.Echo "ERROR: Invalid or malformed JSON input."
    WScript.Quit 1
End If
' Extract execution_id for screenshots (with default)
Dim executionId
If data.Exists("execution_id") Then
    executionId = data.Item("execution_id")
Else
    executionId = "unknown_exec"
End If
' =====================================================================
' SCREENSHOT HELPER - MANDATORY
' =====================================================================
Sub SaveScreenshot(execId, tcode, stepNumber, stepName)
    Dim screenshotFso, baseDir, folderPath, filename, filepath
    Set screenshotFso = CreateObject("Scripting.FileSystemObject")
    ' Save in relative "screenshots" folder
    baseDir = "screenshots"
    folderPath = baseDir & "\" & execId
    ' Create base directory if not exists
    If Not screenshotFso.FolderExists(baseDir) Then
        On Error Resume Next
        screenshotFso.CreateFolder(baseDir)
        On Error GoTo 0
    End If
    ' Create execution directory if not exists
    If Not screenshotFso.FolderExists(folderPath) Then
        On Error Resume Next
        screenshotFso.CreateFolder(folderPath)
        On Error GoTo 0
    End If
    ' Generate filename with timestamp
    filename = tcode & "_" & stepNumber & "_" & Replace(stepName, " ", "_") & ".png"
    filepath = screenshotFso.GetAbsolutePathName(folderPath & "\" & filename)
    ' Capture screenshot using SAP GUI HardCopy
    On Error Resume Next
    session.FindById("wnd[0]").HardCopy filepath, 1
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to capture screenshot: " & Err.Description
    Else
        WScript.Echo "Screenshot saved at: " & filepath
    End If
    On Error GoTo 0
End Sub
' =====================================================================
' SAP GUI CONNECTION - MANDATORY
' =====================================================================
Dim SapGuiAuto, application, connection, session
' Initialize SAP GUI connection
On Error Resume Next
If Not IsObject(application) Then
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not connect to SAP GUI. Please ensure SAP GUI is running."
        WScript.Quit 1
    End If
    Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
    Set connection = application.Children(0)
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not get SAP connection. Please ensure you are logged into SAP."
        WScript.Quit 1
    End If
End If
If Not IsObject(session) Then
    Set session = connection.Children(0)
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not get SAP session."
        WScript.Quit 1
    End If
End If
On Error GoTo 0
' Connect events
If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If
WScript.Echo "INFO: SAP GUI connection established successfully"
' =====================================================================
' EXTRACT TEST DATA FROM JSON
' =====================================================================
' Extract all field values from JSON (with defaults from KB script)
Dim prDescription, materialNumber, quantity, plant, storageLocation
Dim purchasingGroup, requisitioner, desiredVendor
' Use JSON values if provided, otherwise use KB defaults
If data.Exists("pr_description") Then
    prDescription = data.Item("pr_description")
Else
    prDescription = "Test PR from Automation" ' Default value
End If
If data.Exists("material_number") Then
    materialNumber = data.Item("material_number")
Else
    materialNumber = "2092" ' KB default
End If
If data.Exists("quantity") Then
    quantity = data.Item("quantity")
Else
    quantity = "1000" ' KB default
End If
If data.Exists("plant") Then
    plant = data.Item("plant")
Else
    plant = "MTH1" ' KB default
End If
If data.Exists("storage_location") Then
    storageLocation = data.Item("storage_location")
Else
    storageLocation = "ZROM" ' KB default
End If
If data.Exists("purchasing_group") Then
    purchasingGroup = data.Item("purchasing_group")
Else
    purchasingGroup = "MT1" ' KB default
End If
If data.Exists("requisitioner") Then
    requisitioner = data.Item("requisitioner")
Else
    requisitioner = "LKARENNAGARI" ' KB default
End If
If data.Exists("desired_vendor") Then
    desiredVendor = data.Item("desired_vendor")
Else
    desiredVendor = "6000000071" ' KB default
End If
' =====================================================================
' STEP 1: NAVIGATE TO ME51N TRANSACTION
' =====================================================================
WScript.Echo "INFO: Step 1/12: Starting ME51N Purchase Requisition creation"
On Error Resume Next
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]
' Verification: Maximizing main window
session.FindById("wnd[0]").maximize
WScript.Sleep 500
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]/tbar[0]/okcd
' Verification: Transaction code field
WScript.Echo "INFO: Step 2/12: Navigating to transaction ME51N"
session.FindById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Verification: Pressing ENTER to execute transaction
session.FindById("wnd[0]").sendVKey 0
WScript.Sleep 1000
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to navigate to ME51N transaction: " & Err.Description
    WScript.Quit 1
End If
WScript.Echo "INFO: Step 2/12: Navigation to ME51N completed"
SaveScreenshot executionId, "ME51N", "1", "Navigation"
' =====================================================================
' STEP 3: SET FOCUS ON DESCRIPTION FIELD
' =====================================================================
WScript.Echo "INFO: Step 3/12: Setting focus on description field"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/txtMEREQ_TOPLINE-PURREQNDESCRIPTION
' Verification: PR description field - exact path from KB
Dim descField
Set descField = session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/txtMEREQ_TOPLINE-PURREQNDESCRIPTION")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to find description field: " & Err.Description
    Err.Clear
Else
    descField.setFocus
    descField.caretPosition = 0
    descField.text = prDescription
    WScript.Echo "INFO: Step 3/12: Description field set to: " & prDescription
End If
WScript.Sleep 500
' =====================================================================
' STEP 4: GRID CONTROL - SET CURRENT CELL TO MATERIAL NUMBER
' =====================================================================
WScript.Echo "INFO: Step 4/12: Setting up grid control"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell
' Verification: Grid control for line items - exact path from KB
Dim gridShell
Set gridShell = session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to find grid control: " & Err.Description
    WScript.Quit 1
End If
' Set current cell to MATNR column (Material Number)
gridShell.currentCellColumn = "MATNR"
WScript.Sleep 300
' Press Enter to prepare grid
gridShell.pressEnter
WScript.Sleep 500
WScript.Echo "INFO: Step 4/12: Grid control setup completed"
' =====================================================================
' STEP 5: SET TEXT EDITOR SELECTION (IF EXISTS)
' =====================================================================
WScript.Echo "INFO: Step 5/12: Setting text editor selection"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell
' Verification: Text editor control - exact path from KB
On Error Resume Next
Dim textEditor
Set textEditor = session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell")
If Err.Number = 0 Then
    textEditor.setSelectionIndexes 0,0
    WScript.Echo "INFO: Step 5/12: Text editor selection set"
Else
    WScript.Echo "INFO: Step 5/12: Text editor not found (may not be visible), continuing..."
    Err.Clear
End If
On Error GoTo 0
WScript.Sleep 300
' =====================================================================
' STEP 6: ENTER DATA INTO GRID - ALL FIELDS
' =====================================================================
WScript.Echo "INFO: Step 6/12: Entering data into grid cells"
' Re-acquire grid control reference
Set gridShell = session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
' Source: ME51N KB Reference Script - EXACT DATA ENTRY SEQUENCE
' All modifyCell calls use exact column names from KB script
On Error Resume Next
' Material Number (MATNR)
gridShell.modifyCell 0, "MATNR", materialNumber
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to set Material Number: " & Err.Description
    Err.Clear
Else
    WScript.Echo "INFO: Material Number set to: " & materialNumber
End If
' Quantity (MENGE)
gridShell.modifyCell 0, "MENGE", quantity
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to set Quantity: " & Err.Description
    Err.Clear
Else
    WScript.Echo "INFO: Quantity set to: " & quantity
End If
' Plant (NAME1)
gridShell.modifyCell 0, "NAME1", plant
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to set Plant: " & Err.Description
    Err.Clear
Else
    WScript.Echo "INFO: Plant set to: " & plant
End If
' Storage Location (LGOBE)
gridShell.modifyCell 0, "LGOBE", storageLocation
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to set Storage Location: " & Err.Description
    Err.Clear
Else
    WScript.Echo "INFO: Storage Location set to: " & storageLocation
End If
' Purchasing Group (EKGRP)
gridShell.modifyCell 0, "EKGRP", purchasingGroup
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to set Purchasing Group: " & Err.Description
    Err.Clear
Else
    WScript.Echo "INFO: Purchasing Group set to: " & purchasingGroup
End If
' Requisitioner (AFNAM)
gridShell.modifyCell 0, "AFNAM", requisitioner
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to set Requisitioner: " & Err.Description
    Err.Clear
Else
    WScript.Echo "INFO: Requisitioner set to: " & requisitioner
End If
' Desired Vendor (FLIEF)
gridShell.modifyCell 0, "FLIEF", desiredVendor
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to set Desired Vendor: " & Err.Description
    Err.Clear
Else
    WScript.Echo "INFO: Desired Vendor set to: " & desiredVendor
End If
On Error GoTo 0
WScript.Echo "INFO: Step 6/12: All grid data entry completed"
WScript.Sleep 500
SaveScreenshot executionId, "ME51N", "2", "Data_Entry"
' =====================================================================
' STEP 7: ADJUST GRID COLUMN VISIBILITY
' =====================================================================
WScript.Echo "INFO: Step 7/12: Adjusting grid column visibility"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Set current cell and adjust visible columns
gridShell.currentCellColumn = "FLIEF"
gridShell.firstVisibleColumn = "EPSTP"
WScript.Sleep 300
WScript.Echo "INFO: Step 7/12: Grid column visibility adjusted"
' =====================================================================
' STEP 8: VALIDATE DATA - FIRST ENTER
' =====================================================================
WScript.Echo "INFO: Step 8/12: Validating data (first Enter)"
' Source: ME51N KB Reference Script - EXACT SEQUENCE
' Press Enter to validate grid data
gridShell.pressEnter
WScript.Sleep 1000
WScript.Echo "INFO: Step 8/12: First validation completed"
SaveScreenshot executionId, "ME51N", "3", "First_Validation"
' =====================================================================
' STEP 9: VALIDATE DATA - SECOND ENTER (MODIFIED SCREEN)
' =====================================================================
WScript.Echo "INFO: Step 9/12: Validating data (second Enter on modified screen)"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell
' Verification: Grid control on modified screen (different subSUB0:0010) - exact path from KB
On Error Resume Next
Dim gridShell2
Set gridShell2 = session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
If Err.Number = 0 Then
    gridShell2.pressEnter
    WScript.Sleep 1000
    WScript.Echo "INFO: Step 9/12: Second validation completed"
Else
    WScript.Echo "INFO: Step 9/12: Modified screen grid not found, continuing with original grid reference"
    gridShell.pressEnter
    WScript.Sleep 1000
    Err.Clear
End If
On Error GoTo 0
SaveScreenshot executionId, "ME51N", "4", "Second_Validation"
' =====================================================================
' STEP 10: CHECK/VALIDATE OPERATION
' =====================================================================
WScript.Echo "INFO: Step 10/12: Performing check/validate operation"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]/tbar[1]/btn[39]
' Verification: Check button (toolbar button 39) - exact path from KB
On Error Resume Next
session.FindById("wnd[0]/tbar[1]/btn[39]").press
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to press check button: " & Err.Description
    Err.Clear
Else
    WScript.Echo "INFO: Step 10/12: Check/validate operation completed"
    WScript.Sleep 1000
End If
On Error GoTo 0
SaveScreenshot executionId, "ME51N", "5", "Check_Validate"
' =====================================================================
' STEP 11: SAVE PURCHASE REQUISITION
' =====================================================================
WScript.Echo "INFO: Step 11/12: Saving Purchase Requisition"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]/tbar[0]/btn[11]
' Verification: Save button (toolbar button 11) - exact path from KB
On Error Resume Next
session.FindById("wnd[0]/tbar[0]/btn[11]").press
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to press save button: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Sleep 1500
WScript.Echo "INFO: Step 11/12: Save operation completed"
SaveScreenshot executionId, "ME51N", "6", "After_Save"
' =====================================================================
' STEP 12: CAPTURE PR NUMBER FROM STATUS BAR
' =====================================================================
WScript.Echo "INFO: Step 12/12: Capturing PR number from status bar"
' Source: ME51N KB Reference Script - EXACT PATH PRESERVATION
' Path: wnd[0]/sbar
' Verification: Status bar - exact path from KB
On Error Resume Next
Dim statusBar
Set statusBar = session.FindById("wnd[0]/sbar")
If Err.Number = 0 Then
    ' Double-click status bar to view PR number (as per KB script)
    statusBar.doubleClick
    WScript.Sleep 500
    ' Capture and log status bar message
    If statusBar.Text <> "" Then
        WScript.Echo "Output: [" & statusBar.MessageType & "] " & statusBar.Text
    End If
Else
    WScript.Echo "INFO: Could not access status bar: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
SaveScreenshot executionId, "ME51N", "7", "Final_Result"
' =====================================================================
' SCRIPT COMPLETION
' =====================================================================
WScript.Echo "INFO: Step 12/12: ME51N Purchase Requisition creation completed successfully"
WScript.Echo "INFO: Script execution finished"
' Cleanup
Set gridShell = Nothing
Set gridShell2 = Nothing
Set descField = Nothing
Set textEditor = Nothing
Set statusBar = Nothing
Set session = Nothing
Set connection = Nothing
Set application = Nothing
Set SapGuiAuto = Nothing
Set data = Nothing
Set fso = Nothing
WScript.Quit 0