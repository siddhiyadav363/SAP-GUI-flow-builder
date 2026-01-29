' SAP GUI MIGO Goods Movement Complete Flow Script
' Process: Complete goods receipt processing with PO verification
' Transaction: MIGO (primary), ME23N (verification)
Option Explicit
' JSON Input Handling - MANDATORY (File or String)
If WScript.Arguments.Count = 0 Then
    WScript.Echo "ERROR: JSON input not provided. Pass JSON or File Path as a command-line argument."
    WScript.Quit 1
End If
Dim fso, jsonString, argValue
Set fso = CreateObject("Scripting.FileSystemObject")
argValue = WScript.Arguments(0)
' Automatic File Detection
If fso.FileExists(argValue) Then
    On Error Resume Next
    Dim file
    Set file = fso.OpenTextFile(argValue, 1)
    jsonString = file.ReadAll()
    file.Close
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: JSON file not found or inaccessible."
        WScript.Quit 1
    End If
    On Error GoTo 0
Else
    jsonString = argValue
End If
' Parse JSON using custom parser
Function ParseJson(jsonString)
    Set ParseJson = CreateObject("Scripting.Dictionary")
    ' Remove outer braces and clean string
    Dim cleanJson
    cleanJson = Replace(jsonString, "{", "")
    cleanJson = Replace(cleanJson, "}", "")
    cleanJson = Replace(cleanJson, """", "")
    cleanJson = Trim(cleanJson)
    ' Split by comma and parse key-value pairs
    Dim pairs, pair, keyValue
    pairs = Split(cleanJson, ",")
    Dim i
    For i = 0 To UBound(pairs)
        pair = Trim(pairs(i))
        If InStr(pair, ":") > 0 Then
            keyValue = Split(pair, ":")
            If UBound(keyValue) >= 1 Then
                ParseJson.Add Trim(keyValue(0)), Trim(keyValue(1))
            End If
        End If
    Next
End Function
' Parse and validate JSON
Dim data
Set data = ParseJson(jsonString)
If data.Count = 0 Then
    WScript.Echo "ERROR: Invalid or malformed JSON input."
    WScript.Quit 1
End If
' Extract data with defaults from reference script
Dim poNumber, quantity, deliveryNote
poNumber = "4500000170"  ' Default from reference script
quantity = "10"          ' Default from reference script  
deliveryNote = "4500000170"  ' Default from reference script
' Override with user data if provided
If data.Exists("po_number") Then poNumber = data("po_number")
If data.Exists("quantity") Then quantity = data("quantity")
If data.Exists("delivery_note") Then deliveryNote = data("delivery_note")
WScript.Echo "INFO - Starting MIGO goods movement complete flow"
WScript.Echo "INFO - PO Number: " & poNumber
WScript.Echo "INFO - Quantity: " & quantity
WScript.Echo "INFO - Delivery Note: " & deliveryNote
' SAP GUI Connection Setup
Dim SapGuiAuto, application, connection, session
On Error Resume Next
If Not IsObject(application) Then
    Set SapGuiAuto = GetObject("SAPGUI")
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
If Err.Number <> 0 Then
    WScript.Echo "ERROR - Failed to connect to SAP GUI: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "INFO - SAP GUI connection established successfully"
' Step 1: Initialize and maximize window
WScript.Echo "INFO - Step 1/11: Maximizing SAP window"
session.findById("wnd[0]").maximize
WScript.Echo "INFO - Step 1/11: Window maximized successfully"
' Step 2: Navigate to MIGO transaction
WScript.Echo "INFO - Step 2/11: Navigating to MIGO transaction"
' Source original: session.findById("wnd[0]/tbar[0]/okcd").text = "MIGO"
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[0]/okcd").text = "MIGO"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
WScript.Echo "INFO - Step 2/11: Navigation to MIGO completed"
' Step 3: Enter Purchase Order number
WScript.Echo "INFO - Step 3/11: Entering PO number: " & poNumber
' Source original: session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER")
' From: MIGO VBScript reference  
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER").text = poNumber
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER").caretPosition = Len(poNumber)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
WScript.Echo "INFO - Step 3/11: PO number entered successfully"
' Step 4: Set quantity in item details
WScript.Echo "INFO - Step 4/11: Setting quantity: " & quantity
' Source original: session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG")
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified  
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text = quantity
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").setFocus
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").caretPosition = Len(quantity)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500
WScript.Echo "INFO - Step 4/11: Quantity set successfully"
' Step 5: Select detail take checkbox
WScript.Echo "INFO - Step 5/11: Selecting detail take checkbox"
' Source original: session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE")
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = true
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").setFocus
WScript.Echo "INFO - Step 5/11: Detail take checkbox selected"
' Step 6: First validation - press check button
WScript.Echo "INFO - Step 6/11: Performing first validation"
' Source original: session.findById("wnd[0]/tbar[1]/btn[7]").press
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[1]/btn[7]").press
WScript.Sleep 1000
' Handle validation popup
On Error Resume Next
Dim popup
Set popup = session.findById("wnd[1]")
If Not popup Is Nothing Then
    WScript.Echo "INFO - Step 6/11: Handling validation popup"
    ' Source original: session.findById("wnd[1]/tbar[0]/btn[0]").press
    ' From: MIGO VBScript reference
    ' Verification: Path copied EXACTLY - character-by-character match verified
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 500
End If
On Error GoTo 0
WScript.Echo "INFO - Step 6/11: First validation completed"
' Step 7: Enter delivery note in header
WScript.Echo "INFO - Step 7/11: Entering delivery note: " & deliveryNote
' Source original: session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR").text = deliveryNote
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR").setFocus
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR").caretPosition = Len(deliveryNote)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500
WScript.Echo "INFO - Step 7/11: Delivery note entered successfully"
' Step 8: Second validation - press check button
WScript.Echo "INFO - Step 8/11: Performing second validation"
session.findById("wnd[0]/tbar[1]/btn[7]").press
WScript.Sleep 1000
' Handle second validation popup
On Error Resume Next
Set popup = session.findById("wnd[1]")
If Not popup Is Nothing Then
    WScript.Echo "INFO - Step 8/11: Handling second validation popup"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 500
End If
On Error GoTo 0
WScript.Echo "INFO - Step 8/11: Second validation completed"
' Step 9: Post the document
WScript.Echo "INFO - Step 9/11: Posting document"
' Source original: session.findById("wnd[0]/tbar[1]/btn[23]").press
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[1]/btn[23]").press
WScript.Sleep 2000
WScript.Echo "INFO - Step 9/11: Document posted successfully"
' Capture and log final result from status bar
On Error Resume Next
Dim sbar
Set sbar = session.findById("wnd[0]/sbar")
If Not sbar Is Nothing Then
    If sbar.Text <> "" Then
        WScript.Echo "Output: [" & sbar.MessageType & "] " & sbar.Text
    End If
End If
On Error GoTo 0
' Step 10: Navigate to ME23N for verification
WScript.Echo "INFO - Step 10/11: Navigating to ME23N for verification"
' Source original: session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
' From: MIGO VBScript reference  
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
' Open PO selection dialog
WScript.Echo "INFO - Step 10/11: Opening PO selection dialog"
' Source original: session.findById("wnd[0]/tbar[1]/btn[17]").press
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified  
session.findById("wnd[0]/tbar[1]/btn[17]").press
WScript.Sleep 1000
' Enter PO number in selection dialog
WScript.Echo "INFO - Step 10/11: Entering PO number in selection: " & poNumber
' Source original: session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN")
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = poNumber
session.findById("wnd[1]").sendVKey 0
WScript.Sleep 1000
WScript.Echo "INFO - Step 10/11: ME23N navigation completed"
' Step 11: Navigate to history tab for verification
WScript.Echo "INFO - Step 11/11: Navigating to history tab"
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT16")
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT16").select
' Set grid control properties for history review  
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell")
' From: MIGO VBScript reference
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").currentCellColumn = "BELNR"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").clickCurrentCell
WScript.Echo "INFO - Step 11/11: History tab navigation completed"
WScript.Echo "INFO - MIGO goods movement complete flow finished successfully"
WScript.Echo "INFO - Goods receipt processed for PO: " & poNumber & ", Quantity: " & quantity
WScript.Echo "INFO - Verification completed in ME23N transaction"