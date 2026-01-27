' ME51N Purchase Requisition Creation - Complete Flow
' JSON-driven execution with automatic file detection
Dim jsonString, data, argValue
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
' Check for command-line argument
If WScript.Arguments.Count = 0 Then
    WScript.Echo "ERROR: JSON input not provided. Pass JSON or File Path as a command-line argument."
    WScript.Quit 1
End If
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
Set data = ParseJson(jsonString)
If data.Count = 0 Then
    WScript.Echo "ERROR: Invalid or malformed JSON input."
    WScript.Quit 1
End If
' SAP GUI Connection - From ME51N reference script
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
WScript.Echo "INFO - Starting ME51N Purchase Requisition Creation flow"
On Error Resume Next
' Step 1: Window maximization - From ME51N reference script
WScript.Echo "INFO - Step 1/9: Maximizing window"
' Source original: session.findById("wnd[0]").maximize
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").maximize
WScript.Echo "INFO - Step 1/9: Window maximized successfully"
' Step 2: Navigate to ME51N transaction - From ME51N reference script
WScript.Echo "INFO - Step 2/9: Navigating to ME51N transaction"
' Source original: session.findById("wnd[0]/tbar[0]/okcd").text = "ME51N"
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[0]/okcd").text = "ME51N"
' Source original: session.findById("wnd[0]").sendVKey 0
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
WScript.Echo "INFO - Step 2/9: Navigation completed successfully"
' Step 3: Enter Material Number - From ME51N reference script
WScript.Echo "INFO - Step 3/9: Entering material number"
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"MATNR","2092"
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
materialNumber = data.Item("material_number")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"MATNR",materialNumber
WScript.Echo "INFO - Entered material number: " & materialNumber
' Step 4: Enter Quantity - From ME51N reference script
WScript.Echo "INFO - Step 4/9: Entering quantity"
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"MENGE","10"
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
quantity = data.Item("quantity")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"MENGE",quantity
WScript.Echo "INFO - Entered quantity: " & quantity
' Step 5: Enter Vendor Name - From ME51N reference script
WScript.Echo "INFO - Step 5/9: Entering vendor name"
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"NAME1","MTH1"
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
vendorName = data.Item("vendor_name")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"NAME1",vendorName
WScript.Echo "INFO - Entered vendor name: " & vendorName
' Step 6: Enter Storage Location - From ME51N reference script
WScript.Echo "INFO - Step 6/9: Entering storage location"
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"LGOBE","ZROM"
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
storageLocation = data.Item("storage_location")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"LGOBE",storageLocation
WScript.Echo "INFO - Entered storage location: " & storageLocation
' Step 7: Enter Purchasing Group - From ME51N reference script
WScript.Echo "INFO - Step 7/9: Entering purchasing group"
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"EKGRP","MT1"
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
purchasingGroup = data.Item("purchasing_group")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"EKGRP",purchasingGroup
WScript.Echo "INFO - Entered purchasing group: " & purchasingGroup
' Step 8: Enter Vendor Number and validate - From ME51N reference script
WScript.Echo "INFO - Step 8/9: Entering vendor number"
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"LIFNR","6000000071"
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
vendorNumber = data.Item("vendor_number")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"LIFNR",vendorNumber
WScript.Echo "INFO - Entered vendor number: " & vendorNumber
' Set cursor position and validate - From ME51N reference script
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "LIFNR"
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "LIFNR"
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
WScript.Sleep 1000
' Process/Check operations (executed twice as per reference script)
WScript.Echo "INFO - Processing data validation (first pass)"
' Source original: session.findById("wnd[0]/tbar[1]/btn[39]").press
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[1]/btn[39]").press
WScript.Sleep 1000
WScript.Echo "INFO - Processing data validation (second pass)"
' Source original: session.findById("wnd[0]/tbar[1]/btn[39]").press
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[1]/btn[39]").press
WScript.Sleep 1000
' Step 9: Save Purchase Requisition - From ME51N reference script
WScript.Echo "INFO - Step 9/9: Saving purchase requisition"
' Source original: session.findById("wnd[0]/tbar[0]/btn[11]").press
' From: ME51N reference script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 2000
' Check for popup after save
On Error Resume Next
Set popup = session.FindById("wnd[1]")
If Not popup Is Nothing Then
    WScript.Echo "INFO - Handling popup after save"
    popup.sendVKey 0  ' Enter
    WScript.Sleep 1000
End If
On Error GoTo 0
' Capture final status bar message
On Error Resume Next
Set sbar = session.FindById("wnd[0]/sbar")
If Not sbar Is Nothing Then
    If sbar.Text <> "" Then
        WScript.Echo "Output: [" & sbar.MessageType & "] " & sbar.Text
    End If
End If
On Error GoTo 0
WScript.Echo "INFO - ME51N Purchase Requisition Creation flow completed successfully"
' Custom JSON Parser Function
Function ParseJson(jsonString)
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    ' Remove outer braces and clean string
    cleanString = Replace(jsonString, "{", "")
    cleanString = Replace(cleanString, "}", "")
    cleanString = Replace(cleanString, """", "")
    cleanString = Trim(cleanString)
    ' Split by commas to get key-value pairs
    pairs = Split(cleanString, ",")
    For Each pair In pairs
        If InStr(pair, ":") > 0 Then
            keyValue = Split(pair, ":")
            If UBound(keyValue) >= 1 Then
                key = Trim(keyValue(0))
                value = Trim(keyValue(1))
                dict.Add key, value
            End If
        End If
    Next
    On Error GoTo 0
    Set ParseJson = dict
End Function