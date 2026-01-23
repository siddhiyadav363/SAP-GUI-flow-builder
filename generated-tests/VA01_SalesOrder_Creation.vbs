' VA01 Sales Order Creation - Complete Flow
' Dynamic JSON-driven execution with automatic file detection
' Usage: cscript VA01_SalesOrder_Creation.vbs "C:\path\to\data.json"
' Usage: cscript VA01_SalesOrder_Creation.vbs "{""order_type"":""OR"",""customer"":""900005""}"
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
' Custom JSON Parser (String-based)
Function ParseJson(jsonString)
    Set dict = CreateObject("Scripting.Dictionary")
    ' Clean string (remove braces and spaces)
    Dim cleanStr
    cleanStr = Replace(jsonString, "{", "")
    cleanStr = Replace(cleanStr, "}", "")
    cleanStr = Replace(cleanStr, """", "")
    cleanStr = Trim(cleanStr)
    ' Split by comma and parse key:value pairs
    If Len(cleanStr) > 0 Then
        Dim pairs, i, pair, keyValue
        pairs = Split(cleanStr, ",")
        For i = 0 To UBound(pairs)
            pair = Trim(pairs(i))
            If InStr(pair, ":") > 0 Then
                keyValue = Split(pair, ":")
                If UBound(keyValue) >= 1 Then
                    dict.Add Trim(keyValue(0)), Trim(keyValue(1))
                End If
            End If
        Next
    End If
    Set ParseJson = dict
End Function
' Parse JSON input
Dim data
Set data = ParseJson(jsonString)
' Validate JSON parsing
If data.Count = 0 And Len(jsonString) > 10 Then
    WScript.Echo "ERROR: Invalid or malformed JSON input."
    WScript.Quit 1
End If
' Extract values from JSON with defaults from knowledge base
Dim orderType, customer, poReference, material, quantity, conditionType, conditionAmount
orderType = "OR"
customer = "900005"
poReference = "12345"
material = "30010"
quantity = "10"
conditionType = "PR00"
conditionAmount = "500"
' Override with JSON values if provided
If data.Exists("order_type") Then orderType = data("order_type")
If data.Exists("customer") Then customer = data("customer")
If data.Exists("po_reference") Then poReference = data("po_reference")
If data.Exists("material") Then material = data("material")
If data.Exists("quantity") Then quantity = data("quantity")
If data.Exists("condition_type") Then conditionType = data("condition_type")
If data.Exists("condition_amount") Then conditionAmount = data("condition_amount")
' Logging function
Sub LogInfo(message)
    WScript.Echo FormatDateTime(Now, 0) & " - INFO - " & message
End Sub
Sub LogError(message)
    WScript.Echo FormatDateTime(Now, 0) & " - ERROR - " & message
End Sub
' Error handling
On Error Resume Next
' SAP Connection Management - EXACT from reference script
LogInfo "Starting VA01 Sales Order Creation flow"
LogInfo "Step 1/11: Establishing SAP connection"
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
If Err.Number <> 0 Then
    LogError "Failed to establish SAP connection: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 1/11: SAP connection established successfully"
' Step 2: Navigate to VA01 transaction - EXACT from reference script
LogInfo "Step 2/11: Navigating to VA01 transaction"
' Source original: session.findById("wnd[0]").maximize
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").maximize
' Source original: session.findById("wnd[0]/tbar[0]/okcd").text = "VA01"
' From: VA01 Creation script  
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[0]/okcd").text = "VA01"
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then
    LogError "Failed to navigate to VA01: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 2/11: Navigation to VA01 completed successfully"
' Step 3: Set order type - EXACT from reference script
LogInfo "Step 3/11: Setting order type to " & orderType
' Source original: session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = "OR"
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = orderType
' Source original: session.findById("wnd[0]/usr/ctxtVBAK-AUART").caretPosition = 2
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ctxtVBAK-AUART").caretPosition = 2
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then
    LogError "Failed to set order type: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 3/11: Order type set successfully"
' Step 4: Enter customer - EXACT from reference script
LogInfo "Step 4/11: Entering customer " & customer
' Source original: session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = "900005"
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = customer
' Source original: session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").caretPosition = 6
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").caretPosition = 6
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then
    LogError "Failed to enter customer: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 4/11: Customer entered successfully"
' Step 5: Enter PO reference - EXACT from reference script
LogInfo "Step 5/11: Entering PO reference " & poReference
' Source original: session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = "12345"
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = poReference
' Source original: session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 5
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 5
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then
    LogError "Failed to enter PO reference: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 5/11: PO reference entered successfully"
' Step 6: F4 help selection - EXACT from reference script
LogInfo "Step 6/11: Performing F4 help selection"
' Source original: session.findById("wnd[0]").sendVKey 4
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 4
' Wait for F4 popup to appear
WScript.Sleep 1000
' Source original: session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 4
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 4
' Source original: session.findById("wnd[1]").sendVKey 2
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[1]").sendVKey 2
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then
    LogError "Failed to perform F4 selection: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 6/11: F4 help selection completed successfully"
' Step 7: Enter material - EXACT from reference script
LogInfo "Step 7/11: Entering material " & material
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").text = "30010"
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").text = material
If Err.Number <> 0 Then
    LogError "Failed to enter material: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 7/11: Material entered successfully"
' Step 8: Enter quantity - EXACT from reference script
LogInfo "Step 8/11: Entering quantity " & quantity
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[3,0]").text = "10"
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[3,0]").text = quantity
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[3,0]").setFocus
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[3,0]").setFocus
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[3,0]").caretPosition = 19
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[3,0]").caretPosition = 19
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then
    LogError "Failed to enter quantity: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 8/11: Quantity entered successfully"
' Additional material processing - EXACT from reference script
LogInfo "Step 8a/11: Additional material field processing"
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").setFocus
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").setFocus
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").caretPosition = 5
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").caretPosition = 5
' Source original: session.findById("wnd[0]").sendVKey 2
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 2
If Err.Number <> 0 Then
    LogError "Failed in additional material processing: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 8a/11: Additional material processing completed successfully"
' Step 9: Navigate to conditions tab - EXACT from reference script
LogInfo "Step 9/11: Navigating to conditions tab"
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 2
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 2
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 4
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 4
If Err.Number <> 0 Then
    LogError "Failed to navigate to conditions tab: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 9/11: Conditions tab navigation completed successfully"
' Step 10: Set pricing conditions - EXACT from reference script
LogInfo "Step 10/11: Setting pricing conditions - Type: " & conditionType & ", Amount: " & conditionAmount
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,6]").text = "PR00"
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,6]").text = conditionType
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,6]").text = "500"
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,6]").text = conditionAmount
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,6]").setFocus
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,6]").setFocus
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,6]").caretPosition = 16
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,6]").caretPosition = 16
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then
    LogError "Failed to set pricing conditions: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 10/11: Pricing conditions set successfully"
' Step 11: Save operation with popup handling - EXACT from reference script
LogInfo "Step 11/11: Saving sales order"
' Source original: session.findById("wnd[0]/tbar[0]/btn[11]").press
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/tbar[0]/btn[11]").press
' Wait for popup to appear
WScript.Sleep 1000
' Source original: session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
' From: VA01 Creation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
If Err.Number <> 0 Then
    LogError "Failed to save sales order: " & Err.Description
    WScript.Quit 1
End If
LogInfo "Step 11/11: Sales order saved successfully"
LogInfo "VA01 Sales Order Creation flow completed successfully"
' Final success message
WScript.Echo "SUCCESS: VA01 Sales Order Creation completed successfully"
WScript.Echo "Order Type: " & orderType
WScript.Echo "Customer: " & customer
WScript.Echo "Material: " & material
WScript.Echo "Quantity: " & quantity
WScript.Echo "Condition: " & conditionType & " = " & conditionAmount
WScript.Quit 0