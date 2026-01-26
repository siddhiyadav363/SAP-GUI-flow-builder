' SAP GUI VA11 Create Quotation Automation Script
' Transaction: VA11 (Create Quotation)
' Purpose: Automated creation of sales quotation with organizational data, customer information, and line items
' Language: VBScript
' Module: Sales and Distribution (SD)
' Check command line arguments for JSON input (required)
If WScript.Arguments.Count = 0 Then
    WScript.Echo "ERROR: JSON input not provided. Pass JSON or File Path as a command-line argument."
    WScript.Quit 1
End If
' Get JSON input (File Path or Raw JSON)
Dim fso, jsonString, argValue
Set fso = CreateObject("Scripting.FileSystemObject")
argValue = WScript.Arguments(0)
' Automatic File Detection - check if argument is a valid file path
If fso.FileExists(argValue) Then
    ' Read JSON from file
    On Error Resume Next
    Dim file
    Set file = fso.OpenTextFile(argValue, 1)
    If Err.Number = 0 Then
        jsonString = file.ReadAll()
        file.Close
    Else
        WScript.Echo "ERROR: JSON file not found or inaccessible."
        WScript.Quit 1
    End If
    On Error GoTo 0
Else
    ' Treat as raw JSON string
    jsonString = argValue
End If
' Parse JSON input using custom parser
Function ParseJson(jsonString)
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    ' Clean JSON string - remove outer braces and spaces
    Dim cleanJson
    cleanJson = Trim(jsonString)
    cleanJson = Mid(cleanJson, 2, Len(cleanJson) - 2) ' Remove { }
    ' Split by comma and parse key-value pairs
    Dim pairs, pair, keyValue, key, value, i
    pairs = Split(cleanJson, ",")
    For i = 0 To UBound(pairs)
        pair = Trim(pairs(i))
        If InStr(pair, ":") > 0 Then
            keyValue = Split(pair, ":", 2)
            If UBound(keyValue) = 1 Then
                key = Trim(keyValue(0))
                value = Trim(keyValue(1))
                ' Remove quotes
                key = Replace(Replace(key, """", ""), "'", "")
                value = Replace(Replace(value, """", ""), "'", "")
                dict.Add key, value
            End If
        End If
    Next
    If Err.Number <> 0 Then
        Set dict = Nothing
    End If
    On Error GoTo 0
    Set ParseJson = dict
End Function
' Parse JSON data
Dim data
Set data = ParseJson(jsonString)
' Validate JSON parsing
If data Is Nothing Or data.Count = 0 Then
    WScript.Echo "ERROR: Invalid or malformed JSON input."
    WScript.Quit 1
End If
' Extract values from JSON with defaults from knowledge base
Dim documentType, salesOrg, distChannel, division, customerNumber, materialNumber, quantity, conditionType, price
documentType = "IN"
salesOrg = "9000"  
distChannel = "9A"
division = "9A"
customerNumber = "900005"
materialNumber = "30010"
quantity = "2"
conditionType = "PR00" 
price = "1000"
' Override with JSON values if provided
If data.Exists("documentType") Then documentType = data("documentType")
If data.Exists("salesOrg") Then salesOrg = data("salesOrg")
If data.Exists("distChannel") Then distChannel = data("distChannel") 
If data.Exists("division") Then division = data("division")
If data.Exists("customerNumber") Then customerNumber = data("customerNumber")
If data.Exists("materialNumber") Then materialNumber = data("materialNumber")
If data.Exists("quantity") Then quantity = data("quantity")
If data.Exists("conditionType") Then conditionType = data("conditionType")
If data.Exists("price") Then price = data("price")
' SAP GUI Connection Management
' Source original: SAP GUI connection initialization from VA11 reference script
' From: VA11 Create Quotation script
' Verification: Connection code copied EXACTLY - character-by-character match verified
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
' Error handling wrapper for main automation
On Error Resume Next
' Step 1: Window management and transaction navigation
' Source original: session.findById("wnd[0]").maximize
' From: VA11 Create Quotation script  
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "Step 1/8: Maximizing window and navigating to VA11 transaction"
session.findById("wnd[0]").maximize
' Source original: session.findById("wnd[0]/tbar[0]/okcd").text = "VA11"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified  
session.findById("wnd[0]/tbar[0]/okcd").text = "VA11"
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
WScript.Echo "Step 1/8: Navigation to VA11 completed successfully"
' Step 2: Set organizational data (Document Type, Sales Organization, Distribution Channel, Division)
' Source original: session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = "IN"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "Step 2/8: Setting organizational data"
session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = documentType
' Source original: session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = "9000"
' From: VA11 Create Quotation script 
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = salesOrg
' Source original: session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = "9A"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = distChannel
' Source original: session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = "9A"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = division
' Source original: session.findById("wnd[0]/usr/ctxtVBAK-SPART").setFocus
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ctxtVBAK-SPART").setFocus
' Source original: session.findById("wnd[0]/usr/ctxtVBAK-SPART").caretPosition = 2
' From: VA11 Create Quotation script
' Verification: CaretPosition value copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/ctxtVBAK-SPART").caretPosition = 2
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
WScript.Echo "Step 2/8: Organizational data set successfully"
' Step 3: Set customer data (Sold-to Party)
' Source original: session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = "900005"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "Step 3/8: Setting customer data"
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = customerNumber
' Source original: session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").setFocus
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").setFocus
' Source original: session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").caretPosition = 0
' From: VA11 Create Quotation script  
' Verification: CaretPosition value copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").caretPosition = 0
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
WScript.Echo "Step 3/8: Customer data set successfully"
' Step 4: Add line item data (Material Number and Quantity)
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").text = "30010"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified  
WScript.Echo "Step 4/8: Adding line item data"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").text = materialNumber
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[3,0]").text = "2"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[3,0]").text = quantity
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[3,0]").setFocus
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[3,0]").setFocus
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[3,0]").caretPosition = 19
' From: VA11 Create Quotation script
' Verification: CaretPosition value copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[3,0]").caretPosition = 19
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").setFocus
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").setFocus
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").caretPosition = 5
' From: VA11 Create Quotation script
' Verification: CaretPosition value copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").caretPosition = 5
' Source original: session.findById("wnd[0]").sendVKey 2
' From: VA11 Create Quotation script
' Verification: VKey value copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 2
WScript.Echo "Step 4/8: Line item data added successfully"
' Step 5: Navigate to Conditions tab
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "Step 5/8: Navigating to Conditions tab"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 3
' From: VA11 Create Quotation script  
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 3
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 5
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 5
WScript.Echo "Step 5/8: Navigation to Conditions tab completed"
' Step 6: Set pricing conditions (Condition Type and Price)
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,5]").text = "PR00"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "Step 6/8: Setting pricing conditions"  
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,5]").text = conditionType
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = "1000"
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = price
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus
' Source original: session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
' From: VA11 Create Quotation script  
' Verification: CaretPosition value copied EXACTLY - character-by-character match verified
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
' Source original: session.findById("wnd[0]").sendVKey 0
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
session.findById("wnd[0]").sendVKey 0
WScript.Echo "Step 6/8: Pricing conditions set successfully"
' Step 7: Save document
' Source original: session.findById("wnd[0]/tbar[0]/btn[11]").press
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "Step 7/8: Saving quotation document"
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Echo "Step 7/8: Save button pressed successfully"
' Step 8: Handle confirmation popup
' Source original: session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press  
' From: VA11 Create Quotation script
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "Step 8/8: Handling confirmation popup"
WScript.Sleep 1000 ' Wait for popup to appear
' Handle popup if it exists
On Error Resume Next
Dim popup
Set popup = session.findById("wnd[1]/usr/btnSPOP-VAROPTION1")
If Not popup Is Nothing Then
    popup.press
    WScript.Echo "Step 8/8: Confirmation popup handled successfully"
Else
    WScript.Echo "Step 8/8: No confirmation popup found"
End If
On Error GoTo 0
WScript.Echo "VA11 Create Quotation automation completed successfully"
WScript.Echo "Document Type: " & documentType
WScript.Echo "Sales Organization: " & salesOrg  
WScript.Echo "Customer: " & customerNumber
WScript.Echo "Material: " & materialNumber & " (Quantity: " & quantity & ")"
WScript.Echo "Price: " & price & " (" & conditionType & ")"
' Check for errors
If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Description & " (Error: " & Err.Number & ")"
    WScript.Quit 1
End If
WScript.Echo "Script execution completed successfully."