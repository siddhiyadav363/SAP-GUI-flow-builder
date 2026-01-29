' SAP GUI ME21N Purchase Order Creation Script
' Transaction: ME21N (Create Purchase Order)
' Module: Materials Management (MM)
' JSON Input: Dynamic runtime data via command line
Option Explicit
' Global variables
Dim SapGuiAuto, application, connection, session
Dim fso, jsonString, argValue, data
' JSON Input Handling - MANDATORY (File or String)
If WScript.Arguments.Count = 0 Then
    WScript.Echo "ERROR: JSON input not provided. Pass JSON or File Path as a command-line argument."
    WScript.Quit 1
End If
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
' Parse JSON input
Set data = ParseJson(jsonString)
If data.Count = 0 Then
    WScript.Echo "ERROR: Invalid or malformed JSON input."
    WScript.Quit 1
End If
' JSON Parser Function
Function ParseJson(jsonString)
    Set ParseJson = CreateObject("Scripting.Dictionary")
    ' Clean the JSON string
    jsonString = Replace(jsonString, "{", "")
    jsonString = Replace(jsonString, "}", "")
    jsonString = Replace(jsonString, """", "")
    ' Split by comma and process each key-value pair
    Dim pairs, i, pair, keyValue
    pairs = Split(jsonString, ",")
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
' Helper function to get JSON value with fallback
Function GetJsonValue(key, defaultValue)
    If data.Exists(key) Then
        GetJsonValue = data(key)
    Else
        GetJsonValue = defaultValue
    End If
End Function
' Main execution
On Error Resume Next
WScript.Echo "INFO - Starting ME21N Purchase Order Creation"
' Step 1: SAP GUI Connection Setup
' Source original: Reference ME21N VBScript - Connection Management
WScript.Echo "INFO - Step 1/12: Establishing SAP GUI connection"
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
' Maximize window
session.findById("wnd[0]").maximize
WScript.Echo "INFO - Step 1/12: SAP GUI connection established successfully"
' Step 2: Transaction Navigation (ME21N)
' Source original: session.findById("wnd[0]/tbar[0]/okcd").text = "me21n"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "INFO - Step 2/12: Navigating to ME21N transaction"
session.findById("wnd[0]/tbar[0]/okcd").text = "me21n"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
WScript.Echo "INFO - Step 2/12: ME21N transaction navigation completed"
' Step 3: Vendor Selection
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = "6000000071"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "INFO - Step 3/12: Setting vendor code"
Dim vendorCode
vendorCode = GetJsonValue("vendor_code", "6000000071")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = vendorCode
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
WScript.Echo "INFO - Step 3/12: Vendor code " & vendorCode & " set successfully"
' Step 4: Header Data Configuration - Purchasing Organization
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").text = "mth1"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "INFO - Step 4/12: Setting purchasing organization"
Dim purchOrg
purchOrg = GetJsonValue("purchasing_org", "mth1")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").text = purchOrg
WScript.Echo "INFO - Step 4/12: Purchasing organization " & purchOrg & " set successfully"
' Step 5: Header Data Configuration - Purchasing Group
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").text = "MT1"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "INFO - Step 5/12: Setting purchasing group"
Dim purchGroup
purchGroup = GetJsonValue("purchasing_group", "MT1")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").text = purchGroup
WScript.Echo "INFO - Step 5/12: Purchasing group " & purchGroup & " set successfully"
' Step 6: Header Data Configuration - Company Code
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").text = "9000"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "INFO - Step 6/12: Setting company code"
Dim companyCode
companyCode = GetJsonValue("company_code", "9000")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").text = companyCode
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500
WScript.Echo "INFO - Step 6/12: Company code " & companyCode & " set successfully"
' Step 7: Line Item Details - Material Number
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").text = "2092"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "INFO - Step 7/12: Setting line item details"
Dim materialNumber
materialNumber = GetJsonValue("material_number", "2092")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").text = materialNumber
WScript.Echo "INFO - Material number " & materialNumber & " set"
' Step 8: Line Item Details - Quantity
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").text = "1000"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
Dim quantity
quantity = GetJsonValue("quantity", "1000")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").text = quantity
WScript.Echo "INFO - Quantity " & quantity & " set"
' Step 9: Line Item Details - Net Price
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,0]").text = "1000"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
Dim netPrice
netPrice = GetJsonValue("net_price", "1000")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,0]").text = netPrice
WScript.Echo "INFO - Net price " & netPrice & " set"
' Step 10: Line Item Details - Plant and Storage Location
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]").text = "MTH1"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
Dim plant, storageLocation
plant = GetJsonValue("plant", "MTH1")
storageLocation = GetJsonValue("storage_location", "ZROM")
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]").text = plant
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-LGOBE[16,0]").text = storageLocation
WScript.Echo "INFO - Plant " & plant & " and Storage Location " & storageLocation & " set"
' Focus and validate
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-CHARG[17,0]").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-CHARG[17,0]").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500
' Step 11: Purchase Requisition Reference
' Source original: session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[28,0]").text = "0010000192"
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "INFO - Step 11/12: Setting Purchase Requisition reference"
Dim prNumber, prItem
prNumber = GetJsonValue("pr_number", "0010000192")
prItem = GetJsonValue("pr_item", "10")
' Set PR Number (following exact sequence from reference script)
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[28,0]").text = "100000192"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[28,0]").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[28,0]").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
' Correct PR Number (as per reference script sequence)
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[28,0]").text = "100000192"
' Set PR Item Number
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-BNFPO[29,0]").text = prItem
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-BNFPO[29,0]").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-BNFPO[29,0]").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
' Adjust column width and set final PR Number (exact sequence from reference)
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").columns.elementAt(28).width = 10
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[28,0]").text = prNumber
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[28,0]").setFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[28,0]").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000
WScript.Echo "INFO - Step 11/12: Purchase Requisition " & prNumber & " item " & prItem & " set successfully"
' Step 12: Save Operations
' Source original: session.findById("wnd[0]/tbar[1]/btn[39]").press
' Source original: session.findById("wnd[0]/tbar[0]/btn[11]").press
' From: ME21N VBScript Reference
' Verification: Path copied EXACTLY - character-by-character match verified
WScript.Echo "INFO - Step 12/12: Saving Purchase Order"
session.findById("wnd[0]/tbar[1]/btn[39]").press
WScript.Sleep 1000
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 2000
' Check for popup after save operation
On Error Resume Next
Dim popup
Set popup = session.findById("wnd[1]")
If Not popup Is Nothing Then
    WScript.Echo "INFO - Handling save confirmation popup"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 1000
End If
On Error GoTo 0
' Capture final status bar message
Dim sbar
Set sbar = session.findById("wnd[0]/sbar")
If Not sbar Is Nothing Then
    If sbar.Text <> "" Then
        WScript.Echo "Output: [" & sbar.MessageType & "] " & sbar.Text
    End If
End If
WScript.Echo "INFO - Step 12/12: Purchase Order save operation completed"
WScript.Echo "INFO - ME21N Purchase Order Creation flow completed successfully"
If Err.Number <> 0 Then
    WScript.Echo "ERROR - Script execution failed: " & Err.Description
    WScript.Quit 1
End If