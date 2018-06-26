WScript.Sleep(10000)

Set objShell = WScript.CreateObject("WScript.Shell")

objShell.Run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\Saplogon.exe"""
WScript.Sleep(10000)
If Not IsObject(application) Then
	Set SapGuiAuto = GetObject("SAPGUI")
	Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
	Set connection = application.OpenConnection("System name", True)
End If
If Not IsObject(session) Then
	Set session = connection.Children(0)
End If
WScript.Sleep(10000)

' session.findById("wnd[0]").maximize
' session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "Mandant ID"
' session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "Username"
' session.findById("wnd[0]/usr/txtRSYST-BCODE").text = "Password"
' session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"
' session.findById("wnd[0]/tbar[0]/btn[0]").press
' On Error Resume Next
' session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select
' session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").setFocus
' session.findById("wnd[1]/tbar[0]/btn[0]").press
' session.findById("wnd[1]/tbar[0]/btn[12]").press
' On Error GoTo 0

Dim excelApp, docList, worksheet
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False
excelApp.DisplayAlerts = False

Set docList = excelApp.Workbooks.Open("Path",0,False)
Set worksheet = docList.Worksheets(1)

Dim lastline, result
lastline = worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).Row

For i = 2 To lastline Step 1
	result = loadNewestDocumentFB03(worksheet.cells(i,1), "0002", "2018")
	If StrComp(Left(result,6), "Success", 1) = 0 Then
		worksheet.Hyperlinks.Add worksheet.Cells(i,2), worksheet.cells(i,1) & Trim(Split(result,"-")(1)), "", "Link", "Dokument"
		worksheet.Cells(i,1).Interior.Color = RGB(0,255,0)
	Else
		worksheet.Cells(i,1).Interior.Color = RGB(255,0,0)
	End If
Next

docList.Save
docList.Close
excelApp.Quit

session.findById("wns[0]/tbar[0]/okcd").text = "/nex"
session.findById("wnd[0]/tbar[0]/btn[0]").press

objShell.Run "TASKKILL.exe /F /IM SAPLogon.exe"

Function loadNewestDocumentFB03(documentID, companyCode, businessYear)
	Dim application, connection, session, SapGuiAuto
	
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
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nFB03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	If Not (StrComp(session.Info().Program, "SAPMF05L", 1) = 0) Then
		loadNewestDocumentFB03 = "Unknown Error"
		Exit Function
	End If
	' If Not (StrComp(Left(session.findById("wnd[0]").Text,24), "Screen Name", 1) = 0) Then
	' 	loadNewestDocumentFB03 = "Unknown Error"
	' 	Exit Function
	' End If
	session.findById("wnd[0]/usr/txtRF05L-BELNR").text = documentID
	session.findById("wnd[0]/usr/ctxtRF05L-BUKRS").text = companyCode
	session.findByID("wnd[0]/usr/txtRF05L-GJAHR").text = businessYear
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	' If (StrComp(Left(session.findById("wnd[0]").Text,24), "Screen Name", 1) = 0) Then
	' 	loadNewestDocumentFB03 = "Failed"
	' 	Exit Function
	' End If
	If Not (StrComp(session.findByID("wnd[0]/usr/txtBKPF-BELNR").text, docID, 1) = 0) Then
		loadNewestDocumentFB03 = "Unknown Error"
		Exit Function
	End If
	session.findByID("wnd[0]/titl/shellcont/shell").pressContectButton "%GOS_TOOLBOX"
	session.findByID("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
	
	Dim rowSelection, noOfRows, rowToDownload, rowSplit, date1, date2
	session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").SelectAll
	rowSelection = session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").SelectedRows
	If Len(rowSelection) = 0 Then
		loadNewestDocumentFB03 = "Failed"
		session.findById("wnd[1]/tbar[0]/btn[12]").press
		session.findById("wnd[0]/tbar[0]/okcd").text = "/nFB03"
		session.findById("wnd[0]/tbar[0]/btn[0]").press
		Exit Function
	ElseIf Len(rowSelection) = 1 Then
		rowToDownload = rowSelection
	Else
		rowSplit = Split(rowSelection, "-")
		rowToDownload = rowSplit(0)
		For i = CInt(rowSplit(0)) + 1 To CInt(rowSplit(1)) Step 1
			date1 = session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").getCellValue(rowToDownload, "CREADATE")
			date2 = session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").getCellValue(i, "CREADATE")
			If CDate(date1) < CDate(date2) Then
				rowToDownload = i
			End If
		Next
	End If
	session.findById("wnd[1]/usr/cntlCOntAINER_0100/shellcont/shell").currentCellColumn = "CREADATE"
	session.findById("wnd[1]/usr/cntlCOntAINER_0100/shellcont/shell").selectedRows = rowToDownload
	session.findById("wnd[1]/usr/cntlCOntAINER_0100/shellcont/shell").pressToolbarButton "%ATTA_EXPORT"
	session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\Temp\"
	Dim fileExtension
	fileExtension = session.findById("wnd[2]/usr/ctxtDY_FILENAME").text
	session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = docID & fileExtension
	session.findById("wnd[2]/tbar[0]/btn[11]").press
	session.findById("wnd[1]/tbar[0]/btn[12]").press
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nFB03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	loadNewestDocumentFB03 = "Success - " & fileExtension
End Function