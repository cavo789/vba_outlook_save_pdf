' ==============================================================
' Description: Outlook macro to save a selected item in the 
' pdf-format
'
' Requires Word 2007 SP2 or Word 2010
' Requires a reference to "Microsoft Word <version> Object Library"
' (version is 12.0 or 14.0)
' In VBA Editor; Tools-> References...
'
' Author: Robert Sparnaaij
' Modified by : Christophe Avonture
' website: http://www.howto-outlook.com/howto/saveaspdf.htm
'====================================================

Private Const cFolder As String = "C:\Christophe\Mails\PDF"

Sub SaveAsPDFfile()

	Dim MyOlNamespace As Outlook.NameSpace
	Dim FSO As Object, TmpFolder As Object
	Dim fdfs As FileDialogFilters
	Dim fdf As FileDialogFilter
	Dim wrdApp As Object
	Dim wrdDoc As Object
	Dim dlgSaveAs As FileDialog
	Dim I As Integer
	Dim msgFileName As String
	Dim strCurrentFile As String
	
	' Get all selected items
	Set MyOlNamespace = Application.GetNamespace("MAPI")
	Set MyOlSelection = Application.ActiveExplorer.Selection

	' Make sure at least one item is selected
	If MyOlSelection.Count <> 1 Then
		Response = MsgBox("Please select a single item", _
			vbExclamation, "Save as PDF")
		Exit Sub
	End If

	' Retrieve the selected item
	Set MySelectedItem = MyOlSelection.Item(1)

	' Get the user's TempFolder to store the item in	
	Set FSO = CreateObject("scripting.filesystemobject")
	Set tmpFileName = FSO.GetSpecialFolder(2)

	' Construct the filename for the temp mht-file
	strName = "outlook"
	tmpFileName = tmpFileName & "\" & strName & ".mht"

	' Save the mht-file
	MySelectedItem.SaveAs tmpFileName, olMHTML

	' Create a Word object
	Set wrdApp = CreateObject("Word.Application")

	' Open the mht-file in Word without Word visible
	Set wrdDoc = wrdApp.Documents.Open(FileName:=tmpFileName, Visible:=False)
   
	wrdApp.Visible = True

	' Define the SafeAs dialog	
	Set dlgSaveAs = wrdApp.FileDialog(msoFileDialogSaveAs)

	' Determine the FilterIndex for saving as a pdf-file
	' Get all the filters
	Set fdfs = dlgSaveAs.Filters

	' Loop through the Filters and exit when "pdf" is found	
	I = 0
	
	For Each fdf In fdfs
		I = I + 1
		
		If InStr(1, fdf.Extensions, "pdf", vbTextCompare) > 0 Then 
			Exit For
		End If		
	Next fdf

	' Set the FilterIndex to pdf-files
	dlgSaveAs.FilterIndex = I

	' Construct a safe file name from the message subject
	msgFileName = MySelectedItem.Subject

	Set oRegEx = CreateObject("vbscript.regexp")
	oRegEx.Global = True
	oRegEx.Pattern = "[\\/:*?""<>|]"

	' Add the received email date as prefix
	msgFileName = Format(MySelectedItem.ReceivedTime, "yyyy-mm-dd_Hh-Nn") & _
		"_" & Trim(oRegEx.Replace(msgFileName, ""))

	' Set the initial location and file name for SaveAs dialog
	dlgSaveAs.InitialFileName = cFolder & "\" & msgFileName

	' Show the SaveAs dialog and save the message as pdf
	If dlgSaveAs.Show = -1 Then
		strCurrentFile = dlgSaveAs.SelectedItems(1)

		' Verify if pdf is selected
		If Right(strCurrentFile, 4) <> ".pdf" Then
	  
			Response = MsgBox("Sorry, only saving in the pdf-format " & _
				"is supported." & vbNewLine & vbNewLine & _
				"Save as pdf instead?", vbInformation + vbOKCancel)

			If Response = vbCancel Then
				wrdDoc.Close
				wrdApp.Quit
				Exit Sub
			ElseIf Response = vbOK Then
				intPos = InStrRev(strCurrentFile, ".")
				If intPos > 0 Then 
					strCurrentFile = Left(strCurrentFile, intPos - 1)
				End If
				strCurrentFile = strCurrentFile & ".pdf"
			End If
					 
		End If

		' Save as pdf
		wrdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:= _
			strCurrentFile, ExportFormat:= _
			wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
			wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=0, To:=0, _
			Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
			CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
			BitmapMissingFonts:=True, UseISO19005_1:=False
	End If

	Set dlgSaveAs = Nothing

	' Close the document and Word

	wrdDoc.Close
	wrdApp.Quit

	' Cleanup

	Set MyOlNamespace = Nothing
	Set MyOlSelection = Nothing
	Set MySelectedItem = Nothing
	Set wrdDoc = Nothing
	Set wrdApp = Nothing
	Set oRegEx = Nothing

End Sub