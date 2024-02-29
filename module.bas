' --------------------------------------------------
'
' Outlook macro to save a selected item(s) as pdf
' files on your hard-disk. You can select as many mails
' you want and hop hop hop each mails will be saved on
' your disk.
'
' Note : requires Winword (referenced by late-bindings)
'
' @see https://github.com/cavo789/vba_outlook_save_pdf
'
' --------------------------------------------------

Option Explicit

Private Const cFolder As String = "C:\Mails\"

Private objWord As Object

' --------------------------------------------------
'
' Ask the user for the folder where to store emails
'
' --------------------------------------------------
Private Function AskForTargetFolder(ByVal sTargetFolder As String) As String

    Dim dlgSaveAs As FileDialog

    sTargetFolder = Trim(sTargetFolder)

    ' Be sure that sTargetFolder is well ending by a slash
    If Not (Right(sTargetFolder, 1) = "\") Then
        sTargetFolder = sTargetFolder & "\"
    End If

    ' Already initialized before, so it's safe to just get the object
    Set dlgSaveAs = objWord.FileDialog(msoFileDialogFolderPicker)

    With dlgSaveAs
        .Title = "Select a Folder where to save emails"
        .AllowMultiSelect = False
        .InitialFileName = sTargetFolder
        .Show

        On Error Resume Next

        sTargetFolder = .SelectedItems(1)

        If Err.Number <> 0 Then
            sTargetFolder = ""
            Err.Clear
        End If

        On Error GoTo 0

    End With

    ' Be sure that sTargetFolder is well ending by a slash
    If Not (Right(sTargetFolder, 1) = "\") Then
        sTargetFolder = sTargetFolder & "\"
    End If

    AskForTargetFolder = sTargetFolder

End Function

' --------------------------------------------------
'
' Ask the user for a filename
'
' --------------------------------------------------
Private Function AskForFileName(ByVal sFileName As String) As String

    Dim dlgSaveAs As FileDialog
    Dim wResponse As VBA.VbMsgBoxResult
    Dim wPos As Integer

    Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)

    ' Set the initial location and file name for SaveAs dialog
    dlgSaveAs.InitialFileName = sFileName

    ' Show the SaveAs dialog and save the message as pdf
    If dlgSaveAs.Show = -1 Then

        sFileName = dlgSaveAs.SelectedItems(1)

        ' Verify if pdf is selected
        If Right(sFileName, 4) <> ".pdf" Then

            wResponse = MsgBox("Sorry, only saving in the pdf-format " & _
                "is supported." & vbNewLine & vbNewLine & _
                "Save as pdf instead?", vbInformation + vbOKCancel)

            If wResponse = vbCancel Then
                sFileName = ""
            ElseIf wResponse = vbOK Then
                wPos = InStrRev(sFileName, ".")
                If wPos > 0 Then
                    sFileName = Left(sFileName, wPos - 1)
                End If
                sFileName = sFileName & ".pdf"
            End If

        End If
    End If

    ' Return the filename
    AskForFileName = sFileName

End Function

' --------------------------------------------------
'
' Do the job, process every selected emails and
' export them as .pdf files.
'
' If the user has ask for removing mails once exported,
' emails will be removed.
'
' --------------------------------------------------
Sub SaveAsPDFfile()

    Const wdExportFormatPDF = 17
    Const wdExportOptimizeForPrint = 0
    Const wdExportAllDocument = 0
    Const wdExportDocumentContent = 0
    Const wdExportCreateNoBookmarks = 0

    Dim oSelection As Outlook.Selection
    Dim oMail As Outlook.MailItem
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")

    ' Use late-bindings
    Dim objDoc As Object
    Dim oRegEx As Object

    Dim dlgSaveAs As FileDialog
    Dim objFDFS As FileDialogFilters
    Dim fdf As FileDialogFilter
    Dim I As Integer, wSelectedeMails As Integer
    Dim sFileName As String
    Dim sTempFolder As String, sTempFileName As String
    Dim sTargetFolder As String, strCurrentFile As String

    Dim bContinue As Boolean
    Dim bAskForFileName As Boolean
    Dim bRemoveMailAfterExport As Boolean

    ' Get all selected items
    Set oSelection = Application.ActiveExplorer.Selection

    ' Get the number of selected emails
    wSelectedeMails = oSelection.Count

    ' Make sure at least one item is selected
    If wSelectedeMails < 1 Then
        Call MsgBox("Please select at least one email", _
            vbExclamation, "Save as PDF")
        Exit Sub
    End If

    ' --------------------------------------------------
    bContinue = MsgBox("You're about to export " & wSelectedeMails & " " & _
        "emails as PDF files, do you want to continue? If you Yes, you'll " & _
        "first need to specify the name of the folder where to store the files", _
        vbQuestion + vbYesNo + vbDefaultButton1) = vbYes

    If Not bContinue Then
        Exit Sub
    End If

    ' --------------------------------------------------
    ' Start Word and make initializations
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False

    ' --------------------------------------------------
    ' Define the target folder, where to save emails
    sTargetFolder = AskForTargetFolder(cFolder)

    If sTargetFolder = "" Then
        objWord.Quit
        Set objWord = Nothing
        Exit Sub
    End If

    ' --------------------------------------------------
    ' Once the mail has been saved as PDF do we need to
    ' remove it?
    bRemoveMailAfterExport = MsgBox("Once the email has been " & _
        "exported and saved onto your disk, do you wish to keep " & _
        "it in your mailbox or do you want to delete it?" & vbCrLf & vbCrLf & _
        "Press Yes to keep the mail, Press No to delete the mail after exportation", _
        vbQuestion + vbYesNo + vbDefaultButton1) = vbNo

    ' --------------------------------------------------
    ' When more than one email has been selected, just ask the
    ' user if we need to ask for filenames each time (can be
    ' annoying)
    bAskForFileName = True

    If (wSelectedeMails > 1) Then
        bAskForFileName = MsgBox("You're about to save " & wSelectedeMails & " " & _
            "emails as PDF files. Do you want to see " & wSelectedeMails & " " & _
            "prompts so you can update the filename or just use the automated " & _
            "one (so no prompt)." & vbCrLf & vbCrLf & _
            "Press Yes to see prompts, Press No to use automated name", _
            vbQuestion + vbYesNo + vbDefaultButton2) = vbYes

        MsgBox "BE CAREFULL: You'll not see a progression on the screen (unfortunately, " & _
            "Outlook doesn't allow this)." & vbCrLf & vbCrLf & _
            "If you're exporting a lot of mails, the process can take a while. " & _
            "Perhaps the best way to see that things are working is to open a " & _
            "explorer window and see how files are added to the folder." & vbCrLf & vbCrLf & _
            "Once finished, you'll see a feedback message.", _
            vbInformation + vbOKOnly
    End If

    ' --------------------------------------------------
    ' Define the SaveAs dialog
    If bAskForFileName Then

        Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)

        ' --------------------------------------------------
        ' Determine the FilterIndex for saving as a pdf-file
        ' Get all the filters and make sure we've "pdf"
        Set objFDFS = dlgSaveAs.Filters

        I = 0

        For Each fdf In objFDFS
            I = I + 1

            If InStr(1, fdf.Extensions, "pdf", vbTextCompare) > 0 Then
                Exit For
            End If
        Next fdf

        Set objFDFS = Nothing

        ' Set the FilterIndex to pdf-files
        dlgSaveAs.FilterIndex = I

    End If

    ' ----------------------------------------------------
    ' Get the user's TempFolder to store the item in
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sTempFolder = objFSO.GetSpecialFolder(2)
    Set objFSO = Nothing

    ' ----------------------------------------------------
    ' We are ready to start
    ' Process every selected emails; one by one
    On Error Resume Next

    For I = 1 To wSelectedeMails

        ' Retrieve the selected email
        Set oMail = oSelection.Item(I)

        ' Construct the filename for the temp mht-file
        sTempFileName = sTempFolder & "\outlook.mht"

        ' Kill the previous file if already present
        If Dir(sTempFileName) Then Kill (sTempFileName)

        ' Save the mht-file
        oMail.SaveAs sTempFileName, olMHTML

        ' Open the mht-file in Word without Word visible
        Set objDoc = objWord.Documents.Open(FileName:=sTempFileName, Visible:=False, ReadOnly:=True)

        ' Construct a safe file name from the message subject
        sFileName = oMail.Subject

        ' Sanitize filename, remove unwanted characters
        Set oRegEx = CreateObject("vbscript.regexp")
        oRegEx.Global = True
        oRegEx.Pattern = "[\\/:*?""<>|]"

        ' Add the received email date as prefix
        sFileName = sTargetFolder & Format(oMail.ReceivedTime, "yyyy-mm-dd_Hh-Nn") & _
            "_" & Trim(oRegEx.Replace(sFileName, "")) & ".pdf"

        If bAskForFileName Then
            sFileName = AskForFileName(sFileName)
        End If

        If Not (Trim(sFileName) = "") Then

            Debug.Print "Save " & sFileName

            ' If already there, remove the file first
            If Dir(sFileName) <> "" Then
                Kill (sFileName)
            End If

            ' Save as pdf
            objDoc.ExportAsFixedFormat OutputFileName:=sFileName, _
                ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=0, To:=0, _
                Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=True, UseISO19005_1:=False

            ' And close once saved on disk
            objDoc.Close (False)

            ' Kill the mail?
            If bRemoveMailAfterExport Then
                ' Ok but only if the mail has been successfully exported
                If Dir(sFileName) <> "" Then
                    oMail.Delete
                End If
            End If

        End If

    Next I

    Set dlgSaveAs = Nothing

    On Error GoTo 0

    ' Close the document and Word

    On Error Resume Next
    objWord.Quit
    On Error GoTo 0

    ' Cleanup

    Set oSelection = Nothing
    Set oMail = Nothing
    Set objDoc = Nothing
    Set objWord = Nothing
    Set oRegEx = Nothing

    MsgBox "Done, mails have been exported to " & sTargetFolder, vbInformation

End Sub
