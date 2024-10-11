Sub DownloadAttachmentsFromSelectedEmails()
    Dim objSelection As Selection
    Dim objMail As MailItem
    Dim objAttachment As Attachment
    Dim saveFolder As String
    Dim i As Integer
    
    ' Set the folder path where attachments will be saved
    saveFolder = "C:\YourFolderPath\"

    ' Ensure the folder path ends with a backslash
    If Right(saveFolder, 1) <> "\" Then
        saveFolder = saveFolder & "\"
    End If
    
    ' Get the selected emails
    Set objSelection = Application.ActiveExplorer.Selection
    
    ' Loop through each selected item
    For i = 1 To objSelection.Count
        ' Check if the selected item is a mail item
        If TypeName(objSelection.Item(i)) = "MailItem" Then
            Set objMail = objSelection.Item(i)
            
            ' Loop through attachments in the mail
            If objMail.Attachments.Count > 0 Then
                For Each objAttachment In objMail.Attachments
                    ' Save the attachment
                    objAttachment.SaveAsFile saveFolder & objAttachment.FileName
                Next objAttachment
            End If
        End If
    Next i
    
    MsgBox "Attachments downloaded successfully to " & saveFolder
End Sub
