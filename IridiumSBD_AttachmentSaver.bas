Attribute VB_Name = "Outlook_SBDAttachmentSaver"
Public Sub saveAttachtoDisk(itm As Outlook.MailItem)
Dim objAtt As Outlook.Attachment
Dim saveFolder As String
saveFolder = "C:\Iridium\SBD_IMEI"

     'simply save to a folder
     'For Each objAtt In itm.Attachments
          'objAtt.SaveAsFile saveFolder & "\" & objAtt.DisplayName
          'objAtt.SaveAsFile saveFolder & "\" & dateFormat & objAtt.DisplayName
          'Set objAtt = Nothing
     'Next
     
     'based on some extension
     'If InStr(objAtt.DisplayName, ".xml") Then
          'objAtt.SaveAsFile saveFolder & "\" & objAtt.DisplayName
     'End If
     
     'Write to Textfile
     'Dim fso As Object
     'Set fso = CreateObject("Scripting.FileSystemObject")
     'Dim Fileout As Object
     'Dim TextFilePath As String
     'TextFilePath = saveFolder & "\SBD_Files_Log.txt"
     'Set Fileout = fso.CreateTextFile(TextFilePath, True, True)
     
     'split extension, IMEINumber, MessageNumber
     For Each objAtt In itm.Attachments
          Dim Full_FileName As String
          Dim Split_FullFileName() As String
          Dim Split_FileName() As String
          Dim IMEI_Number As String
          Dim Message_Number As String
          Full_FileName = objAtt.DisplayName
          Split_FullFileName = Split(Full_FileName, ".")
          Split_FileName = Split(Split_FullFileName(0), "_")
          IMEI_Number = Split_FileName(0)
          Message_Number = Split_FileName(1)
          
     'Save Email Body as Text file in a folder based on IMEI_TXT
     If InStr(Full_FileName, IMEI_Number) Then
          'Check for Folder else create
          Dim EmailBodySavePath As String
          EmailBodySavePath = saveFolder & "\" & IMEI_Number & "_TXT"
     If Dir(EmailBodySavePath, vbDirectory) = "" Then
          MkDir EmailBodySavePath
     End If
          'Check for file
          Dim EmailTxtFileName As String
          EmailTxtFileName = EmailBodySavePath & "\" & objAtt.DisplayName & ".txt"
     If Dir(EmailTxtFileName) <> "" Then
     Else
          'Save Email Body as Text file in IMEI_TXT Folder
          Const olTXT = 0
          itm.SaveAs EmailTxtFileName, olTXT
     End If
     End If
     
     'Save Attachment in a folder based on IMEI_Number
     If InStr(Full_FileName, IMEI_Number) Then
          'Check for Folder else create
          Dim AttachmentSavePath As String
          AttachmentSavePath = saveFolder & "\" & IMEI_Number
     If Dir(AttachmentSavePath, vbDirectory) = "" Then
          MkDir AttachmentSavePath
     End If
          'Check for file
          Dim AttachmentFileName As String
          AttachmentFileName = AttachmentSavePath & "\" & objAtt.DisplayName
     If Dir(AttachmentFileName) <> "" Then
     Else
          'save Attachment file in IMEI Folder
          objAtt.SaveAsFile AttachmentFileName
          Set objAtt = Nothing
     End If
     End If
     Next
End Sub
