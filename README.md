# iridium sbd email
- IridiumSBD_AttachmentSaver.bas
- created by mahemys 2017.10.22
- !perfect, but works!
- MIT; no license; free to use!
- update 2017-10-22; initial review

**purpose**
- Iridium Short Burst Data (SBD) Macros for Microsoft Outlook. Auto save attachments received via email to respective IMEI folders.
- Iridium SBD delivers data to the user to specified TCP/IP Socket, and or to specified Email address.

**how to use**
- Open Outlook VBA Editor => Alt + F11 => copy the *bas code and save.
- Try to Run, if error occurs follow Troubleshooting below.

**requirements**
- Iridium SBD Hardware with active Subscription or Plan.

**recommendations**
- Use TCP/IP Socket as it is robust and can handle multiple connections.
- Email is a delayed process with too many failure points as listed below.

**limitations**
- multiple devices means too many emails.
- Inbox may become Full, if hosting space is not managed properly.
- Server/Device CPU, Memory, Disk Space limitations.
- Post Processing of data or Parsing may be delayed.

**process**
- When Outlook receives an email, it will save the attachment, message body to specified file or folder.
- Auto save *.sbd attachments into 'IMEINumber' folder. Ex. 300xxxxxxxxxxx0_0xxx79.sbd
- Auto save Email Message into 'IMEINumber_TXT' folder. Ex. 300xxxxxxxxxxx0_0xxx79.sbd.txt
- SBD Email Attachment file format => IMEINumber_MOMSN.sbd

**Sample of Email from Iridium SBD Service**
```
From:	sbd service @ sbd dot iridium dot com
Sent:	Friday, November 3, 2017 7:17 PM
To:	your email @ your domain dot com
Subject:	SBD Msg From Unit: 300xxxxxxxxxxx0
Attachments:	300xxxxxxxxxxx0_0xxx79.sbd

MOMSN: 0xxx79
MTMSN: 0
Time of Session (UTC): Fri Nov  3 19:16:50 2017
Session Status: 00 - Transfer OK
Message Size (bytes): 6
```

**help with VBA** ![Alt text](/iridium-sbd-email-screenshot.png)
- refer https://www.datanumen.com/blogs/run-vba-code-outlook/

**Troubleshooting**
- If you can't get the rule to work, try adjusting your Outlook security settings:
- File tab, choose Outlook Options to open the Outlook Options dialog box, and then click Trust Center.
- Click Trust Center Settings, and then the Macro Settings option on the left.
- Select Notifications for all macros and then click OK. 
- The option allows macros to run in Outlook, but before the macro runs, Outlook prompts you to verify that you want to run the macro.
- Restart Outlook for the configuration change to take effect.

**Disable the feature that removes extra line breaks**
- This method disables the feature for all plain text items.
- To do this, follow these steps:

**For Outlook 2010 and later versions:**
- Open Outlook.
- On the File tab, click Options.
- In the Options dialog, click Mail.
- In the Message format section, clear the Remove extra line breaks in plain text messages check box.
- Click OK.

**For Outlook 2007 or earlier versions:**
- Open Outlook.
- On the Tools menu, click Options.
- On the Preferences tab, click the E-mail Options button.
- Click to clear the Remove extra line breaks in plain text messages check box.
- Click OK two times.

**Only if required**
- Restore Missing Run A Script Option In Outlook Rule
- https://www.extendoffice.com/documents/outlook/4640-outlook-rule-run-a-script-missing.html
```
add Registry -> HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Security
New >> DWORD (32-BIT Value) >> EnableUnsafeClientMailRules
Hex Value >> 0 >> Disabled
Hex Value >> 1 >> Enabled
OK
```

**footnote**
- let me know if you find any bugs!
- Thank you mahemys
