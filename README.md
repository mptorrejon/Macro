# Macro #Outlook macro that syncs a template email body for a certain email into a Excel sheet.Run by the user.

#SETUP OUTLOOK TO ALLOW MACROS
1.- Select "file">"Options"
2.- Select “Trust Center” on the left pane, then select the “Trust Center Settings…” button.
3.-Select “Macro Settings” on the left pane, then the desired setting.
	+Disable all macros without notification.
	+Notifications for digitally signed macros, all other macros disabled.
	+Notifications for all macros.
	+Enable all macros

You may have to check the “Apply macro security settings to installed add-ins” to allow macros to work with add-ins.
4.-Click “OK“, then close and re-open Outlook for the setting to take effect.

#ADD DEVELOPER RIBBON IN OUTLOOK
1.-Select "File"->"Options"->"Customize Ribbon" and check "Developer check box" on the right panel.
2.- Click "Ok" to close the dialog box.

#ADD MACRO TO OUTLOOK
1.- Click the Developer tab and then click on the "Macro" dropdown and select "Macros"
2.-Name your macro Download
3.- Double click "ThisOutlookSession", copy and paste OutlookSyncExcel.txt file into the module window.

#LINES TO BE EDITED BEFORE RUNNING MACRO
+line13:  MailBoxName = "MTorrejon@teamwash.com" #change this to your outlook email account
+Line16: Pst_Folder_Name = "Inbox" #change this to pst folder you will be using 'Inbox' for this case
+line40: Set a = fs.CreateTextFile("\\tw-pdc\documents$\mtorrejon\my documents\testfile.xls", True) #Change this to the folder path you will be utilizing
+line53: Set xlWB = xlAppl.Workbooks.Open("\\tw-pdc\documents$\mtorrejon\my documents\testfile.xls", notify:=False) #Change this to the folder path you will be utilizing
+line 68: And Folder.Items.item(iRow).SenderEmailAddress = "mptorrejon@gmail.com" Then #change this to the email that you will be filtering out with.