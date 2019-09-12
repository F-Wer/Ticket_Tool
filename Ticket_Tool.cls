VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub EmailDirektSenden()
Dim olApp As Outlook.Application
Dim diaFolder As FileDialog
Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim lngCount As Long


   olMail.To = ""
' Recipent must be defined
   olMail.Subject = Worksheets("Tabelle1").Cells(2, 1).Value
'   Topic is hard coded in the Worksheet; Values a from a defined list
   olMail.Body = Worksheets("Tabelle1").Cells(2, 2).Value


With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show

     
        For lngCount = 1 To .SelectedItems.Count
            attFilePath = Application.FileDialog(msoFileDialogOpen).SelectedItems(lngCount)
            olMail.Attachments.Add (attFilePath)
        Next lngCount

    End With

    olMail.Display

End Sub
