VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bes_Change()
Dim ws As Worksheet

Set ws = Worksheets("Tabelle1")
With ws
.Range("B2").Value = Bes.Text
End With
End Sub

Private Sub SendE_Click()
Dim olApp As Outlook.Application
Dim diaFolder As FileDialog
Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim lngCount As Long


   olMail.To = ""
' Hier oben zwischen die "" Email adresse eintragen an die es gehen soll.
   olMail.Subject = Worksheets("Tabelle1").Cells(2, 1).Value
'   Das Thema ist in der Tabelle 2 hart drin; dies kann durch ändern der Liste geändert werden
   olMail.Body = Worksheets("Tabelle1").Cells(2, 2).Value
'   Die Beschreibung darf nur in die Spalte B Zeile 2;

'   Es müssen die Felder noch als nicht änderbar eingestellt werden ( Überprüfen dann schützen und dort beides wählen und ein Passwort reinmachen) und die Tabelle 2 muss versteckt werden

'   Dieses Skript muss noch pw geschützt werden. Oben auf Extras-> Eigenschaften von VBAProject dort auf Schutz und dort ändern


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




Private Sub UserForm_Initialize()
With Me
       .Width = 500
       .Height = 370
    End With
CB.List = Tabelle2.Range("A1:A10").Value
'Hier müssten die Range der Liste der Betreffe reingeschrieben werden.
End Sub
