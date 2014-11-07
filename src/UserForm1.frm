VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "QA Progress"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btnAbout_Click()
    FrmAbout.Show
End Sub

Private Sub CommandButton1_Click()
UserForm1.Hide
End Sub


Private Sub CommandButton2_Click()
 'print the issuelog

'Create a FileSystemObject.
    Set fso = New FileSystemObject
    ' Declare a TextStream.
    Dim stream As TextStream

    'Create a TextStream.
    Set stream = fso.CreateTextFile("h:\MSP_Issues.log", True)
    stream.Write (issueLog)
    
    ' Close the file.
    stream.Close

    'Code to print text file
    Shell ("c:\windows\notepad.exe /p h:\MSP_Issues.log")
End Sub

