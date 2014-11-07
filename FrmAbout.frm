VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAbout 
   Caption         =   "About QA Analyser"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    FrmAbout.Hide
End Sub


Private Sub UserForm_Initialize()
    TextBox1.Value = "This Macro was written by Sean Boyle & Kevin McAleer in November 2014, and remains the intellectual property of the British Council."
End Sub
