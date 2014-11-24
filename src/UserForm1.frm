VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "QA Progress"
   ClientHeight    =   8715.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630.001
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

Private Sub btnStop_Click()
    ' Stop the macro from running
    
    ' First reset the status bar
    Application.StatusBar = ""
    
    ' then make sure the options are returned to normal
    Application.ScreenUpdating = True
    Application.Calculation = pjAutomatic
    DoEvents
    
    ' then stop the macro
    End
End Sub




Private Sub cbSelectAll_Click()
    If UserForm1.cbSelectAll Then
        UserForm1.cb5days.Value = True
        UserForm1.cbOutBound.Value = True
        UserForm1.cb20days.Value = True
        UserForm1.cbNoPred.Value = True
        UserForm1.cbWorkPast.Value = True
        UserForm1.cbFuture.Value = True
        UserForm1.cbSummary.Value = True
        UserForm1.cbSuccess.Value = True
        UserForm1.cbManual.Value = True
        UserForm1.cbNegFloat.Value = True
        UserForm1.cbHardConstraints.Value = True
        UserForm1.cbMilestones.Value = True
        UserForm1.cbOutBound.Value = True
    Else
        UserForm1.cb5days.Value = False
        UserForm1.cbOutBound.Value = False
        UserForm1.cb20days.Value = False
        UserForm1.cbNoPred = False
        UserForm1.cbWorkPast.Value = False
        UserForm1.cbFuture.Value = False
        UserForm1.cbSummary.Value = False
        UserForm1.cbSuccess.Value = False
        UserForm1.cbManual.Value = False
        UserForm1.cbNegFloat.Value = False
        UserForm1.cbHardConstraints.Value = False
        UserForm1.cbMilestones.Value = False
        UserForm1.cbOutBound.Value = False
    End If
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


Private Sub UserForm_Click()

End Sub
