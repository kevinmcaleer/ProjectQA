VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectQAView 
   Caption         =   "QA Progress"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   OleObjectBlob   =   "ProjectQAView.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectQAView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ProjectQAView

Public loopstate As Boolean
Public initstate As Boolean

Private Sub btnAbout_Click()
    FrmAbout.Show
End Sub

Private Sub btnOk_Click()
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

Private Sub btnstart_Click()
    ' Start the QA Process
    loopstate = True
    Me.btnstart.Enabled = False
    Me.btnStop.Enabled = True
    Call ProjectQA
End Sub

Private Sub btnStop_Click()
    loopstate = False
    Me.btnstart.Enabled = True
End Sub

Private Sub cbSelectAll_Click()
    If Me.cbSelectAll Then
        Me.cb5days.Value = True
        Me.cbOutBound.Value = True
        Me.cb20days.Value = True
        Me.cbNoPred.Value = True
        Me.cbWorkPast.Value = True
        Me.cbFuture.Value = True
        Me.cbSummary.Value = True
        Me.cbSuccess.Value = True
        Me.cbManual.Value = True
        Me.cbNegFloat.Value = True
        Me.cbHardConstraints.Value = True
        Me.cbMilestones.Value = True
        Me.cbOutBound.Value = True
    Else
        Me.cb5days.Value = False
        Me.cbOutBound.Value = False
        Me.cb20days.Value = False
        Me.cbNoPred.Value = False
        Me.cbWorkPast.Value = False
        Me.cbFuture.Value = False
        Me.cbSummary.Value = False
        Me.cbSuccess.Value = False
        Me.cbManual.Value = False
        Me.cbNegFloat.Value = False
        Me.cbHardConstraints.Value = False
        Me.cbMilestones.Value = False
        Me.cbOutBound.Value = False
    End If
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
End Sub


Private Sub CommandButton2_Click()
Dim fso As New FileSystemObject
 'print the issuelog

'Create a FileSystemObject.
    'Set fso = New FileSystemObject
    ' Declare a TextStream.
    Dim stream As TextStream

    'Create a TextStream.
    Set stream = fso.CreateTextFile("h:\MSP_Issues.log", True)
    stream.Write (Me.TextBox1.Value)
    
    ' Close the file.
    stream.Close

    'Code to print text file
    Shell ("c:\windows\notepad.exe /p h:\MSP_Issues.log")
End Sub



Public Sub initialise(pr As ProjectQAModel)
    ProjectQAView.TextBox1.Value = "" ' empty the textbox content
    
    'Setup the Userform with values from the project file
    ProjectQAView.lblFinishDate.Caption = pr.FinishDate
    ProjectQAView.lblStatusDate.Caption = pr.StatusDate
    ProjectQAView.lblImboundDependencies.Caption = pr.DIcount
    ProjectQAView.lblOutboundDependencies.Caption = pr.DOcount
    ProjectQAView.lblRemainingTasks.Caption = pr.ITcount
    ProjectQAView.lblKeyMilestones.Caption = pr.TMcount
    ProjectQAView.lblOutBoundPred.Caption = pr.MPcount
    ProjectQAView.lbl5dayslong.Caption = pr.LTcount
    ProjectQAView.lblTasksFinishingSoon.Caption = pr.Fcount
    ProjectQAView.CommandButton2.Enabled = False
    ProjectQAView.btnOk.Caption = "Ok"
    ProjectQAView.lblover20D.Caption = pr.TLcount
    ProjectQAView.lblWorkInPast.Caption = pr.NUcount
    ProjectQAView.lblCompleteInFuture.Caption = pr.IFcount
    ProjectQAView.lblResourcesAssigned.Caption = pr.SRcount
    ProjectQAView.lblNoPred.Caption = pr.NScount
    ProjectQAView.lblNoSuccessors.Caption = pr.NPcount
    ProjectQAView.LblManualTasks.Caption = pr.MAcount
    ProjectQAView.lblMilestoneSuccess.Caption = pr.MScount
    ProjectQAView.lblNegFloat.Caption = pr.NFcount
    ProjectQAView.lblHardConstraints.Caption = pr.HCcount
    ProjectQAView.lblBaseline = pr.BaselineUnformated ' baseline date

    'set all the checkboxes to true / ticked
    If initstate Then
    ProjectQAView.cb5days.Value = True
    ProjectQAView.cbOutBound.Value = True
    ProjectQAView.cb20days.Value = True
    ProjectQAView.cbNoPred.Value = True
    ProjectQAView.cbWorkPast.Value = True
    ProjectQAView.cbFuture.Value = True
    ProjectQAView.cbSummary.Value = True
    ProjectQAView.cbSuccess.Value = True
    ProjectQAView.cbManual.Value = True
    ProjectQAView.cbNegFloat.Value = True
    ProjectQAView.cbHardConstraints.Value = True
    ProjectQAView.cbOutBound.Value = True
    ProjectQAView.cbMilestones.Value = True
    initstate = False
    End If
End Sub

Function update_outbound(pr As ProjectQAModel)
    Me.lblOutboundDependencies.Caption = pr.DOcount
End Function

Function update_inbound(pr As ProjectQAModel)
    Me.lblImboundDependencies.Caption = pr.DIcount
End Function

Function update_remainingTasks(pr As ProjectQAModel)
    Me.lblRemainingTasks.Caption = pr.ITcount
End Function

Function update_milestone(pr As ProjectQAModel)
    Me.lblKeyMilestones.Caption = pr.TMcount
End Function

Function update_task8weeks(pr As ProjectQAModel)
    Me.lblTasksFinishingSoon.Caption = pr.Fcount
End Function

Function update_Outbound_withoutPred(pr As ProjectQAModel)
    Me.lblOutBoundPred.Caption = pr.MPcount
End Function

Function update_Tasks5Days(pr As ProjectQAModel)
    Me.lbl5dayslong.Caption = pr.LTcount
End Function

Function update_MilestonesNoSuccess(pr As ProjectQAModel)
    Me.lblMilestoneSuccess.Caption = pr.MScount
End Function

Function update_TasksOver20d(pr As ProjectQAModel)
    Me.lblover20D.Caption = pr.TLcount
End Function

Function update_NoSuccess(pr As ProjectQAModel)
    Me.lblNoSuccessors = pr.NScount
End Function

Function update_NoPred(pr As ProjectQAModel)
    Me.lblNoPred.Caption = pr.NPcount
End Function

Function update_NegFloat(pr As ProjectQAModel)
    Me.lblNegFloat = pr.NFcount
End Function

Function update_WorkInPast(pr As ProjectQAModel)
    Me.lblWorkInPast.Caption = pr.NUcount
End Function

Function update_WorkInFuture(pr As ProjectQAModel)
    Me.lblCompleteInFuture.Caption = pr.IFcount
End Function

Function update_SummaryResources(pr As ProjectQAModel)
    Me.lblResourcesAssigned.Caption = pr.SRcount
End Function

Function update_ManuallyScheduled(pr As ProjectQAModel)
    Me.LblManualTasks = pr.MAcount
End Function

Function update_HardConstraints(pr As ProjectQAModel)
    Me.lblHardConstraints = pr.HCcount
End Function

Public Sub startup(pr As ProjectQAModel)
    Application.ScreenUpdating = True
    DoEvents
    Call Me.refreshAll(pr)
End Sub

Public Sub finished(pr As ProjectQAModel)
    Me.CommandButton2.Enabled = True ' show the print button
    Me.btnOk.Caption = "Close"
    Application.StatusBar = "" ' set the status bar back to normal
    Me.Show vbModeless
    Me.lblTask.Caption = "Task No: " & pr.TCount & "/" & pr.TaskCount ' display the task number
    Me.Caption = "Microsoft Project Quality Assurance Check | " & pr.percentComplete & "% Complete"
    Me.Frame2.Caption = "Issues: " & pr.totalIssues
    Me.btnstart.Enabled = True
    Me.btnStop.Enabled = False
    pr.percentComplete = 100
End Sub

Public Sub refreshIssueLog(pr As ProjectQAModel)
    Me.TextBox1.Value = pr.issueLog
End Sub

Public Sub refreshAll(pr As ProjectQAModel)
    ' Update the userform with the latest
    Me.lblFinishDate.Caption = pr.FinishDate
    Me.lblStatusDate.Caption = pr.StatusDate
    Me.lblImboundDependencies.Caption = pr.DIcount
    Me.lblOutboundDependencies.Caption = pr.DOcount
    Me.lblRemainingTasks.Caption = pr.ITcount
    Me.lblKeyMilestones.Caption = pr.TMcount
    Me.lblOutBoundPred.Caption = pr.MPcount
    Me.lbl5dayslong.Caption = pr.LTcount
    Me.lblTasksFinishingSoon.Caption = pr.Fcount
    Me.CommandButton2.Enabled = False
    Me.btnOk.Caption = "Ok"
    Me.lblover20D.Caption = pr.TLcount
    Me.lblWorkInPast.Caption = pr.NUcount
    Me.lblCompleteInFuture.Caption = pr.IFcount
    Me.lblResourcesAssigned.Caption = pr.SRcount
    Me.lblNoPred.Caption = pr.NPcount
    Me.lblNoSuccessors = pr.NScount
    Me.LblManualTasks = pr.MAcount
    Me.lblNegFloat = pr.NFcount
    Me.lblMilestoneSuccess.Caption = pr.MScount
    Me.lblTask.Caption = "Task No: " & pr.TCount + pr.SLcount & "/" & pr.TaskCount
    ' update the title with percentage complete and time remaining.
    Me.Caption = "Microsoft Project Quality Assurance Check | " & pr.percentComplete & "% Complete"
    
End Sub

Public Function TogglePrintButton()
    If Me.CommandButton2.Enabled Then
    Me.CommandButton2.Enabled = False
    Else
    Me.CommandButton2.Enabled = True
    End If
End Function
