Attribute VB_Name = "ProjectQAController"
Option Explicit
' +---------------------------------------------------------+
' | AUTHOR: Kevin McAleer / Sean Boyle                      |
' | EMAIL: kevin.mcaleer@advicefactory.co.uk                |
' | DATE: 26/11/2014                                        |
' | VERSION: 2.3                                            |
' | PURPOSE: Analyses a Microsoft Project file for issues   |
' +---------------------------------------------------------+
'
' This module contains a macro which will display
' QA info in a message box
' Created by Sean Boyle, SDB Projects Ltd, for the British Council September 2014
' Updated by Kevin McAleer, Advice Factory Ltd, for the British Council November 2014
' Code optimised to run in a single pass, with a status message
' Also added a dialogbox to show progress whilst running as well as updating the statusbar with progress.
' There is now an about box as well

'TODO: Refactor the variable names to improve legability
'TODO: Do file check and project status date checks before running any code.
' ChangeLog
' Added Start/Restart button
' Added new ProjectQAModel class


' ProjectQAContoller

Public Sub ProjectQA()

Dim p As New ProjectQAModel
Dim n As Integer

p.initialise
Call ProjectQAView.initialise(p)

p.continue = True ' let the macro continue
p.StatusdateUnformatted = ActiveProject.StatusDate

Call ProjectQAView.startup(p)

If p.StatusdateUnformatted = "N/A" Then
    MsgBox ("The project Status Date must be set before running this macro."), vbOKOnly, "British Council - Plan Quality Dashboard"
    GoTo Err:
End If

p.StatusDate = Format(ActiveProject.StatusDate, "dd/mm/yy")
p.FinishDate = Format(ActiveProject.Finish, "dd/mm/yy")
p.BaselineFinish = Format(ActiveProject.ProjectSummaryTask.BaselineFinish, "dd/mm/yy")
p.BaselineUnformated = ActiveProject.BaselineSavedDate(pjBaseline)
If p.BaselineUnformated <> "00:00:00" Then p.BaselinedDated = Format(ActiveProject.BaselineSavedDate(pjBaseline), "dd/mm/yy") Else p.BaselinedDated = ActiveProject.BaselineSavedDate(pjBaseline)

If ActiveProject.ProjectSummaryTask.Finish > ActiveProject.ProjectSummaryTask.BaselineFinish Then p.ProjectStatus = " late."
If ActiveProject.ProjectSummaryTask.Finish < ActiveProject.ProjectSummaryTask.BaselineFinish Then p.ProjectStatus = " early."
If ActiveProject.ProjectSummaryTask.Finish = ActiveProject.ProjectSummaryTask.BaselineFinish Then p.ProjectStatus = " on track."

ProjectQAView.Show vbModeless

p.TaskCount = ActiveProject.Tasks.Count   ' count the number of tasks in the project plan
p.codeStartTime = Now()                   ' capture the start time of the code
p.issueLog = p.issueLog & "Project Quality Assurance Analysis" & vbLf
p.issueLog = p.issueLog & "----------------------------------" & vbLf
p.issueLog = p.issueLog & "Starting Analysis ->" & vbLf

' Main code loop
If ActiveProject.StatusDate = "NA" Then MsgBox "The Project Status Date has not been set", vbOKOnly

' Set the calulation mode to manual - this speeds up the macro
Application.Calculation = pjManual
Application.ScreenUpdating = False
n = 1
p.continue = ProjectQAView.loopstate
While (p.continue = True) And (n <= p.TaskCount)
    
    ' calculate the time remaining to run
    If p.percentComplete <> 0 Then
        p.eta = (Now() - p.codeStartTime) / (p.percentComplete) * (p.percentComplete - 100)
    End If
    
    ' draw the dialog box and update the fields
    ProjectQAView.lblTask.Caption = "Task No: " & p.TCount & "/" & p.TaskCount ' display the task number
    
    ' update the title with percentage complete and time remaining.
    ProjectQAView.Caption = "Microsoft Project Quality Assurance Check | " & p.percentComplete & "% Complete | time remaining: " & p.eta
    
    ' add the summary tasks
    If ActiveProject.Tasks(n).Summary = True Then p.SLcount = p.SLcount + 1
1
    'If (Not ActiveProject.Tasks(n) Is Nothing) And (ActiveProject.Tasks(n).percentComplete <> 100) Then
    If (ActiveProject.Tasks(n).percentComplete <> 100) And (Not ActiveProject.Tasks(n).Summary) Then
        p.percentComplete = p.TCount / p.TaskCount * 100              ' calculate the percent complete
        p.TCount = p.TCount + 1
        ActiveProject.Tasks(n).Text25 = ""                                           ' 1 Clear issue field (Needed if errors written to Text25)
        
        ' check that are just for info
        If ProjectQAView.cbInfoSelectAll Then
            Call p.check_Inbound(n)
            Call ProjectQAView.update_inbound(p)
            Call p.check_Outbound(n)
            Call ProjectQAView.update_outbound(p)
            Call p.count_RemainingTasks(n)
            Call ProjectQAView.update_remainingTasks(p)
            Call p.count_Milestones(n)
            Call ProjectQAView.update_milestone(p)
            Call p.count_tasks8weeks(n)
            Call ProjectQAView.update_task8weeks(p)
        End If
        
        ' Update the Issue log display
        Call ProjectQAView.refreshIssueLog(p)
        
        ' checks that detect issues
        If ProjectQAView.cbOutBound Then
            Call p.count_Outbound_withoutPred(n)
            Call ProjectQAView.update_Outbound_withoutPred(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cb5days.Value Then
            Call p.count_tasks5days(n)
            Call ProjectQAView.update_Tasks5Days(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbMilestones.Value Then
            Call p.count_MilestonesNoSuccess(n)
            Call ProjectQAView.update_MilestonesNoSuccess(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cb20days.Value Then
            Call p.count_TasksOver20d(n)
            Call ProjectQAView.update_TasksOver20d(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbSuccess.Value Then
            Call p.check_NoSuccess(n)
            Call ProjectQAView.update_NoSuccess(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbNoPred.Value Then
            Call p.check_NoPred(n)
            Call ProjectQAView.update_NoPred(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbNegFloat.Value Then
            Call p.check_NegFloat(n)
            Call ProjectQAView.update_NegFloat(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbWorkPast.Value Then
            Call p.check_WorkInPast(n)
            Call ProjectQAView.update_WorkInPast(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbFuture.Value Then
            Call p.check_WorkInFuture(n)
            Call ProjectQAView.update_WorkInFuture(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbSummary.Value Then
            Call p.check_SummaryResources(n)
            Call ProjectQAView.update_SummaryResources(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbManual.Value Then
            Call p.check_ManuallyScheduled(n)
            Call ProjectQAView.update_ManuallyScheduled(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        If ProjectQAView.cbHardConstraints.Value Then
            Call p.check_HardConstraints(n)
            Call ProjectQAView.update_HardConstraints(p)
            Call ProjectQAView.refreshIssueLog(p)
        End If
        
    End If
    Application.StatusBar = "Quality Assurance Analyser Running | Reading Task: " & p.TCount & "/" & p.TaskCount & " | " & p.percentComplete & "%"
    DoEvents
    p.continue = ProjectQAView.loopstate
    n = n + 1
Wend

' ********************************************************************************************
' END OF LOOP
' ********************************************************************************************

' update the form buttons to show print button and close caption
Call ProjectQAView.finished(p)

' calculate the stats for the summary
p.calculate

' add the summary information to the issue log
p.addSummaryToIssueLog

' update the issuelog textbox
Call ProjectQAView.refreshIssueLog(p)
Call ProjectQAView.refreshAll(p)
Call ProjectQAView.TogglePrintButton
'TODO The next line caused a crash when Sean ran it on Project 2010. the message was 'error 1100: the method is not available in this situation'
Application.Calculation = pjAutomatic
Err:
End Sub
