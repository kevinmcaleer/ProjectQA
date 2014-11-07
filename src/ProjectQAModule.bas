Attribute VB_Name = "ProjectQAModule"
Option Explicit
' +---------------------------------------------------------+
' | AUTHOR: Kevin McAleer / Sean Boyle                      |
' | EMAIL: kevin.mcaleer@advicefactory.co.uk                |
' | DATE: 06/11/2014                                        |
' | VERSION: 2.1                                            |
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

' Initialise variables
Dim MNcount As Integer              'milestones without preds
Dim DIcount As Integer              'dep ins
Dim DOcount As Integer              'dep outs
Dim ITcount As Integer              'Incomplete tasks
Dim LTcount As Integer              'Over 5d in 8w
Dim Fcount As Integer   'finishes in 8w count
Dim TMcount As Integer 'Key milestones
Dim MPcount As Integer 'Missing pred on outbound
Dim MScount As Integer 'Missing succ on inbound
Dim TLcount As Integer 'long tasks
Dim TCount As Integer  'task count
Dim NPcount As Integer 'no preds
Dim NScount As Integer 'no succs
Dim NFcount As Integer 'negative float
Dim NUcount As Integer 'not updated
Dim IFcount As Integer 'in future

Dim SLcount As Integer 'summary task links
Dim SRcount As Integer 'summary tasks with resources
Dim MAcount As Integer 'Manually scheduled tasks
Dim HCcount As Integer 'Hard constraints

Dim ps As String       'project status
Dim SD As String       'status date
Dim SU As String       'unformatted status date
Dim FD As Date         'project finish date
Dim BF As String       'project baseline finish
Dim BU As String       'unformatted last baselined date
Dim BD As String        'last baselined date
Dim fso As FileSystemObject ' file system object - for writing out the log file
Dim proj As Project ' the project itself
Dim t As Task ' a task object

' These Variables help the program display status items
Dim codeStartTime As Date      ' measures the start time of the code, to measure how long it took to run.
Dim codeFinishtime As Date     ' measures the finish time of the code, to measure how long it took to run.
Dim codeRunTime As Date        ' stores the result of the time it took to run the code.
Dim taskCount As Integer        ' stores the number of tasks in the project plan
Dim percentComplete As Integer
Dim issueLog As String          ' Issue Log
Dim totalIssues As Integer      ' stores the total number of issues

Sub projectQualityAssurance()

SU = ActiveProject.StatusDate
If SU = "N/A" Then
    MsgBox ("The project Status Date must be set before running this macro."), vbOKOnly, "British Council - Plan Quality Dashboard"
    GoTo Err:
End If
SD = Format(ActiveProject.StatusDate, "dd/mm/yy")
FD = Format(ActiveProject.Finish, "dd/mm/yy")
BF = Format(ActiveProject.ProjectSummaryTask.BaselineFinish, "dd/mm/yy")
BU = ActiveProject.BaselineSavedDate(pjBaseline)
If BU <> "00:00:00" Then BD = Format(ActiveProject.BaselineSavedDate(pjBaseline), "dd/mm/yy") Else BD = ActiveProject.BaselineSavedDate(pjBaseline)

If ActiveProject.ProjectSummaryTask.Finish > ActiveProject.ProjectSummaryTask.BaselineFinish Then ps = " late."
If ActiveProject.ProjectSummaryTask.Finish < ActiveProject.ProjectSummaryTask.BaselineFinish Then ps = " early."
If ActiveProject.ProjectSummaryTask.Finish = ActiveProject.ProjectSummaryTask.BaselineFinish Then ps = " on track."

' New Optimised Loop, includes all checks
'
' 1 Clear issue field (Needed if errors written to Text25)
' 2 Count inbound dependencies. Information
' 3 Count outbound dependencies. Information
' 4 Count remaining tasks. Information
' 5 Count key milestones. Information
' 6 Count outbound milestones without predeccessors. Issue
' 7 Count tasks finishing within 8w. Information
' 8 Count tasks over 5 days long within next 8 weeks. Issue
' 9 Count inbound milestones with no successors. Issue
' 10 Count tasks over 20d long. Issue
' 11 Count tasks with no successors. Issue
' 12 Count tasks with no predecessors. Issue
' 13 Count Tasks with negative float. Issue
' 14 Count tasks with work in the past. Issue
' 15 Count tasks with work complete in the future. Issue
' 16 Count summary tasks with resources assigned

' Clear Variables
TCount = 0                              ' Clears the number of tasks
DIcount = 0                             ' 2 Clears Inbound Dependencies (Info)
DOcount = 0                             ' 3 Clears Outbound Dependencies (Info)
ITcount = 0                             ' 4 Clears remaining tasks (Info)
TMcount = 0                             ' 5 Clears key milestones (Info)
MPcount = 0                             ' 6 Clears outbound milestones without predeccessors (Issue)
Fcount = 0                              ' 7 Clears tasks Finishing within the next 8 weeks (Info)
LTcount = 0                             ' 8 Clears long tasks (over 5 days) in the next 8 weeks (Issue)
MScount = 0                             ' 9 Clears inbound milestone with no succcessors (Issue)
TLcount = 0                             ' 10 Clears tasks over 20 days (Issue)
NScount = 0                             ' 11 Clears tasks with no successors (Issue)
NPcount = 0                             ' 12 Clears tasks with no precessors (Issue)
NFcount = 0                             ' 13 Clears tasks with negative float (Issue)
NUcount = 0                             ' 14 Clears tasks with work in the past (Issue)
IFcount = 0                             ' 15 Clears tasks with work complete in the future (Issue)
SRcount = 0                             ' 16 Clears sumary tasks with assigned resource (Issue)
MAcount = 0                             ' 17 Clears manually scheduled tasks (Issue)
HCcount = 0                             ' 18 Clears Hard Constraints (Issue)
percentComplete = 0                     ' Clear the percentage complete
UserForm1.TextBox1.Value = ""           ' empty the textbox content
issueLog = ""                           ' clear the issue log
totalIssues = 0                         ' clear the total number of issues

'Application.StatusBar = "Project QA Running"
'Application.ScreenUpdating = True
'Application.DisplayAlerts = True

taskCount = ActiveProject.Tasks.Count   ' count the number of tasks in the project plan
UserForm1.Show vbModeless

'Setup the Userform with values from the project file
UserForm1.lblFinishDate.Caption = FD
UserForm1.lblStatusDate.Caption = SD
UserForm1.lblImboundDependencies.Caption = DIcount
UserForm1.lblOutboundDependencies.Caption = DOcount
UserForm1.lblRemainingTasks.Caption = ITcount
UserForm1.lblKeyMilestones.Caption = TMcount
UserForm1.lblOutBoundPred.Caption = MPcount
UserForm1.lbl5dayslong.Caption = LTcount
UserForm1.lblTasksFinishingSoon.Caption = Fcount
UserForm1.CommandButton2.Enabled = False
UserForm1.CommandButton1.Caption = "Ok"
UserForm1.lblover20D.Caption = TLcount
UserForm1.lblWorkInPast.Caption = NUcount
UserForm1.lblCompleteInFuture.Caption = IFcount
UserForm1.lblResourcesAssigned.Caption = SRcount
UserForm1.lblNoPred.Caption = NScount
UserForm1.lblNoSuccessors = NPcount
UserForm1.LblManualTasks = MAcount

codeStartTime = Now()                   ' capture the start time of the code
issueLog = issueLog & "Project Quality Assurance Analysis" & vbLf
issueLog = issueLog & "----------------------------------" & vbLf
issueLog = issueLog & "Starting Analysis ->" & vbLf

' Main code loop
'MsgBox ActiveProject.StatusDate, vbOKOnly

If ActiveProject.StatusDate = "NA" Then MsgBox "The Project Status Date has not been set", vbOKOnly
For Each t In ActiveProject.Tasks
    
    
    UserForm1.lblTask.Caption = "Task No: " & TCount & "/" & taskCount ' display the task number
    UserForm1.Caption = "Microsoft Project Quality Assurance Check | " & percentComplete & "% Complete"
    If Not t Is Nothing Then
        percentComplete = TCount / taskCount * 100              ' calculate the percent complete
        TCount = TCount + 1
        t.Text25 = ""                                           ' 1 Clear issue field (Needed if errors written to Text25)
        If t.Text14 = "In" Then                                 ' 2 Count inbound dependencies. Information
            DIcount = DIcount + 1
            UserForm1.lblImboundDependencies.Caption = DIcount
        End If
        
        If t.Text14 = "Out" Then                                ' 3 Count outbound dependencies. Information
            DOcount = DOcount + 1
            UserForm1.lblOutboundDependencies.Caption = DOcount
        End If
        
        If t.percentComplete <> 100 And t.Summary = False Then  ' 4 Count remaining tasks. Information
            ITcount = ITcount + 1
            UserForm1.lblRemainingTasks.Caption = ITcount
        End If
        
        If t.Text10 = "Yes" Then                                ' 5 Count key milestones. Information
            TMcount = TMcount + 1
            UserForm1.lblKeyMilestones.Caption = TMcount
        End If
        
        If t.Text14 = "Out" And t.Predecessors = "" And t.percentComplete <> 100 Then
            MPcount = MPcount + 1                               ' 6 Count outbound milestones without predeccessors. Issue
            UserForm1.lblOutBoundPred.Caption = MPcount
            issueLog = issueLog + "Task no " & t & " has an outbound milestone without a predeccessor" & vbLf
            t.Text25 = t.Text25 & ". has an outbound milestone without a predeccessor"
        End If
        
        If t.Finish < (ActiveProject.StatusDate + 56) And t.percentComplete <> 100 And t.Summary = False Then
            Fcount = Fcount + 1                                 ' 7 Count tasks finishing within 8w. Information
            UserForm1.lblTasksFinishingSoon.Caption = Fcount
        End If
        
        If t.Start < (ActiveProject.StatusDate + 56) And t.Duration > 2400 And t.Summary = False And t.percentComplete <> 100 Then
            LTcount = LTcount + 1                               ' 8 Count tasks over 5 days long within next 8 weeks. Issue
            UserForm1.lbl5dayslong.Caption = LTcount
            issueLog = issueLog + "Task no " & t & " is within the next 8 weeks and is more than 5 days in duration" & vbLf
            t.Text25 = t.Text25 & ". is within the next 8 weeks and is more than 5 days in duration"
        End If
        
        If t.Text14 = "In" And t.Successors = "" And t.percentComplete <> 100 Then
            MScount = MScount + 1                               ' 9 Count inbound milestones with no successors. Issue
            issueLog = issueLog + "Task no " & t & " is an inbound milestone with no successor" & vbLf
            t.Text25 = t.Text25 & ". is an inbound milestone with no successor"
        End If
        
        If t.Duration > 9600 And t.Summary = False And t.percentComplete <> 100 Then
            TLcount = TLcount + 1                               ' 10 Count tasks over 20d long. Issue
            t.Text25 = t.Text25 & "Over 20d. "
            UserForm1.lblover20D.Caption = TLcount
            issueLog = issueLog + "Task no " & t & " is over 20 days long" & vbLf
            
        End If
        
        If t.Successors = "" And t.Summary = False And t.percentComplete <> 100 And t.ExternalTask = False Then
            NScount = NScount + 1
            t.Text25 = t.Text25 & "No Successor. "              ' 11 Count tasks with no successors. Issue
            issueLog = issueLog + "Task no " & t & " has no successors " & vbLf
            UserForm1.lblNoSuccessors = NPcount
        End If
        
        If t.Predecessors = "" And t.Summary = False And t.percentComplete <> 100 And t.ExternalTask = False Then
            NPcount = NPcount + 1                               ' 12 Count tasks with no predecessors. Issue
            t.Text25 = t.Text25 & "No Predecessor. "
            issueLog = issueLog + "Task no " & t & " has no predeccessors " & vbLf
            UserForm1.lblNoPred.Caption = NScount
        End If
        
        If t.TotalSlack < 0 And t.Summary = False And t.percentComplete <> 100 Then
            NFcount = NFcount + 1                               ' 13 Count Tasks with negative float. Issue
            t.Text25 = t.Text25 & "Negative Float. "
            issueLog = issueLog + "Task no " & t & " has a negative float " & vbLf
            
        End If
        
        If t.percentComplete <> 100 And t.Summary = False And t.Finish < ActiveProject.StatusDate Then
            NUcount = NUcount + 1                               ' 14 Count tasks with work in the past. Issue
            t.Text25 = t.Text25 & "Incomplete in past. "
            issueLog = issueLog + "Task no " & t & " has work in the past " & vbLf
            UserForm1.lblWorkInPast.Caption = NUcount
        End If
        
        If t.percentComplete <> 100 And t.Summary = False And t.Finish > ActiveProject.StatusDate And t.Resume <= ActiveProject.StatusDate Then
            NUcount = NUcount + 1
            t.Text25 = t.Text25 & "Incomplete in past. "
            issueLog = issueLog + "Task no " & t & " has work in the past " & vbLf
            UserForm1.lblWorkInPast.Caption = NUcount
        End If
        
        If t.percentComplete = 0 And t.Summary = False And t.Finish > ActiveProject.StatusDate And t.Start < ActiveProject.StatusDate Then
            NUcount = NUcount + 1
            t.Text25 = t.Text25 & "Incomplete in past. "
            issueLog = issueLog + "Task no " & t & " has work in the past " & vbLf
            UserForm1.lblWorkInPast.Caption = NUcount
        End If
        
        If t.ActualStart <> "NA" And t.ActualStart > ActiveProject.StatusDate Then
            IFcount = IFcount + 1                               ' 15 Count tasks with work complete in the future. Issue
            t.Text25 = t.Text25 & "Complete in future. "
            issueLog = issueLog + "Task no " & t & " has work completed in the future " & vbLf
            UserForm1.lblCompleteInFuture.Caption = IFcount
        End If
        
        If t.Summary = "True" And t.percentComplete <> 100 Then
            If Not t.ResourceNames = "" Then                        ' 16 Count Summary resources with resources assigned
                SRcount = SRcount + 1
                t.Text25 = t.Text25 & "Summary resourced. "
                issueLog = issueLog + "Task no " & t & " is a summary task with resources assigned " & vbLf
                UserForm1.lblResourcesAssigned.Caption = SRcount
            End If
        End If
        
        If t.Manual = "True" And t.percentComplete <> 100 And t.ExternalTask = False Then
            MAcount = MAcount + 1                               ' 17 Count manually scheduled tasks. issue
            t.Text25 = t.Text25 & "Manually Scheduled. "
            UserForm1.LblManualTasks = MAcount
            issueLog = issueLog + "Task no " & t & " is manually assigned " & vbLf
        End If
        
        If t.ConstraintType <> 0 And t.percentComplete <> 100 And t.ExternalTask = False Then
            HCcount = HCcount + 1                               ' 18 Count hard constraints. Issue
            t.Text25 = t.Text25 & "Constrained. "
            issueLog = issueLog + "Task no " & t & " has hard constraints " & vbLf
        End If
    End If
    Application.StatusBar = "Quality Assurance Analyser Running | Reading Task: " & TCount & "/" & taskCount & " | " & percentComplete & "%"
    UserForm1.TextBox1.Value = issueLog
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    DoEvents
Next t
'UserForm1.Hide
UserForm1.lblFinishDate.Caption = FD
UserForm1.lblStatusDate.Caption = SD
UserForm1.lblImboundDependencies.Caption = DIcount
UserForm1.lblOutboundDependencies.Caption = DOcount
UserForm1.lblRemainingTasks.Caption = ITcount
UserForm1.lblKeyMilestones.Caption = TMcount
UserForm1.lblOutBoundPred.Caption = MPcount
UserForm1.lbl5dayslong.Caption = LTcount
UserForm1.lblTasksFinishingSoon.Caption = Fcount
UserForm1.CommandButton2.Enabled = False
UserForm1.CommandButton1.Caption = "Ok"
UserForm1.lblover20D.Caption = TLcount
UserForm1.lblWorkInPast.Caption = NUcount
UserForm1.lblCompleteInFuture.Caption = IFcount
UserForm1.lblResourcesAssigned.Caption = SRcount
UserForm1.lblNoPred.Caption = NScount
UserForm1.lblNoSuccessors = NPcount
UserForm1.LblManualTasks = MAcount

percentComplete = TCount / taskCount * 100
totalIssues = MPcount + LTcount + MScount + TLcount + NScount + NPcount + NFcount + NUcount + IFcount + SRcount + MAcount + HCcount ' calculate the total number of issues
UserForm1.Show vbModeless
UserForm1.lblTask.Caption = "Task No: " & TCount & "/" & taskCount ' display the task number
UserForm1.Caption = "Microsoft Project Quality Assurance Check | " & percentComplete & "% Complete"
UserForm1.Frame2.Caption = "Issues: " & totalIssues
Application.StatusBar = "" ' set the status bar back to normal

codeFinishtime = Now()                          ' capture the finish time of the code
codeRunTime = codeStartTime - codeFinishtime    ' work out how long the code took to run

issueLog = issueLog & "----------------------------------" & vbLf
issueLog = issueLog & "Analysis Complete." & vbLf
issueLog = issueLog & taskCount & " tasks took " & codeRunTime & " seconds to check." & vbLf
issueLog = issueLog & "A total of " & totalIssues & " issues were found" & vbLf
issueLog = issueLog & vbLf & vbLf & "Please contact the PMO if you need any assistance to resolve these issues."
UserForm1.TextBox1.Value = issueLog
' ********************************************************************************************
' END OF LOOP
' ********************************************************************************************

UserForm1.CommandButton2.Enabled = True ' show the print button
UserForm1.CommandButton1.Caption = "Close"
Err:
End Sub



