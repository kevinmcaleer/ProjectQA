VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectQAModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ProjectQAModel

' Initialise variables
Public MNcount As Integer              'milestones without preds
Public DIcount As Integer              'dep ins
Public DOcount As Integer              'dep outs
Public ITcount As Integer              'Incomplete tasks
Public LTcount As Integer              'Over 5d in 8w
Public Fcount As Integer   'finishes in 8w count
Public TMcount As Integer 'Key milestones
Public MPcount As Integer 'Missing pred on outbound
Public MScount As Integer 'Missing succ on inbound
Public TLcount As Integer 'long tasks
Public TCount As Integer  'task count
Public NPcount As Integer 'no preds
Public NScount As Integer 'no succs
Public NFcount As Integer 'negative float
Public NUcount As Integer 'not updated
Public IFcount As Integer 'in future
Public SLcount As Integer 'summary task links
Public SRcount As Integer 'summary tasks with resources
Public MAcount As Integer 'Manually scheduled tasks
Public HCcount As Integer 'Hard constraints
Public ProjectStatus As String       'project status
Public StatusDate As String       'status date
Public StatusdateUnformatted As String       'unformatted status date
Public FinishDate As Date         'project finish date
Public BaselineFinish As String       'project baseline finish
Public BaselineUnformated As String       'unformatted last baselined date
Public BaselinedDated As String        'last baselined date
Public fso As FileSystemObject ' file system object - for writing out the log file
Public proj As Project ' the project itself
Private t As task ' a task object
Public continue As Boolean

' These Variables help the program display status items
Public codeStartTime As Date      ' measures the start time of the code, to measure how long it took to run.
Public codeFinishtime As Date     ' measures the finish time of the code, to measure how long it took to run.
Public codeRunTime As Date        ' stores the result of the time it took to run the code.
Public TaskCount As Integer        ' stores the number of tasks in the project plan
Public percentComplete As Integer
Public issueLog As String           ' Issue Log
Public totalIssues As Integer      ' stores the total number of issues
'Private n As Integer ' loop counter
Public eta As Date ' time remaining to calculate


Public Sub initialise()
    ' Clear Variables
    Me.TCount = 0                              ' Clears the number of tasks
    Me.DIcount = 0                             ' 2 Clears Inbound Dependencies (Info)
    Me.DOcount = 0                             ' 3 Clears Outbound Dependencies (Info)
    Me.ITcount = 0                             ' 4 Clears remaining tasks (Info)
    Me.TMcount = 0                             ' 5 Clears key milestones (Info)
    Me.MPcount = 0                             ' 6 Clears outbound milestones without predeccessors (Issue)
    Me.Fcount = 0                              ' 7 Clears tasks Finishing within the next 8 weeks (Info)
    Me.LTcount = 0                             ' 8 Clears long tasks (over 5 days) in the next 8 weeks (Issue)
    Me.MScount = 0                             ' 9 Clears inbound milestone with no succcessors (Issue)
    Me.TLcount = 0                             ' 10 Clears tasks over 20 days (Issue)
    Me.NScount = 0                             ' 11 Clears tasks with no successors (Issue)
    Me.NPcount = 0                             ' 12 Clears tasks with no precessors (Issue)
    Me.NFcount = 0                             ' 13 Clears tasks with negative float (Issue)
    Me.NUcount = 0                             ' 14 Clears tasks with work in the past (Issue)
    Me.IFcount = 0                             ' 15 Clears tasks with work complete in the future (Issue)
    Me.SRcount = 0                             ' 16 Clears sumary tasks with assigned resource (Issue)
    Me.MAcount = 0                             ' 17 Clears manually scheduled tasks (Issue)
    Me.HCcount = 0                             ' 18 Clears Hard Constraints (Issue)
    Me.percentComplete = 0                     ' Clear the percentage complete
    Me.issueLog = ""                           ' clear the issue log
    Me.totalIssues = 0                         ' clear the total number of issues
End Sub

Function check_Outbound(taskno As Integer)
    If ActiveProject.Tasks(taskno).Text14 = "Out" Then                                ' 3 Count outbound dependencies. Information
        Me.DOcount = Me.DOcount + 1
    End If
End Function

Function check_Inbound(taskno As Integer)
    If ActiveProject.Tasks(taskno).Text14 = "In" Then                                 ' 2 Count inbound dependencies. Information
        Me.DIcount = Me.DIcount + 1
    End If
End Function

Function count_RemainingTasks(taskno As Integer)
    If ActiveProject.Tasks(taskno).percentComplete <> 100 And ActiveProject.Tasks(taskno).Summary = False Then  ' 4 Count remaining tasks. Information
        Me.ITcount = Me.ITcount + 1
    End If
End Function

Function count_Milestones(taskno As Integer)
    If ActiveProject.Tasks(taskno).Text10 = "Yes" Then                                ' 5 Count key milestones. Information
        Me.TMcount = Me.TMcount + 1
        
    End If
End Function

Function count_tasks8weeks(taskno As Integer)
    If ActiveProject.Tasks(taskno).Finish < (ActiveProject.StatusDate + 56) And ActiveProject.Tasks(taskno).percentComplete <> 100 And ActiveProject.Tasks(taskno).Summary = False Then
        Me.Fcount = Me.Fcount + 1                                 ' 7 Count tasks finishing within 8w. Information
    End If
End Function

Function count_Outbound_withoutPred(taskno As Integer)
    If ActiveProject.Tasks(taskno).Text14 = "Out" And ActiveProject.Tasks(taskno).Predecessors = "" And ActiveProject.Tasks(taskno).percentComplete <> 100 Then
        Me.MPcount = Me.MPcount + 1                               ' 6 Count outbound milestones without predeccessors. Issue
        Me.issueLog = issueLog + "Task no " & taskno & " has an outbound milestone without a predeccessor" & vbLf
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & ". has an outbound milestone without a predeccessor"
    End If
End Function

Function count_tasks5days(taskno As Integer)
    If ActiveProject.Tasks(taskno).start < (ActiveProject.StatusDate + 56) And ActiveProject.Tasks(taskno).duration > 2400 And ActiveProject.Tasks(taskno).Summary = False And ActiveProject.Tasks(taskno).percentComplete <> 100 Then
        Me.LTcount = Me.LTcount + 1                               ' 8 Count tasks over 5 days long within next 8 weeks. Issue
        Me.issueLog = Me.issueLog + "Task no " & taskno & " is within the next 8 weeks and is more than 5 days in duration" & vbLf
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & ". is within the next 8 weeks and is more than 5 days in duration"
    End If
End Function

Function addSummaryToIssueLog()
    Me.issueLog = Me.issueLog & "----------------------------------" & vbLf
    Me.issueLog = Me.issueLog & "Analysis Complete." & vbLf
    Me.issueLog = Me.issueLog & Me.TaskCount & " tasks took " & Me.codeRunTime & " seconds to check." & vbLf
    Me.issueLog = Me.issueLog & "A total of " & Me.totalIssues & " issues were found" & vbLf
    Me.issueLog = Me.issueLog & vbLf & vbLf & "Please contact the PMO if you need any assistance to resolve these issues."
End Function

Function count_MilestonesNoSuccess(taskno As Integer)
    If ActiveProject.Tasks(taskno).Text14 = "In" And ActiveProject.Tasks(taskno).Successors = "" And ActiveProject.Tasks(taskno).percentComplete <> 100 Then
        Me.MScount = Me.MScount + 1                               ' 9 Count inbound milestones with no successors. Issue
        Me.issueLog = Me.issueLog + "Task no " & taskno & " is an inbound milestone with no successor" & vbLf
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & ". is an inbound milestone with no successor"
    End If
End Function

Function calculate()
    Me.codeFinishtime = Now()                          ' capture the finish time of the code
    Me.codeRunTime = Me.codeStartTime - Me.codeFinishtime    ' work out how long the code took to run
    
    'calculate stats
    Me.totalIssues = Me.MPcount + Me.LTcount + Me.MScount + Me.TLcount + Me.NScount + Me.NPcount + Me.NFcount + Me.NUcount + Me.IFcount + Me.SRcount + Me.MAcount + Me.HCcount ' calculate the total number of issues
End Function

Function count_TasksOver20d(taskno As Integer)
    If ActiveProject.Tasks(taskno).duration > 9600 And ActiveProject.Tasks(taskno).Summary = False And ActiveProject.Tasks(taskno).percentComplete <> 100 Then
        Me.TLcount = Me.TLcount + 1                               ' 10 Count tasks over 20d long. Issue
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Over 20d. "
        Me.issueLog = Me.issueLog + "Task no " & taskno & " is over 20 days long" & vbLf
    End If
End Function

Function check_NoSuccess(taskno As Integer)
    If ActiveProject.Tasks(taskno).Successors = "" And ActiveProject.Tasks(taskno).Summary = False And ActiveProject.Tasks(taskno).percentComplete <> 100 And ActiveProject.Tasks(taskno).ExternalTask = False Then
        Me.NScount = Me.NScount + 1
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "No Successor. "              ' 11 Count tasks with no successors. Issue
        Me.issueLog = Me.issueLog + "Task no " & taskno & " has no successors " & vbLf
    End If
End Function

Function check_NoPred(taskno As Integer)
    If ActiveProject.Tasks(taskno).Predecessors = "" And ActiveProject.Tasks(taskno).Summary = False And ActiveProject.Tasks(taskno).percentComplete <> 100 And ActiveProject.Tasks(taskno).ExternalTask = False Then
        Me.NPcount = Me.NPcount + 1                               ' 12 Count tasks with no predecessors. Issue
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "No Predecessor. "
        Me.issueLog = Me.issueLog + "Task no " & taskno & " has no predeccessors " & vbLf
    End If
End Function

Function check_NegFloat(taskno As Integer)
    If ActiveProject.Tasks(taskno).TotalSlack < 0 And ActiveProject.Tasks(taskno).Summary = False And ActiveProject.Tasks(taskno).percentComplete <> 100 Then
        Me.NFcount = Me.NFcount + 1                               ' 13 Count Tasks with negative float. Issue
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Negative Float. "
        Me.issueLog = Me.issueLog + "Task no " & taskno & " has a negative float " & vbLf
        
    End If
End Function

Function check_WorkInPast(taskno As Integer)
    If ActiveProject.Tasks(taskno).percentComplete <> 100 And ActiveProject.Tasks(taskno).Summary = False And ActiveProject.Tasks(taskno).Finish < ActiveProject.StatusDate Then
        Me.NUcount = Me.NUcount + 1                               ' 14 Count tasks with work in the past. Issue
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Incomplete in past. "
        issueLog = issueLog + "Task no " & taskno & " has work in the past " & vbLf
    End If
    If ActiveProject.Tasks(taskno).percentComplete <> 100 And ActiveProject.Tasks(taskno).Summary = False And ActiveProject.Tasks(taskno).Finish > ActiveProject.StatusDate And ActiveProject.Tasks(taskno).Resume <= ActiveProject.StatusDate Then
        Me.NUcount = Me.NUcount + 1
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Incomplete in past. "
        issueLog = issueLog + "Task no " & taskno & " has work in the past " & vbLf
    End If
    If ActiveProject.Tasks(taskno).percentComplete = 0 And ActiveProject.Tasks(taskno).Summary = False And ActiveProject.Tasks(taskno).Finish > ActiveProject.StatusDate And ActiveProject.Tasks(taskno).start < ActiveProject.StatusDate Then
        Me.NUcount = Me.NUcount + 1
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Incomplete in past. "
        Me.issueLog = Me.issueLog + "Task no " & taskno & " has work in the past " & vbLf
    End If
End Function

Function check_WorkInFuture(taskno As Integer)
    If ActiveProject.Tasks(taskno).ActualStart <> "NA" And ActiveProject.Tasks(taskno).ActualStart > ActiveProject.StatusDate Then
        Me.IFcount = Me.IFcount + 1                               ' 15 Count tasks with work complete in the future. Issue
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Complete in future. "
        Me.issueLog = Me.issueLog + "Task no " & taskno & " has work completed in the future " & vbLf
    End If
End Function

Function check_SummaryResources(taskno As Integer)
     If ActiveProject.Tasks(taskno).Summary = "True" And ActiveProject.Tasks(taskno).percentComplete <> 100 Then
        If Not ActiveProject.Tasks(taskno).ResourceNames = "" Then                        ' 16 Count Summary resources with resources assigned
            Me.SRcount = Me.SRcount + 1
            ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Summary resourced. "
            Me.issueLog = Me.issueLog + "Task no " & taskno & " is a summary task with resources assigned " & vbLf
        End If
    End If
End Function

Function check_ManuallyScheduled(taskno As Integer)
    If ActiveProject.Tasks(taskno).Manual = "True" And ActiveProject.Tasks(taskno).percentComplete <> 100 And ActiveProject.Tasks(taskno).ExternalTask = False Then
        Me.MAcount = Me.MAcount + 1                               ' 17 Count manually scheduled tasks. issue
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Manually Scheduled. "
        Me.issueLog = Me.issueLog + "Task no " & taskno & " is manually assigned " & vbLf
    End If
End Function

Function check_HardConstraints(taskno As Integer)
    If ActiveProject.Tasks(taskno).ConstraintType <> 0 And ActiveProject.Tasks(taskno).percentComplete <> 100 And ActiveProject.Tasks(taskno).ExternalTask = False Then
        Me.HCcount = Me.HCcount + 1                               ' 18 Count hard constraints. Issue
        ActiveProject.Tasks(taskno).Text25 = ActiveProject.Tasks(taskno).Text25 & "Constrained. "
        Me.issueLog = Me.issueLog + "Task no " & taskno & " has hard constraints " & vbLf
    End If
End Function
