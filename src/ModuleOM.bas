Attribute VB_Name = "ModuleOM"
'This module contains a macro which will display
'QA info in a message box
'Created by Sean Boyle, SDB Projects Ltd, for the British Council September 2014
Option Explicit


Dim MNcount As Integer 'milestones without preds
Dim DIcount As Integer 'dep ins
Dim DOcount As Integer 'dep outs
Dim ITcount As Integer 'Incomplete tasks
Dim LTcount As Integer 'Over 5d in 8w
Dim Fcount As Integer  'finishes in 8w count
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
Dim SRcount As Integer 'summarty tasks with resources
Dim MAcount As Integer 'Manually scheduled tasks
Dim HCcount As Integer 'Hard constraints
Dim ps As String       'project status
Dim SD As String       'status date
Dim SU As String       'unformatted status date
Dim FD As Date         'project finish date
Dim BF As String       'project baseline finish
Dim BU As String       'unformatted last baselined date
Dim BD As String       'last baselined date
Dim fso As FileSystemObject
Dim proj As Project
Dim t As Task



Sub QA_Chex()

SU = ActiveProject.StatusDate
If SU = "NA" Then
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



'Clear issue field (Needed if errors written to Text25)
For Each t In ActiveProject.Tasks
    If Not t Is Nothing Then
        t.Text25 = ""
    End If
Next t

'Count inbound dependencies. Information
DIcount = 0
For Each t In ActiveProject.Tasks
    If Not t Is Nothing Then
        If t.Text14 = "In" Then
            DIcount = DIcount + 1
        End If
    End If
Next t


'Count outbound dependencies. Information
DOcount = 0
For Each t In ActiveProject.Tasks
    If Not t Is Nothing Then
        If t.Text14 = "Out" Then
            DOcount = DOcount + 1
        End If
    End If
Next t


'Count remaining tasks. Information
ITcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.percentComplete <> 100 And t.Summary = False Then
        ITcount = ITcount + 1
     End If
End If
Next t


' Count key milestones. Information
TMcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Text10 = "Yes" Then
            TMcount = TMcount + 1
    End If
End If
Next t

'Count summary task predecessors. Issue
SLcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Summary = "True" And t.percentComplete <> 100 Then
        If Not t.Predecessors = "" Then
            SLcount = SLcount + 1
            t.Text25 = t.Text25 & "Summary predeceessor. "
        End If
    End If
End If
Next t

 'Count summary task successors. Issue
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Summary = "True" And t.percentComplete <> 100 Then
        If Not t.Successors = "" Then
            SLcount = SLcount + 1
            t.Text25 = t.Text25 & "Summary successor. "
        End If
    End If
End If
Next t


'Count summary task with assigned resource. Issue
SRcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Summary = "True" And t.percentComplete <> 100 Then
        If Not t.ResourceNames = "" Then
            SRcount = SRcount + 1
            t.Text25 = t.Text25 & "Summary resourced. "
        End If
    End If
End If
Next t

'Count outbound milestones without predeccessors. Issue
MPcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Text14 = "Out" And t.Predecessors = "" And t.percentComplete <> 100 Then
        MPcount = MPcount + 1
    End If
End If
Next t


'Count tasks finishing within 8w. Information
Fcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Finish < (ActiveProject.StatusDate + 56) And t.percentComplete <> 100 And t.Summary = False Then
        Fcount = Fcount + 1
    End If
End If
Next t


'Count tasks over 5 days long within next 8 weeks. Issue
LTcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Start < (ActiveProject.StatusDate + 56) And t.Duration > 2400 And t.Summary = False And t.percentComplete <> 100 Then
        LTcount = LTcount + 1
    End If
End If
Next t


'Count inbound milestones with no successors. Issue
MScount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Text14 = "In" And t.Successors = "" And t.percentComplete <> 100 Then
        MScount = MScount + 1
    End If
End If
Next t


'Count tasks over 20d long. Issue
TLcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Duration > 9600 And t.Summary = False And t.percentComplete <> 100 Then
        TLcount = TLcount + 1
        t.Text25 = t.Text25 & "Over 20d. "
    End If
End If
Next t


'Count tasks with no successors. Issue
NScount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Successors = "" And t.Summary = False And t.percentComplete <> 100 And t.ExternalTask = False Then
        NScount = NScount + 1
        t.Text25 = t.Text25 & "No Successor. "
    End If
End If
Next t


'Count tasks with no predecessors. Issue
NPcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.Predecessors = "" And t.Summary = False And t.percentComplete <> 100 And t.ExternalTask = False Then
        NPcount = NPcount + 1
        t.Text25 = t.Text25 & "No Predecessor. "
    End If
End If
Next t



'Count Tasks with negative float. Issue
NFcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.TotalSlack < 0 And t.Summary = False And t.percentComplete <> 100 Then
        NFcount = NFcount + 1
        t.Text25 = t.Text25 & "Negative Float. "
    End If
End If
Next t


'Count tasks with work in the past. Issue
NUcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.percentComplete <> 100 And t.Summary = False And t.Finish < ActiveProject.StatusDate Then
        NUcount = NUcount + 1
        t.Text25 = t.Text25 & "Incomplete in past. "
    End If
End If
Next t
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    
    If t.percentComplete <> 100 And t.Summary = False And t.Finish > ActiveProject.StatusDate And t.Resume <= ActiveProject.StatusDate Then
        NUcount = NUcount + 1
        t.Text25 = t.Text25 & "Incomplete in past. "
    End If
End If
Next t
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    
    If t.percentComplete = 0 And t.Summary = False And t.Finish > ActiveProject.StatusDate And t.Start < ActiveProject.StatusDate Then
        NUcount = NUcount + 1
        t.Text25 = t.Text25 & "Incomplete in past. "
    End If
End If
Next t


'Count tasks with work complete in the future. Issue
IFcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.ActualStart <> "NA" And t.ActualStart > ActiveProject.StatusDate And t.Summary = False Then
        IFcount = IFcount + 1
        t.Text25 = t.Text25 & "Complete in future. "
    End If
End If
Next t

'Count manually scheduled tasks. issue
MAcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
If t.Manual = "True" And t.percentComplete <> 100 And t.ExternalTask = False Then
        MAcount = MAcount + 1
        t.Text25 = t.Text25 & "Manually Scheduled. "
    End If
End If
Next t

'count hard constraints
HCcount = 0
For Each t In ActiveProject.Tasks
If Not t Is Nothing Then
    If t.ConstraintType = 2 Or t.ConstraintType = 3 Or t.ConstraintType = 5 Or t.ConstraintType = 7 Then
        If t.percentComplete <> 100 Then
            HCcount = HCcount + 1
            t.Text25 = t.Text25 & "Hard Constraint. "
        End If
    End If
End If
Next t


    MsgBox (ActiveProject.Name & Chr$(10) _
& Chr$(10) & "Its current Status Date is " & SD & Chr$(10) _
& "Its current Finish Date is " & FD & Chr$(10) _
& "Its current Baseline Finish Date is " & BF & Chr$(10) _
& "Its current Baseline was set on " & BD & Chr$(10) _
& "It has " & ITcount & " remaining tasks" & Chr$(10) _
& Chr$(10) & "It has " & IFcount + NFcount + MScount + MNcount + MAcount + SRcount + HCcount + MPcount + NUcount + SLcount + TLcount + NScount + NPcount _
& " issues in the following areas..." _
& Chr$(10) & Chr$(10) & NUcount & " incomplete tasks in the past." _
& Chr$(10) & IFcount & " started tasks in the future." _
& Chr$(10) & MScount & " inbound dependencies have no successors." _
& Chr$(10) & MPcount & " outbound dependencies have no predecessors." _
& Chr$(10) & NFcount & " tasks have negative float." _
& Chr$(10) & TLcount & " tasks have durations greater than 20 days." _
& Chr$(10) & NScount & " tasks have no successors." _
& Chr$(10) & NPcount & " tasks have no predecessors." _
& Chr$(10) & SLcount & " links on summary tasks." _
& Chr$(10) & SRcount & " resourced summary tasks." & Chr$(10) & MAcount & " manually scheduled activities." & Chr$(10) & HCcount & " hard constraints." _
& Chr$(10) & Chr(10) & "Also... " _
& Chr$(10) & Fcount & " finishes in next 8w." _
& Chr$(10) & TMcount & " key milestones present." _
& Chr$(10) & DIcount & " inbound dependencies." _
& Chr$(10) & DOcount & " outbound dependencies." _
& Chr$(10) & Chr$(10) & "Please contact the PMO if you need" _
& Chr$(10) & "any assistance to resolve these issues."), vbOKOnly, "British Council PMO - Plan Quality Dashboard"


If MsgBox("Do you wish to print the Dashboard report?", 4, "British Council PMO - Plan Quality Dashboard") = vbYes Then

'code to create text file
'Create a FileSystemObject.
Set fso = New FileSystemObject
' Declare a TextStream.
Dim stream As TextStream

'Create a TextStream.
Set stream = fso.CreateTextFile("h:\MSP_Issues.log", True)
stream.WriteLine
stream.WriteLine
stream.WriteLine "Project Name: " & ActiveProject.Name
stream.WriteLine
stream.WriteLine "Current Status Date: " & SD
stream.WriteLine "Current Finish Date: " & FD
stream.WriteLine "Current Baseline Finish Date: " & BF
stream.WriteLine "Current Baseline was set on " & BD
stream.WriteLine
stream.WriteLine ActiveProject.Name & " has " & ITcount & " remaining tasks"
stream.WriteLine
stream.WriteLine "It has " & IFcount + NFcount + MScount + MNcount + SRcount + MPcount + NUcount + SLcount + LTcount + TLcount + NScount + NPcount _
& " issues in the following areas..."
stream.WriteLine
stream.WriteLine NUcount & " incomplete tasks in the past."
stream.WriteLine IFcount & " started tasks in the future."
stream.WriteLine MScount & " inbound dependencies have no successors."
stream.WriteLine MPcount & " outbound dependencies have no predecessors."
stream.WriteLine NFcount & " tasks have negative float."
stream.WriteLine LTcount & " tasks over 5d within 8w."
stream.WriteLine TLcount & " tasks have durations greater than 20 days."
stream.WriteLine NScount & " tasks have no successors."
stream.WriteLine NPcount & " tasks have no predecessors."
stream.WriteLine SLcount & " links on summary tasks."
stream.WriteLine SRcount & " resourced summary tasks."
stream.WriteLine HCcount & " resourced summary tasks."
stream.WriteLine
stream.WriteLine "Other information... "
stream.WriteLine
stream.WriteLine Fcount & " finishes in next 8w."
stream.WriteLine TMcount & " key milestones present."
stream.WriteLine DIcount & " inbound dependencies."
stream.WriteLine DOcount & " outbound dependencies."
stream.WriteLine
stream.WriteLine "Please contact the PMO if you need"
stream.WriteLine "any assistance to resolve these issues."

' Close the file.
stream.Close

'Code to print text file
Shell ("c:\windows\notepad.exe /p h:\MSP_Issues.log")

End If
Err:
End Sub

