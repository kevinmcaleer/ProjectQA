VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private name As String
Private description As String
Private state As Boolean
'Private count As Integer

Public Property Set count(c As Integer)
    count = c
End Property

Public Property Let count(c As Integer)
    c = count
End Property

'Run the test
Public Sub runTest()

End Sub
