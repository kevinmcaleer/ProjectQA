VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is the Check Class

Private name As String
Public checkPredecessors As clsCheck


Public Property Set name(s As String)
    name = s
End Property

Public Property Let name(s As String)
    s = name
End Property

'need code here to implement the class.runTest() code
