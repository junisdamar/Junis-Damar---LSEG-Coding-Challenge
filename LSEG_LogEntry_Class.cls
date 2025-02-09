VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LSEG_LogEntry_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StartTime    As Date
Public EndTime      As Date
Public TaskName     As String
Public ID           As Long

Property Get ProcessTime() As Date
    If EndTime > 0 And StartTime > 0 Then
        ProcessTime = EndTime - StartTime
    Else
        ProcessTime = -1
    End If
End Property

