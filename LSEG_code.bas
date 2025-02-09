Attribute VB_Name = "LSEG_code"
Option Explicit

Const INPUT_LOG_PATH = "C:\Users\junis\OneDrive\Documents\Job Search\LSEG - code test\"
Const INPUT_LOG_FILE = "logs.log"
Const OUTPUT_LOG_FILE = "output.log"

Const HANDLED_ERROR As Long = 65535

Const TIMESTAMP As Byte = 0
Const TASK As Byte = 1
Const ACTIONNAME As Byte = 2
Const ID As Byte = 3

Const WARNING_THRESHOLD As Date = "00:05"
Const ERROR_THRESHOLD As Date = "00:10"

Sub Test_App()
    Call Parse_Log_File(INPUT_LOG_PATH, INPUT_LOG_FILE, OUTPUT_LOG_FILE)
End Sub

Function Parse_Log_File(szPath As String, szInputFile As String, szOutputFile As String)
    
Dim FileNum         As Integer
Dim DataLine        As String
Dim szLines()       As String
Dim vLine           As Variant

Dim Log_Entry       As Log_Entry
Dim Log_Entries     As Collection
Dim vValues         As Variant
Dim bNewObj         As Boolean

Dim oFileSystem     As Object
Dim oFile           As Object

FileNum = FreeFile()
Open szPath & szInputFile For Input As #FileNum

' The example file reads in as a single file, not line by line, so workaround is to load the whole file into an array
Line Input #FileNum, DataLine
szLines = Split(DataLine, Chr$(10))
Set Log_Entries = New Collection

'Process each line of the file, adding that data to the Log_Entries collection
For Each vLine In szLines
    'Stop
    If vLine <> "" Then                                                                     'Check we have a line with data
        vValues = Split(vLine, ",")                                                             'Parse the line into constituent parts
        Set Log_Entry = GetLogEntryObject(vValues(TASK), Log_Entries, bNewObj)                  'Call function to either return correct obeject or create new one
        If Trim(vValues(ACTIONNAME)) = "START" Then Log_Entry.StartTime = vValues(TIMESTAMP)    'Add the start time, if that was the data
        If Trim(vValues(ACTIONNAME)) = "END" Then Log_Entry.EndTime = vValues(TIMESTAMP)        'Add the end time, if that was the data
        Log_Entry.ID = vValues(ID)
        If bNewObj Then Log_Entries.Add Log_Entry
    End If
Next vLine
Stop

'Write out the results to an output file
Set oFileSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oFileSystem.CreateTextFile(szPath & szOutputFile)

'Write out the ERRORs
For Each Log_Entry In Log_Entries
    If Log_Entry.ProcessTime = -1 Then
        oFile.WriteLine Log_Entry.TaskName & " ERROR"
    ElseIf Log_Entry.ProcessTime > ERROR_THRESHOLD Then
        oFile.WriteLine Log_Entry.TaskName & " ERROR"
    End If
Next Log_Entry

'Write out the WARNINGs
For Each Log_Entry In Log_Entries
    If Log_Entry.ProcessTime = -1 Then
        'Do Nothing
    ElseIf Log_Entry.ProcessTime > ERROR_THRESHOLD Then
        'Do Nothing
    ElseIf Log_Entry.ProcessTime > WARNING_THRESHOLD Then
        oFile.WriteLine Log_Entry.TaskName & " WARNING"
    End If
Next Log_Entry


oFile.Close

Set oFileSystem = Nothing
Set oFile = Nothing

End Function

'Function to either retrieve Log_Entry object, or create a new one
Function GetLogEntryObject(ByVal szLogName As String, Log_Entries As Collection, NewObj As Boolean) As Log_Entry
    Dim objLogEntry As Log_Entry
    For Each objLogEntry In Log_Entries             'Check the 'TaskName' property in each object of the 'Log_Entries' collection to see if this is a new object, or an existing object
        If objLogEntry.TaskName = szLogName Then
            NewObj = False                          'If we found this object, just return it and set the ByRef NewObj flag to false
            Set GetLogEntryObject = objLogEntry
            Exit Function
        End If
    Next objLogEntry
    NewObj = True                                   'If we checked the whole 'Log_Entries' collection and did not find this object, then create a new one and pass it back
    Set objLogEntry = New Log_Entry
    objLogEntry.TaskName = szLogName
    Set GetLogEntryObject = objLogEntry
End Function
