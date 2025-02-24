Attribute VB_Name = "LSEG_code"
Option Explicit

'File names and path location
Const INPUT_LOG_PATH = "C:\Users\junis\OneDrive\Documents\Job Search\LSEG - code test\"
Const INPUT_LOG_FILE = "logs.log"
Const INPUT_LOG_FILE_TEST = "logs - TEST.log"
Const OUTPUT_LOG_FILE = "output.log"

'Variables to set thresholds for WARNINGs and ERRORs
Const WARNING_THRESHOLD As Date = "00:05"
Const ERROR_THRESHOLD As Date = "00:10"


'Constants for locations of data in the in input CSV file
Const TIMESTAMP As Byte = 0
Const TASK As Byte = 1
Const ACTIONNAME As Byte = 2
Const ID As Byte = 3

Sub Main_Procedure()

    Dim Log_Entries     As Collection
    Dim oFileSystem     As Object
    Dim oFile           As Object
    
    'Read the input file
    Set Log_Entries = Parse_Log_File(INPUT_LOG_PATH, INPUT_LOG_FILE)
    
    'Create an obejct to write out the results
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFileSystem.CreateTextFile(INPUT_LOG_PATH & OUTPUT_LOG_FILE)
    
    'Write out the results to an output file
    Call WriteToFile("ERROR", Log_Entries, oFile)
    Call WriteToFile("WARNING", Log_Entries, oFile)
    
    'Close the file and clear the objects
    oFile.Close
    Set oFileSystem = Nothing
    Set oFile = Nothing

End Sub

'Test function that uses explict file to validate processes
Function Test_File_Read() As Boolean
    
    Dim Log_Entries As Collection
    
    'Call the main function and print the results
    Set Log_Entries = Parse_Log_File(INPUT_LOG_PATH, INPUT_LOG_FILE_TEST)
    Debug.Print "TEST NO PROBLEM name", Log_Entries(1).TaskName = "TEST NO PROBLEM"
    Debug.Print "TEST WARNING name", Log_Entries(2).TaskName = "TEST WARNING"
    Debug.Print "TEST WARNING time", Log_Entries(2).ProcessTime > WARNING_THRESHOLD
    Debug.Print "TEST ERROR name", Log_Entries(3).TaskName = "TEST ERROR"
    Debug.Print "TEST ERROR time", Log_Entries(3).ProcessTime > ERROR_THRESHOLD
    Debug.Print "TEST NO END name", Log_Entries(4).TaskName = "TEST NO END"
    
    'Cause an error if unexpected data
    If Not Log_Entries(2).ProcessTime > WARNING_THRESHOLD Then Exit Function
    If Not Log_Entries(3).ProcessTime > ERROR_THRESHOLD Then Exit Function
    If Not Log_Entries(1).TaskName = "TEST NO PROBLEM" Then Exit Function
    If Not Log_Entries(2).TaskName = "TEST WARNING" Then Exit Function
    If Not Log_Entries(2).ProcessTime > WARNING_THRESHOLD Then Exit Function
    If Not Log_Entries(3).TaskName = "TEST ERROR" Then Exit Function
    If Not Log_Entries(3).ProcessTime > ERROR_THRESHOLD Then Exit Function
    If Not Log_Entries(4).TaskName = "TEST NO END" Then Exit Function
    
    'Exit function as a success
    Test_File_Read = True
End Function

Function Parse_Log_File(szPath As String, szInputFile As String) As Collection
        
    Dim FileNum         As Integer
    Dim DataLine        As String
    Dim szLines()       As String
    Dim vLine           As Variant
    
    Dim Log_Entry       As Log_Entry
    Dim Log_Entries     As Collection
    Dim vValues         As Variant
    Dim bNewObj         As Boolean
    
    'Open the input file for reading
    FileNum = FreeFile()
    Open szPath & szInputFile For Input As #FileNum
    
    'The file given reads the whole file in a single entry not line by line, so workaround is to load the whole file into an array
    Line Input #FileNum, DataLine
    szLines = Split(DataLine, Chr$(10))
    Set Log_Entries = New Collection
    
    'Process each line of the file, adding that data to the Log_Entries collection
    For Each vLine In szLines
        If vLine <> "" Then                                                                     'Check we have a line with data
            vValues = Split(vLine, ",")                                                             'Parse the line into constituent parts
            Set Log_Entry = GetLogEntryObject(vValues(TASK), Log_Entries, bNewObj)                  'Call function to either return correct obeject or create new one
            If Trim(vValues(ACTIONNAME)) = "START" Then Log_Entry.StartTime = vValues(TIMESTAMP)    'Add the start time, if that was the data
            If Trim(vValues(ACTIONNAME)) = "END" Then Log_Entry.EndTime = vValues(TIMESTAMP)        'Add the end time, if that was the data
            Log_Entry.ID = vValues(ID)                                                              'Add the ID of the process
            If bNewObj Then Log_Entries.Add Log_Entry                                               'Only add this object to the collection, if it is actually new
        End If
    Next vLine
    Set Parse_Log_File = Log_Entries
End Function

'Function to either retrieve Log_Entry object, or create a new one
Function GetLogEntryObject(ByVal szLogName As String, Log_Entries As Collection, NewObj As Boolean) As Log_Entry
    Dim objLogEntry As Log_Entry
    For Each objLogEntry In Log_Entries             'Check the 'TaskName' property in each object of the 'Log_Entries' collection to see if this is a new object, or an existing object
        If objLogEntry.TaskName = szLogName Then
            NewObj = False                          'If we found this object, just return it and set the ByRef NewObj flag to false
            Set GetLogEntryObject = objLogEntry
            Exit Function                           'Jump out of the object without creating a new one, but just passing back the correct object
        End If
    Next objLogEntry
    NewObj = True                                   'If we checked the whole 'Log_Entries' collection and did not find this object, then create a new one and pass it back
    Set objLogEntry = New Log_Entry                 'Create a new object
    objLogEntry.TaskName = szLogName                'Add the name of the process
    Set GetLogEntryObject = objLogEntry             'Return this object in this function
End Function

'Function to handle writing out to a file
Function WriteToFile(szDataToOutput As String, Log_Entries As Collection, oFile As Object)
Dim Log_Entry As Log_Entry

'Write out the data requested if they were ERRORs
If szDataToOutput = "ERROR" Then
    For Each Log_Entry In Log_Entries
        If Log_Entry.ProcessTime = -1 Then oFile.WriteLine Log_Entry.TaskName & " ERROR"                'Treat incomplete entries as an ERROR
        If Log_Entry.ProcessTime > ERROR_THRESHOLD Then oFile.WriteLine Log_Entry.TaskName & " ERROR"
    Next Log_Entry
End If

'Write out the data requested if they were WARNINGs
If szDataToOutput = "WARNING" Then
    For Each Log_Entry In Log_Entries
        If Log_Entry.ProcessTime > WARNING_THRESHOLD And Log_Entry.ProcessTime <= ERROR_THRESHOLD Then oFile.WriteLine Log_Entry.TaskName & " WARNING"
    Next Log_Entry
End If

End Function




