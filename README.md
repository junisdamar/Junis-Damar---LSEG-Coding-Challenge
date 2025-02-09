# Junis-Damar---LSEG-Coding-Challenge

PURPOSE:
This is a coding challeng from LSEG completed by Junis Damar

CODE FILES:
There are 2 code files: 

  LSEG_code.bas 
  This file has the VB code with 2 functions and a testing subroutine
  
  LSEG_LogEntry_Class.cls 
  This is a small class file to hold the log entry instances

LOGIC:

  The main function is 'Parse_Log_File', which reads a file set by the INPUT_LOG_FILE variable and parses it.
  The process is effectivly in 2 parts, firstly reading the log file, then writing an output
  The class file is just a simple way to hold instances of each process, so we can track the start and end times, which appear on different rows of the input file

IMPROVMENTS:

  There is no error handling around the format of the input file, so code is not resilient to misformated inputs
  There is no logic to handle the output file being locked
