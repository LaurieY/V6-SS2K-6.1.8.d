VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum eLogModes
          mLogAuto = 0
          mLogOff = 1
          mLogToFile = 2
          mLogToNT = 3
          mLogOverwrite = &H10
          mLogThreadID = &H20
      End Enum

      Public Sub StartLog(logTarget As String, logMode As eLogModes)
          App.StartLogging logTarget, logMode

      End Sub

      Public Sub AddLogEvent(LogBuffer As String, EventType As _
      LogEventTypeConstants)
          App.LogEvent LogBuffer, EventType
      End Sub


