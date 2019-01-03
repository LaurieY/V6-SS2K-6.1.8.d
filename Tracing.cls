VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tracing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_bTracing As Boolean
'Private m_ologging As AppMessages.Logging  ' for stand-alone debug
Private m_ologging As Logging

'Public g_AxisRates As AxisRates         ' rates available for MoveAxis
'Private Declare Function GetTickCount Lib "kernel32" () As Long
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Class_Initialize()
       Dim buf
       Dim tempPath As String
       On Error Resume Next
    m_bTracing = False   'default
       'Dim m_logPath, m_logMode
       'm_logPath = App.LogPath
       'm_logMode = App.LogMode
    buf = g_Profile.GetValue(g_ID, REGVAL_TRACE, REGSEC_TRACE)
     If buf <> "" Then
        m_bTracing = buf  ' 0 is false 1  is true
    
    End If
  
 

    'App.LogEvent (" Tracing Started")
     'Set m_ologging = New AppMessages.Logging  ' For sdebug logging
     If m_bTracing Then
     Set m_ologging = New Logging
     tempPath = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp"))

           'm_ologging.StartLog "C:\TEMP\SS2K.log", mLogToFile
           m_ologging.StartLog tempPath + "\ss2k.log", mLogToFile
           m_ologging.AddLogEvent "---------Tracing Started " & Format$(Now(), "General Date") & " -----------------", vbLogEventTypeInformation
         
    End If
    
End Sub
Public Sub Trace(msg As String)
    If m_bTracing Then
        'm_TraceFile.WriteLine "[" & Format$(Now(), "Hh:Nn:Ss") & "] " & msg
        'App.LogEvent (msg)
         m_ologging.AddLogEvent "[" & Format$(Now(), "Hh:Nn:Ss") & "] " & msg, vbLogEventTypeInformation
    End If
End Sub
Public Property Let SetTracing(ByVal newVal As Boolean)
         g_Profile.WriteValue g_ID, REGVAL_TRACE, newVal, REGSEC_TRACE

End Property
