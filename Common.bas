Attribute VB_Name = "Common"
'   ============
'   Common.bas
'   ============
'
' Common utility functions
'
' 17 March 2008 Laurie Yates
Option Explicit





Private m_previousRA
Private m_previousDEC
'Private m_Serial As DriverHelper.Serial
Private m_bActive As Boolean        ' set while sending and receiving a command
Private m_preMoveRA, m_premoveDEC, m_movingRA, m_movingDEC As Boolean
Private FSO As Scripting.FileSystemObject
'Public g_AxisRates As AxisRates         ' rates available for MoveAxis
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)




Public Sub Wait()
    ' wait for any send/receive process to finish before sending another
    Static Count As Integer
    Count = 0
    While m_bActive
        Sleep 100
        DoEvents
        Count = Count + 1
        If Count > 100 Then m_bActive = False   ' 10 second timeout
    Wend
End Sub





