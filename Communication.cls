VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Communication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------
' Copyright © 2000-2002 SPACE.com Inc., New York, NY
'
' Permission is hereby granted to use this Software for any purpose
' including combining with commercial products, creating derivative
' works, and redistribution of source or binary code, without
' limitation or consideration. Any redistributed copies of this
' Software must include the above Copyright Notice.
'
' THIS SOFTWARE IS PROVIDED "AS IS". SPACE.COM, INC. MAKES NO
' WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
' SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'---------------------------------------------------------------------
'   =============
'   COMMUNICATION.CLS
'   =============
'
' Implementation of class for storing and retrieving communication settings.
'
' Written:  04-Jan-01   Arne Danielsen
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 04-Jan-01 ad      Initial edit
' 24-Jan-01 ad      Updated code to confirm with Beta 2 of the SDK
' 30-Dec-01 ad      Create registry settings programmatically if they
'                   don't exist allready.
'---------------------------------------------------------------------

Option Explicit

'Default values
Private Const DEFAULT_PORT As Integer = 1
Private Const DEFAULT_SPEED As Integer = 9600

'Registry location
Private Const REGSEC_COMMUNICATION As String = "Comunication"
Private Const REGVAL_PORT As String = "Port"
Private Const REGVAL_SPEED As String = "Speed"

Private m_nPort As Integer
Private m_nSpeed As Integer

Property Let Port(nPort As Integer)
    m_nPort = nPort
End Property

Property Get Port() As Integer
    Port = m_nPort
End Property

Property Let Speed(nSpeed As Integer)
    m_nSpeed = nSpeed
End Property

Property Get Speed() As Integer
    Speed = m_nSpeed
End Property

Public Sub SaveSettings()
    g_Profile.WriteValue g_ID, REGVAL_PORT, CStr(m_nPort), REGSEC_COMMUNICATION
    g_Profile.WriteValue g_ID, REGVAL_SPEED, CStr(m_nSpeed), REGSEC_COMMUNICATION
End Sub

Public Sub LoadSettings()
    Dim buf As String
    
'Allow empty value or missing key in the registry
On Error Resume Next
    
    buf = g_Profile.GetValue(g_ID, REGVAL_PORT, REGSEC_COMMUNICATION)
    If buf <> "" Then
        m_nPort = CInt(buf)
    Else
        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_PORT, CStr(DEFAULT_PORT), REGSEC_COMMUNICATION
    End If
    
    buf = g_Profile.GetValue(g_ID, REGVAL_SPEED, REGSEC_COMMUNICATION)
    If buf <> "" Then
        m_nSpeed = CInt(buf)
    Else
        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_SPEED, CStr(DEFAULT_SPEED), REGSEC_COMMUNICATION
    End If
End Sub

Private Sub Class_Initialize()
    'Initialize members with default values.
    m_nPort = DEFAULT_PORT
    m_nSpeed = DEFAULT_SPEED
    
    'Override with registry settings
    LoadSettings
End Sub
