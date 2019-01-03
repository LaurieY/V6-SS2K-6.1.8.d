VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scope"
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
'   SCOPE.CLS
'   =============
'
' Implementation of class for storing and retrieving scope settings.
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
Private Const DEFAULT_SLEWSETTLETIME As Integer = 5

'Registry location
Private Const REGSEC_SCOPE As String = "Scope"
Private Const REGVAL_SLEWSETTLETIME As String = "SlewSettleTime"

Private m_nSlewSettleTime As Integer

Property Let SlewSettleTime(nSlewSettleTime As Integer)
    m_nSlewSettleTime = nSlewSettleTime
End Property

Property Get SlewSettleTime() As Integer
    SlewSettleTime = m_nSlewSettleTime
End Property

Public Sub SaveSettings()
    g_Profile.WriteValue g_ID, REGVAL_SLEWSETTLETIME, CStr(m_nSlewSettleTime), REGSEC_SCOPE
End Sub

Public Sub LoadSettings()
    Dim buf As String
    
'Allow empty value or missing key in the registry
On Error Resume Next
    
    buf = g_Profile.GetValue(g_ID, REGVAL_SLEWSETTLETIME, REGSEC_SCOPE)
    If buf <> "" Then
        m_nSlewSettleTime = CInt(buf)
    Else
        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_SLEWSETTLETIME, CStr(DEFAULT_SLEWSETTLETIME), REGSEC_SCOPE
    End If
End Sub

Private Sub Class_Initialize()
   'Initalize members with default values
    m_nSlewSettleTime = DEFAULT_SLEWSETTLETIME
        
    'Override with registry settings
    LoadSettings
End Sub

