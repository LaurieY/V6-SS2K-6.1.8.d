VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ROMVersion"
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
'   ROMVERSION.CLS
'   =============
'
' Implementation of class for storing and retrieving ROM version.
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
Private Const DEFAULT_VERSION As Integer = 201

'Registry location
Private Const REGSEC_ROM As String = "ROM"
Private Const REGVAL_VERSION As String = "Version"

Private m_nVersion As Integer

Property Let Version(nVersion As Integer)
    m_nVersion = nVersion
End Property

Property Get Version() As Integer
    Version = m_nVersion
End Property

Public Sub SaveSettings()
    g_Profile.WriteValue g_ID, REGVAL_VERSION, CStr(m_nVersion), REGSEC_ROM
End Sub

Public Sub LoadSettings()
    Dim buf As String
    
'Allow empty value or missing key in the registry
On Error Resume Next
    
    buf = g_Profile.GetValue(g_ID, REGVAL_VERSION, REGSEC_ROM)
    If buf <> "" Then
        m_nVersion = CInt(buf)
    Else
        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_VERSION, CStr(DEFAULT_VERSION), REGSEC_ROM
    End If
End Sub

Private Sub Class_Initialize()
    'Initialize members with default values
    m_nVersion = DEFAULT_VERSION
    
    'Override with registry settings
    LoadSettings
End Sub
