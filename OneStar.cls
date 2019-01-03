VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OneStar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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


Option Explicit

'Default values
Private Const DEFAULT_ONESTAR As Boolean = False

'Registry location
Private Const REGSEC_STAR As String = "OneStar"
Private Const REGVAL_ONESTAR As String = "Alignments"

Private m_bOneStar As Boolean

Property Let AStar(bOneStar As Boolean)
    m_bOneStar = bOneStar
End Property

Property Get AStar() As Boolean
    AStar = m_bOneStar
End Property

Public Sub SaveSettings()
g_Tracing.Trace ("saving OneStar settings  6.1.7d_trace")
    g_Profile.WriteValue g_ID, REGVAL_ONESTAR, m_bOneStar, REGSEC_STAR
End Sub

Public Sub LoadSettings()
    Dim buf As String
 g_Tracing.Trace ("Loading OneStar settings  6.1.7d_trace")
   
'Allow empty value or missing key in the registry
On Error Resume Next
    
    buf = g_Profile.GetValue(g_ID, REGVAL_ONESTAR, REGSEC_STAR)
    If buf <> "" Then
        m_bOneStar = CBool(buf)
    Else
    g_Tracing.Trace ("saving OneStar defaults  6.1.7d_trace")

        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_ONESTAR, DEFAULT_ONESTAR, REGSEC_STAR
    End If
End Sub

Private Sub Class_Initialize()
    'Initialize members with default values
    m_bOneStar = DEFAULT_ONESTAR
    
    'Override with registry settings
    LoadSettings
End Sub

