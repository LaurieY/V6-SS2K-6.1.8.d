VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------
' Copyright © 2001, Arne Danielsen, Norway
'
' Permission is hereby granted to use this source code for non-
' commercial purposes only. No warranties are made regarding
' this source code, implied or express, as to the suitability or
' fitness for a particular purpose.
'---------------------------------------------------------------------
'   =============
'   LOCATION.CLS
'   =============
'
' Implementation of class for storing and retrieving location settings.
'
' Written:  05-Jan-01   Arne Danielsen
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 05-Jan-01 ada     Initial edit
'---------------------------------------------------------------------

Option Explicit

'Default values
Private Const DEFAULT_LATITUDE As Double = 0
Private Const DEFAULT_LONGITUDE As Double = 0
Private Const DEFAULT_ELEVATION As Integer = 0

'Registry location
Private Const REGSEC_GENERAL As String = "General"
Private Const REGVAL_LATITUDE As String = "Latitude"
Private Const REGVAL_LONGITUDE As String = "Longitude"
Private Const REGVAL_ELEVATION As String = "Elevation"

Private m_dLatitude As Double
Private m_dLongitude As Double
Private m_nElevation As Integer

Property Let Latitude(dLatitude As Double)
    m_dLatitude = dLatitude
End Property

Property Get Latitude() As Double
    Latitude = m_dLatitude
End Property

Property Let Longitude(dLongitude As Double)
    m_dLongitude = dLongitude
End Property

Property Get Longitude() As Double
    Longitude = m_dLongitude
End Property

Property Let Elevation(nElevation As Integer)
    m_nElevation = nElevation
End Property

Property Get Elevation() As Integer
    Elevation = m_nElevation
End Property

Public Sub SaveSettings()
    WriteProfile REGSEC_GENERAL, REGVAL_LATITUDE, CStr(m_dLatitude)
    WriteProfile REGSEC_GENERAL, REGVAL_LONGITUDE, CStr(m_dLongitude)
    WriteProfile REGSEC_GENERAL, REGVAL_ELEVATION, CStr(m_nElevation)
End Sub

Public Sub LoadSettings()
    Dim buf As String
    
    buf = GetProfile(REGSEC_GENERAL, REGVAL_LATITUDE)
    If buf <> "" Then
        m_dLatitude = CDbl(buf)
    End If
    
    buf = GetProfile(REGSEC_GENERAL, REGVAL_LONGITUDE)
    If buf <> "" Then
        m_dLongitude = CDbl(buf)
    End If
    
    buf = GetProfile(REGSEC_GENERAL, REGVAL_ELEVATION)
    If buf <> "" Then
        m_nElevation = CInt(buf)
    End If
End Sub

Private Sub Class_Initialize()
    m_dLatitude = DEFAULT_LATITUDE
    m_dLongitude = DEFAULT_LONGITUDE
    m_nElevation = DEFAULT_ELEVATION
End Sub

