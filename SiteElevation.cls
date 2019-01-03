VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiteElevation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SiteElevation"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------
'   =============
'   SITEELEVATION.CLS
'   =============
'
' Implementation of class for storing and retrieving Site Elevation settings as SS2K does not hold them.
'
' Written:  07-April   Laurie Yates
'
' Changes:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 07-April-08 LEY      Initial version added to SS2K driver 5.1.4
'---------------------------------------------------------------------
Option Explicit

'Default values
Private Const DEFAULT_SITEELEVATION As Double = 0

'Registry location
Private Const REGSEC_SITE As String = "Site"
Private Const REGVAL_SITEELEVATION As String = "SiteElevation"

Private m_dSiteElevation As Double

Private Sub Class_Initialize()
   'Initalize members with default values
    m_dSiteElevation = DEFAULT_SITEELEVATION
        
    'Override with registry settings
    LoadSettings

End Sub

Property Get SiteElevation() As Double
SiteElevation = m_dSiteElevation

End Property

Property Let SiteElevation(ByVal vNewValue As Double)
 m_dSiteElevation = vNewValue
 SaveSettings
End Property
Public Sub SaveSettings()
If m_dSiteElevation < -300 Then m_dSiteElevation = -300
If m_dSiteElevation > 10000 Then m_dSiteElevation = 10000

    g_Profile.WriteValue g_ID, REGVAL_SITEELEVATION, CStr(m_dSiteElevation), REGSEC_SITE
End Sub

Public Sub LoadSettings()
    Dim buf As String
    
'Allow empty value or missing key in the registry
On Error Resume Next
    
    buf = g_Profile.GetValue(g_ID, REGVAL_SITEELEVATION, REGSEC_SITE)
    If buf <> "" Then
        m_dSiteElevation = CDbl(buf)
    Else
        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_SITEELEVATION, CStr(DEFAULT_SITEELEVATION), REGSEC_SITE
    End If
End Sub
