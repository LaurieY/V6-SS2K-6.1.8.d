VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AxisRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' -----------------------------------------------------------------------------'
' ==================
'   AxisRates.cls
' ==================
'
' Implementation of the ASCOM AxisRates Class
'
' Written: Laurie Yates (from Jon Brewster Celestron Code)
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 23-May-08 ley     Initial edit (Chris Rowland & Jon Brewster created an early version)
' -----------------------------------------------------------------------------

Option Explicit

Implements IAxisRates
'Default values
Private Const DEFAULT_AXISRATES0 As Double = 0.5

'Registry location
Private Const REGSEC_RATE0 As String = "RATE0"
Private Const REGVAL_AXISRATES0 As String = "Rate_for_Guiding"

Private Const REGVAL_SINGLEAXIS As String = "Single Axis Moves Only"
Private m_dAxisRate0x As Double  ' guide rate as n x sidereal
'Private m_dAxisRate0 As Double ' true guide rate
Private m_bSingleAxis As Boolean

' local variable to hold collection
Private mCol As Collection

'
' Creates the collection when this class is created
'
Private Sub Class_Initialize()

    Set mCol = New Collection
   LoadSettings
End Sub

'
' Destroys collection when this class is terminated
'
Private Sub Class_Terminate()

    Set mCol = Nothing
    
End Sub


' Add an element to the collection
' Example: set foo = x.Add(3.0, 1.0), or just x.Add(3.0, 1.0)

Friend Function Add(Maximum As Double, Minimum As Double) As Rate
g_Tracing.Trace (" axisrates.add " & CStr(Maximum))
    'create a new object
    Dim objNewMember As Rate
    Set objNewMember = New Rate
    
    'set the properties passed into the method
    objNewMember.Maximum = Maximum
    objNewMember.Minimum = Minimum
    mCol.Add objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

' Remove an element from the collection
' Example: x.Remove(5)

Friend Sub Remove(Index As Long)

    mCol.Remove Index
    
End Sub

' ============================
' IMPLEMENTATION OF IAxisRates
' ============================

' Get the number of elements in the collection.
' Example: Debug.Print x.Count

Private Property Get IAxisRates_Count() As Long
    IAxisRates_Count = mCol.Count
End Property

' Get an element in the collection
' Example: Set foo = x.Item(5)

Private Property Get IAxisRates_Item(ByVal Index As Long) As IRate
    Set IAxisRates_Item = mCol(Index)
End Property

' this property allows you to enumerate
' this collection with the For...Each syntax

Private Property Get IAxisRates_NewEnum() As stdole.IEnumVARIANT
'g_Tracing.Trace ("Into get IAxisRates_NewEnum")
    Set IAxisRates_NewEnum = mCol.[_NewEnum]
End Property

' ==============
' PUBLIC MEMBERS
' ==============

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get Item(ByVal Index As Long) As Rate
    Set Item = mCol(Index)
End Property

Public Property Get NewEnum() As stdole.IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'g_Tracing.Trace ("Into get NewEnum")

    Set NewEnum = mCol.[_NewEnum]
End Property
Public Sub SaveSettings()
'g_AxisRates.Remove (1)
 g_AxisRates.Add AxisRate0 * SIDEREAL_RATE_DEG_SEC, AxisRate0 * SIDEREAL_RATE_DEG_SEC
    g_Profile.WriteValue g_ID, REGVAL_AXISRATES0, CStr(AxisRate0), REGSEC_RATE0
    
    g_Profile.WriteValue g_ID, REGVAL_SINGLEAXIS, m_bSingleAxis, REGSEC_RATE0
End Sub
Property Get AxisRate0() As Double
AxisRate0 = m_dAxisRate0x * SIDEREAL_RATE_DEG_SEC

End Property
Property Let AxisRate0(ByVal newVal As Double)
m_dAxisRate0x = newVal / SIDEREAL_RATE_DEG_SEC
End Property

Property Get SingleAxis() As Boolean
SingleAxis = m_bSingleAxis

End Property
Property Let SingleAxis(ByVal newVal As Boolean)
m_bSingleAxis = CBool(newVal)
End Property

Public Sub LoadSettings()
    Dim buf As String
    
'Allow empty value or missing key in the registry
On Error Resume Next
    
    buf = g_Profile.GetValue(g_ID, REGVAL_AXISRATES0, REGSEC_RATE0)
    If buf <> "" Then
        m_dAxisRate0x = CDbl(buf)
    Else
        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_AXISRATES0, CStr(DEFAULT_AXISRATES0), REGSEC_RATE0
        m_dAxisRate0x = DEFAULT_AXISRATES0
    End If
     buf = g_Profile.GetValue(g_ID, REGVAL_SINGLEAXIS, REGSEC_RATE0)
    If buf <> "" Then
        m_bSingleAxis = CBool(buf)
    Else
        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_SINGLEAXIS, False, REGSEC_RATE0
        m_bSingleAxis = False
     
    End If
    
    
  'frmSetup.edt_GuideRate(0).Text = CStr(m_dAxisRate0)
   Me.Add m_dAxisRate0x * SIDEREAL_RATE_DEG_SEC, m_dAxisRate0x * SIDEREAL_RATE_DEG_SEC
  g_Tracing.Trace ("ending AxisRate load_settings  ")
End Sub
