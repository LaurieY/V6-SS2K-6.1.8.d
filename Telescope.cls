VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Telescope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ASCOM Telescope driver for Vixen SkySensor 2000-PC"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SiteElevation"
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
'   TLESCOPE.CLS
'   =============
'
' Implementation of ASCOM Vixen SkySensor 2000-PC telescope driver Telescope class
'
' Written:                             22-Aug-00   Robert B. Denny <rdenny@dc3.com>
' Adapted to support SkySensor2000-PC  14-Dec-00   Arne Danielsen
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 22-Aug-00 rbd     Initial edit
' 28-Aug-00 rbd     Integration testing completed. Add setup dialog.
' 14-Oct-00 rbd     Slew settling time
' 15-Oct-00 rbd     Fix Com Port registry setting name
' 13-Nov-00 rbd     Description returns String not Double, Park() does
'                   not take a command. ByVal on parameters and Lets.
'                   Allow Description when not connected.
' 14-Nov-00 rbd     TargetDeclination used HoursToHMS to format for
'                   Autostar, which barfs if there are spaces in the
'                   sexagesimal (DegreesToDMS generates spaces).
' 21-Nov-00 rbd     1.0.3 - DriverInfo is a String. Add DriverInfo
'                   code to take from project properties.
' 14-Dec-00 ad    Created SS2K project using the LX200 driver made by Robert B. Denny as template
' 16-Dec-00 ad    Implemented the following properties and methods:
'                 Properties
'                    AlignmentMode
'                    Connected
'                    Declination
'                    Description
'                    DriverInfo
'                    Name
'                    RightAscension
'                    SlewSettleTime
'                    TargetDeclination
'                    TargetRightAscension
'                 Methods
'                    AbortSlew
'                    CommandBlind
'                    CommandBool
'                    CommandString
'                    SlewToCoordinates
'                    SlewToCoordinatesAsync
'                    SlewToTarget
'                    SlewToTargetAsync
'                    SyncToCoordinates
'                    SyncToTarget
' 17-Dec-00 ad    Testing and modification of implemented properties and methods
' 04-Jan-01 ad    Moved settings into dedicated classes
' 14-Jan-01 ad    Finished implementing the following properties.
'                 Properties
'                    Slewing
'                    Altitude          (NB Only suported with ROM Version 2.03 and higer)
'                    Azimuth           (NB Only suported with ROM Version 2.03 and higer)
'                    SiteLongitude     (NB Only suported with ROM Version 2.03 and higer)
'                    SiteLatitude      (NB Only suported with ROM Version 2.03 and higer)
'                    SiderealTime      (NB Only suported with ROM Version 2.03 and higer)
'                 Added call to CheckConnection in CommandBlind, CommandBool &
'                 CommandString. Removed explicit call to CheckConnection from
'                 all methods which calls CommandBlind, CommandBool or CommandString
'                 directly or indirectly.
' 24-Jan-01 ad    Updated code to confirm with Beta 2 of the SDK
' 10-Feb-01 ad    Removed Abort flag
'                 Added call to AbortSlewing in Let TargetRightAscension and
'                 Let TargetDeclination to avoid conflict with async slewing
'                 Changed implementation of SyncToTarget to use CommandBool instead of
'                 CommandString
' 11-Feb-01 ad    In order to comply with the v1.2 of the telescope interface the following
'                 changes where made.
'                 Added the following properties:
'                    CanSlew
'                    CanSlewAsync
'                    CanSync
'                    CanPark
'                    CanFindHome
'                    CanSetPark
'                    CanUnpark
'                    CanSetTracking
'                    Tracking
'                 Add description and version info display to SetupDialog box.
' 08-Jun-01 rbd   New Profile.DeviceType to generalize Chooser and
'                 Profile.
' 08-Jun-01 rbd   Telescope 1.4: Add PulseGuide() & CanPulseGuide
' 05-Oct-01 rbd   1.4.2: Add ByVals as needed
' 12-Oct-01 rbd   1.4.3: Fix slew settle times > 30 sec
' 30-Dec-01 ad    1.4.4: Create driver specific registry settings
'                 programmatically if they don't exist allready.
'                 Added SkySensor ROM v2.06 as option. No new
'                 functionality available with this option.
' 21-Mar-03 rbd   2.2.1: Version change for ASCOM Platform 2.2
' 24-Jun-03 rbd   2.3.1 - Version change for Platform 2.3
' 03-Sep-03 rbd   3.0.1 - For Platform 3.0
' 23-Nov-04 rbd   4.0.1 - Version change for Platform 4.0
' 01-Mar-08 cdr   5.0.0 - Version change for Platform 5.0
'---------------------------------------------------------------------
' 12-Mar-08       5.1.0 - Modified to remove 180 degree adjustment on Get Azimuth
' 13 March 2008   5.1.1 - Adjusted Get SiderealTime() to be in hours instead of days
' 17 March 2008   5.1.2 - Implemented IsPulseGuiding, internally tracks number of pulse guide commands in progress to
'                         allow multi-axis guide pulses.
'                         Although currently the pulse guides wait to complete so only 1 at a time can be issued.
'                         Much biger change to go async and allow simultaneous pulse guides on 2 axes
' 2 April 2008    5.1.3 - Implemented a kludge for get and let Tracking.  If Tracking is set to off then send a series of :Q# commands with 1sec async waits
'                         until such time as Tracking is set to ON
' 6 April 2008    5.1.4 - Implemented property SiteElevation, the value entered at setup time is stored in the registry
'                         include a radiobutton for version 2.10, though there is no program action implemented on it.
' 11 July 2008    5.1.6 - Implemented Early Binding loading (using TLB) and the iTelescope inteface for Metaguide 4.0.2
'
' 11 July 2008    5.1.6b - Fixed the tracing to use the system derived temp directory
'
' 16 Nov 2008     5.1.6d - Fixed required properties to return a value rather than Error
'                          Get CanSetPierSide, CanSlewAltAz, CanSlewAltAzAsync, CanSyncAltAz, AtPark, UTCTime
' 8 Dec 2008      5.1.7 - Fixed errors in Conform caused by using 010 for value aout of range errors whereas Conform expects 80040405
'                           Include class for TrackingRates copied from MeadeLX200GPS project
'                           Passes Conform tests  for Conform V2.0.2.0
'                           MoveAxis not implemented and neither is SlewAltAz yet
'                 5.1.7b - Altered Synce for ROM < 2.10 to accept a single char return rather than 1#
' 22 March 2010   5.1.7c - Altered install  script to allow for comma separators & hence allow for Ascom versions with decimals in version nymber

' 7 July 2011     6.1.7b - RTe built for Ascom 6 and Released

' 20 August 2013  6.1.8a - Added an option to Force Single Star Alignment on ALL Sync  Commands
'---------------------------------------------------------------------

Option Explicit
Implements ITelescope

'Public Enum AlignmentModes
  '  algUnknown = -1
  '  algAltAz = 0
  '  algPolar = 1
  '  algGermanPolar = 2
'End Enum

'Public Enum GuideDirections
'    guideNorth = 0
'    guideSouth = 1
 '   guideEast = 2
 '   guideWest = 3
'End Enum

Private m_bSlewing As Boolean
Private m_bScopeSettled As Boolean
Private m_lAsyncEndTix As Long
Private m_bLongFormat As Boolean
Private m_dTargetDec As Double
Private m_dTargetRA As Double

Private m_bTracking As Boolean
  Private m_previousRA
Private m_previousDEC

Private m_bActive As Boolean        ' set while sending and receiving a command
Private m_preMoveRA, m_premoveDEC, m_movingRA, m_movingDEC As Boolean

Private m_Serial As DriverHelper.Serial
Private m_TrackingRates As TrackingRates

Private m_iIsPulseGuiding As Integer
Private WithEvents m_Timer As DriverHelper.Timer
Attribute m_Timer.VB_VarHelpID = -1


'
' Locally used Win32 API
'
Private Declare Function GetTickCount Lib "kernel32" () As Long
'local variable(s) to hold property value(s)
Private mvarSideOfPier As PierSide 'local copy
Public Sub SlewToAltAz(Azimuth As Double, Altitude As Double)
      Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property SlewToAltAz" & MSG_NOT_IMPLEMENTED
End Sub
Public Sub SlewToAltAzAsync(Azimuth As Double, Altitude As Double)
      Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property SlewToAltAz" & MSG_NOT_IMPLEMENTED
End Sub
Public Sub SyncToAltAz(ByVal Azimuth As Double, ByVal Altitude As Double)
             Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property SyncToAltAz" & MSG_NOT_IMPLEMENTED
End Sub

Public Property Let SideOfPier(ByVal vData As PierSide)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SideOfPier = Form1
          Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Let)SideOfPier" & MSG_NOT_IMPLEMENTED
End Property


Public Property Get SideOfPier() As PierSide
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SideOfPier
          Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get)SideOfPier" & MSG_NOT_IMPLEMENTED
End Property




Private Sub Class_Initialize()

   'g_Tracing.Trace ("starting telescope class initialize 5.1.4_trace")
   'Create helper classes

   Set g_Util = New DriverHelper.Util
   Set g_Profile = New DriverHelper.Profile
   Set g_Tracing = New Tracing
    g_Tracing.Trace ("starting telescope class initialize 6.1.8c_trace")
   g_Profile.DeviceType = "Telescope"       ' We're a Telescope driver
   g_Tracing.Trace ("setting serial")
      Set m_Serial = New DriverHelper.Serial
   
   'Create classes for storing settings
   Set g_Communication = New Communication
   Set g_Scope = New Scope
   Set g_ROMVersion = New ROMVersion
   Set g_SiteElevation = New SiteElevation
       Set g_AxisRates = New AxisRates  '************
         'g_AxisRates.LoadSettings '***********
         
         Set g_OneStar = New OneStar
Set g_DriveRates = New TrackingRates
    g_DriveRates.Add driveSidereal
g_DriveRate = driveSidereal
   g_Profile.Register g_ID, g_DESC
   
   m_dTargetDec = 999
   m_dTargetRA = 999 ' indicates NOT yet written to
   
   m_bTracking = True
   
End Sub

Private Sub Class_Terminate()
    
   On Error Resume Next
   Me.Connected = False                ' Force disconnect
   
   Set m_Serial = Nothing
   Set g_Util = Nothing
   Set g_Profile = Nothing
   
   Set g_Communication = Nothing
   Set g_Scope = Nothing
   Set g_ROMVersion = Nothing
   Set g_SiteElevation = Nothing

End Sub

Public Property Get AlignmentMode() As AlignmentModes

   'Make sure we are connected
   CheckConnected
   
   'It's impossible to determine the alignment mode of the SS2K programatically
   AlignmentMode = algPolar
    
End Property

Public Property Get Altitude() As Double

   If g_ROMVersion.Version >= 203 Then
      'Altitude is only supported by v 2.03 and higer.
      
      Altitude = g_Util.DMSToDegrees(Me.CommandString("GA"))
   Else
      Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) Altitude" & MSG_NOT_IMPLEMENTED
   End If
    
End Property
Public Property Get ApertureArea() As Double
          Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) ApertureArea" & MSG_NOT_IMPLEMENTED
End Property
Public Property Get ApertureDiameter() As Double

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) ApertureDiamater" & MSG_NOT_IMPLEMENTED

End Property
Public Property Get AtHome() As Boolean
AtHome = False
         ' Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) AtHome" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get AtPark() As Boolean
    
        AtPark = False
         ' Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) AtPark" & MSG_NOT_IMPLEMENTED
End Property

Public Function AxisRates(ByVal Axis As TelescopeAxes) As IAxisRates
'g_Tracing.Trace ("In AxisRates")
    Select Case Axis
    Case axisPrimary, axisSecondary
        Set AxisRates = g_AxisRates
    Case Else
        Set AxisRates = New AxisRates
    End Select
    '
        '  Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) AxisRates" & MSG_NOT_IMPLEMENTED
End Function

Public Property Get Azimuth() As Double
g_Tracing.Trace ("In Get Azimuth")
   If g_ROMVersion.Version >= 203 Then
      'Azimuth is only supported by v 2.03 and higer.
      
      Azimuth = g_Util.DMSToDegrees(Me.CommandString("GZ"))
      '
      ' Wrap it around 180 deg to the "real" azimuth.
      '  LEY Remove this to see if its better
'      Azimuth = Azimuth + 180#
 '     If Azimuth >= 360# Then
 '        Azimuth = Azimuth - 360#
  '    End If
  '  LEY TO HERE
   Else
      Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) Azimuth" & MSG_NOT_IMPLEMENTED
   End If
End Property

Public Property Get CanFindHome() As Boolean

    CanFindHome = False

End Property

Public Property Get CanPark() As Boolean

    CanPark = False
    
End Property

Public Property Get CanPulseGuide() As Boolean
 g_Tracing.Trace ("In CanPulseGuide ")
    CanPulseGuide = True

End Property

Public Property Get CanSetPark() As Boolean

    CanSetPark = False

End Property

Public Property Get CanSetPierSide() As Boolean
    CanSetPierSide = False
          'Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (CanSetPierSide" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get CanSetTracking() As Boolean

    'CanSetTracking = False
    CanSetTracking = True
End Property

Public Property Get CanSlew() As Boolean

    CanSlew = True
    
End Property

Public Property Get CanSlewAltAz() As Boolean
   CanSlewAltAz = False
         ' Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property CanSlewAltAz" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get CanSlewAsync() As Boolean

    CanSlewAsync = True

End Property

Public Property Get CanSlewAltAzAsync() As Boolean
  CanSlewAltAzAsync = False
          'Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property CanSlewAltAzAsync" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get CanSync() As Boolean

    CanSync = True
    
End Property

Public Property Get CanSyncAltAz() As Boolean
    CanSyncAltAz = False
          '    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property CanSyncAltAz" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get CanUnpark() As Boolean

    CanUnpark = False

End Property

Public Property Get Connected() As Boolean

    Connected = m_Serial.Connected
    
End Property

Public Property Let Connected(ByVal newVal As Boolean)
   Dim buf As String

   If Not m_Serial.Connected And newVal Then           ' CONNECTING
      g_Util.LoadDLL App.Path & "\astro32.dll"                           ' Assure this is loaded and available

      '
      ' (1) Set up the communications link.
      '
      m_Serial.Port = g_Communication.Port
      m_Serial.Speed = g_Communication.Speed
      m_Serial.ReceiveTimeout = 10                     ' 5 second timeout
      m_Serial.Connected = True                       ' Grab the serial port
      '
      ' (2) Determine that there is an SS2K present.
      '
      On Error GoTo NO_SCOPE                          ' -- error trap --
      Me.CommandString "GR"                           ' Try getting RA
      On Error GoTo 0                                 ' -- end trap --

      m_Serial.ClearBuffers                           ' Clear serial buffers
      m_bLongFormat = SetLongFormat(True)             ' Try for long format

   ElseIf m_Serial.Connected And Not newVal Then       ' DISCONNECTING
      On Error Resume Next                            ' Best efforts...
      SetLongFormat False
      m_Serial.ClearBuffers                           ' Clear serial buffers
      m_Serial.Connected = False                      ' Release COM port
   End If

   Exit Property

NO_SCOPE:
   m_Serial.ClearBuffers
   m_Serial.Connected = False                          ' Release the port
   Err.Raise SCODE_NO_SCOPE, ERR_SOURCE, MSG_NO_SCOPE
End Property

Public Property Get Declination() As Double
g_Tracing.Trace ("In Get Declination")
   'Send command to SS2K
   Declination = g_Util.DMSToDegrees(Me.CommandString("GD"))

End Property

Public Property Get DeclinationRate() As Double

    'Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
      '   "Property (Get) DeclinationRate" & MSG_NOT_IMPLEMENTED
      DeclinationRate = 0 'altered Dec 2008 for Conform

End Property

Public Property Let DeclinationRate(ByVal newVal As Double)

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Let) DeclinationRate" & MSG_NOT_IMPLEMENTED

End Property

Public Function DestinationSideOfPier(ByVal RightAscension As Double, ByVal Declination As Double) As PierSide
   
          Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property DestinationSideOfPier" & MSG_NOT_IMPLEMENTED
End Function

Public Property Let DoesRefraction(ByVal RHS As Boolean)
  
          Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Let)DoesRefraction" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get DoesRefraction() As Boolean
      Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get)DoesRefraction" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get Description() As String

       Description = "SkySensor 2000-PC controlled telescope."
    
End Property
Public Property Get DriverVersion() As String

    '
    DriverVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    
End Property
Public Property Get DriverInfo() As String
    '
    ' Use the Project/Properties sheet, Make tab, to set these
    ' items. That way they will show in the Version tab of the
    ' Explorer property sheet, and the exact same data will
    ' show in Telescope.DriverInfo.
    '
    DriverInfo = App.FileDescription & " " & _
                App.Major & "." & App.Minor & "." & App.Revision
    If App.CompanyName <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.CompanyName
    If App.LegalCopyright <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.LegalCopyright
    If App.Comments <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.Comments
    
End Property

Public Property Get FocalLength() As Double

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) FocalLength" & MSG_NOT_IMPLEMENTED

End Property


Public Property Let GuideRateDeclination(ByVal vData As Double)
       g_Tracing.Trace ("In Let GuideRateDEC with " & CStr(vData) & " --IGNORED")
        '  Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
      '   "Property (Let)GuideRateDeclination" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get GuideRateDeclination() As Double
      GuideRateDeclination = AxisRates(axisPrimary).Item(1).Minimum  ' item 1 is the default GUIDING rate as set in the chooser
g_Tracing.Trace ("In Get GuideRateDEC Value " & GuideRateDeclination)
       '   Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
       '  "Property (Get)GuideRateDeclination" & MSG_NOT_IMPLEMENTED
End Property

Public Property Let GuideRateRightAscension(ByVal vData As Double)
     g_Tracing.Trace ("In Let GuideRateRA with " & CStr(vData) & " --IGNORED")
     '     Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
       '  "Property (Let)GuideRateRightAscension" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get GuideRateRightAscension() As Double
  GuideRateRightAscension = AxisRates(axisPrimary).Item(1).Minimum ' item 1 is the default GUIDING rate as set in the chooser
   g_Tracing.Trace ("In Get GuideRateRA Value " & GuideRateRightAscension)

         ' Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
        ' "Property (Get)GuideRateRightAscension" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get Name() As String
       
       Name = "SS2K"

End Property


Public Sub MoveAxis(ByVal Axis As TelescopeAxes, ByVal Rate As Double)
g_Tracing.Trace ("In  MoveAxis with Axis " & CStr(Axis) & " Rate " & CStr(Rate))

  Me.ScopeMoveAxis Axis, Rate
         ' Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property MoveAxis" & MSG_NOT_IMPLEMENTED
End Sub

Public Property Get RightAscension() As Double
g_Tracing.Trace ("In Get RightAscension")
   'Send command to SS2K
   RightAscension = g_Util.HMSToHours(Me.CommandString("GR"))
   
End Property

Public Property Get RightAscensionRate() As Double

   ' Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
     '    "Property (Get) RightAscensionRate" & MSG_NOT_IMPLEMENTED
      RightAscensionRate = 0  ' altered Dec 2008 for Conform

End Property

Public Property Let RightAscensionRate(ByVal newVal As Double)

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Let) RightAscensionRate" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get SiderealTime() As Double
g_Tracing.Trace ("In Get SiderealTime")
   If g_ROMVersion.Version >= 203 Then
      'SiderealTime is only supported by v 2.03 and higer.
      
      SiderealTime = CDate(Me.CommandString("GS")) * 24 ' the x 24 added by LY 13/3/2008
      '  as the value returned by "GS" is in days, telescope 2.0 ref says it should be hours
   Else
      Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
            "Property (Get) SiderealTime" & MSG_NOT_IMPLEMENTED
   End If
End Property

Public Property Get SiteElevation() As Double
g_Tracing.Trace ("In Get SiteElevation")
   ' Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) SiteElevation" & MSG_NOT_IMPLEMENTED
    SiteElevation = g_SiteElevation.SiteElevation
End Property

Public Property Let SiteElevation(ByVal newVal As Double)
g_Tracing.Trace ("SiteElevation write with value " + CStr(newVal))
    'Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Let) SiteElevation" & MSG_NOT_IMPLEMENTED
  
   'Make sure we have legal values
   If (newVal < -300) Or (newVal > 10000) Then
       ' Err.Raise SCODE_PROP_RANGE_ERROR, ERR_SOURCE, , MSG_PROP_RANGE_ERROR
       Err.Raise SCODE_INVALID_VALUE, ERR_SOURCE, , MSG_INVALID_VALUE
   Else
        
    g_SiteElevation.SiteElevation = newVal
    End If
End Property
         



Public Property Get SiteLatitude() As Double
g_Tracing.Trace ("In Get SiteLatitude")
   If g_ROMVersion.Version >= 203 Then
      'SiteLatitude is only supported by v 2.03 and higer.
      
      SiteLatitude = g_Util.DMSToDegrees(Me.CommandString("Gt"))
   Else
      Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
            "Property (Get) SiteLatitude" & MSG_NOT_IMPLEMENTED
   End If

End Property

Public Property Let SiteLatitude(ByVal newVal As Double)

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Let) SiteLatitude" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get SiteLongitude() As Double
g_Tracing.Trace ("In Get SiteLongitude")
   If g_ROMVersion.Version >= 203 Then
      'SiteLongitude is only supported by v 2.03 and higer.
      
      ' Convert from 0-359 westward to +/- 180 +E
      SiteLongitude = -g_Util.DMSToDegrees(Me.CommandString("Gg"))
      If SiteLongitude < -180# Then
         SiteLongitude = SiteLongitude + 360#
      End If
   Else
      Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
            "Property (Get) SiteLongitude" & MSG_NOT_IMPLEMENTED
   End If
End Property

Public Property Let SiteLongitude(ByVal newVal As Double)

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Let) SiteLongitude" & MSG_NOT_IMPLEMENTED

End Property
Public Property Let Slewing(ByVal newValue As Boolean)
m_bSlewing = newValue
End Property

Public Property Get Slewing() As Boolean
   Dim buf As String
   Dim dCurrentDec As Double
   Dim dCurrentRA As Double
   
   If m_bSlewing Then
      'Check to see if scope is still slewing
      
      'The SS2K does not support any distance bars, so we have to match current RA & Dec against
      'target RA & Dec in order to see if the slew have finished
      
      'Retrieve current RA & Dec
      dCurrentRA = RightAscension
      dCurrentDec = Declination
      g_Tracing.Trace ("In Slewing?  RA is  " & g_Util.HoursToHMS(RightAscension) & "  DEC is " & g_Util.DegreesToDMS(Declination, ":", ":", ""))
      'Due to rounding errors the current RA & Dec might actually never becomes exactly equal to target RA & Dec.
      'I choose to consider the slewing as finished when current RA & Dec is within +- 0.1 degrees of the target RA & Dec
      'The remaining part of the slewing has to be incorperated in the settle time
      If (dCurrentRA < m_dTargetRA + 0.1) And (dCurrentRA > m_dTargetRA - 0.1) And _
         (dCurrentDec < m_dTargetDec + 0.1) And (dCurrentDec > m_dTargetDec - 0.1) Then
         'We are within +- 0.1 degrees of are target RA &Dec
         
         'Now check if enough time has elapsed for the scope to settle
         If (GetTickCount() > m_lAsyncEndTix) Then
            'Scope settle time has elapsed.
            
            'Flag slewing as finished
            m_bSlewing = False
         End If
         
      Else
         'Slewing still in progress
         'Update end time for scope to have settled.
         m_lAsyncEndTix = GetTickCount() + (g_Scope.SlewSettleTime * 1000)
      End If
   End If
   
   Slewing = m_bSlewing
End Property

Public Property Get SlewSettleTime() As Integer
g_Tracing.Trace ("In Get SlewSettleTime")
    SlewSettleTime = g_Scope.SlewSettleTime
    
End Property

Public Property Let SlewSettleTime(ByVal newVal As Integer)
   
   'Make sure we have legal values
   If (newVal < 0) Or (newVal > 100) Then
        Err.Raise SCODE_INVALID_VALUE, , MSG_INVALID_VALUE
   End If
        
    g_Scope.SlewSettleTime = newVal
    
End Property


Public Property Get TargetDeclination() As Double

   'Make sure we are connected even though we are not going to talk with the SS2K
   'This is done to in order to behave similar to other telescope drivers
   'Also it doesn't make sense to be able to retrieve TargetDeclination without
   'beeing connected to a scope
   CheckConnected
   If m_dTargetDec = 999 Then     ' it has been set previously if 999 not yet set so error
  '  ' indicates NOT yet written to
   Err.Raise SCODE_PROP_NOT_SET, ERR_SOURCE, MSG_PROP_NOT_SET
   
   Else
    g_Tracing.Trace ("In Get TargetDeclination " & g_Util.DegreesToDMS(m_dTargetDec, ":", ":", ""))
   'SS2K does not support retrieval of Target Declination, so we return the internally stored Declination.
   TargetDeclination = m_dTargetDec
   End If

End Property

Public Property Let TargetDeclination(ByVal newVal As Double)
   Dim sexa As String
  g_Tracing.Trace ("In Let TargetDeclination " & g_Util.DegreesToDMS(newVal, ":", ":", ""))
   'Make sure we have legal values.
   If newVal < -90# Or newVal > 90# Then
       Err.Raise SCODE_INVALID_VALUE, ERR_SOURCE, MSG_INVALID_VALUE
   End If
   
   'Make sure any slewing in process is terminated before setting new target Dec
   If Me.Slewing Then
      Me.AbortSlew
   End If
   
   'SS2K does not support retrieval of Target Declination, so we store this internally instead.
   m_dTargetDec = newVal
   
   'Prepare Declination string
   sexa = g_Util.DegreesToDMS(newVal, ":", ":", "")    ' SS2K compatible
   If Left$(sexa, 1) <> "-" Then sexa = "+" & sexa     ' Both need leading '+'
   
   'Send command to SS2K
   Me.CommandBool "Sd " & sexa
    
End Property

Public Property Get TargetRightAscension() As Double

   'Make sure we are connected even though we are not going to talk with the SS2K
   'This is done to in order to behave similar to other telescope drivers
   'Also it doesn't make sense to be able to retrieve TargetRightAscension without
   'beeing connected to a scope
   CheckConnected
   If m_dTargetRA = 999 Then      ' it has been set previously if 999 not yet set so error
   Err.Raise SCODE_PROP_NOT_SET, ERR_SOURCE, MSG_PROP_NOT_SET
   
   Else
    g_Tracing.Trace ("In Get TargetRightAscension " & g_Util.HoursToHMS(m_dTargetRA))
      'SS2K does not support retrieval of Target RightAscension, so we return the internally stored RA.
   TargetRightAscension = m_dTargetRA
   End If

End Property

Public Property Let TargetRightAscension(ByVal newVal As Double)
g_Tracing.Trace ("In Let TargetRightAscension " & g_Util.HoursToHMS(newVal))
   'Make sure we have legal values.
   If newVal < 0# Or newVal >= 24# Then
       'Err.Raise SCODE_PROP_RANGE_ERROR, ERR_SOURCE, MSG_PROP_RANGE_ERROR
       Err.Raise SCODE_INVALID_VALUE, ERR_SOURCE, MSG_INVALID_VALUE
   End If
   
   'Make sure any slewing in process is terminated before setting new target RA
   If Me.Slewing Then
      Me.AbortSlew
   End If
   
   'SS2K does not support retrieval of Target RightAscension, so we store this internally instead.
   m_dTargetRA = newVal
       
   'Send command to SS2K
   Me.CommandBool "Sr " & g_Util.HoursToHMS(newVal)

End Property

Public Property Get Tracking() As Boolean

    'Tracking = True
    Tracking = m_bTracking
    
End Property

Public Property Let Tracking(ByVal newVal As Boolean)

    'Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property (Let) Tracking" & MSG_NOT_IMPLEMENTED
    m_bTracking = newVal
    If Not m_bTracking Then
    Set m_Timer = New DriverHelper.Timer
    m_Timer.Interval = 500
    m_Timer.Enabled = True          ' Start it
    Me.CommandBlind ("Q")
    
    'While Not m_bTracking
    
    'Me.CommandBlind ("Q")
    'g_Util.WaitForMilliseconds (1000)
    End If

End Property

Public Property Let TrackingRate(ByVal RHS As DriveRates)
 
 ' implemented for equatorial mounts only
' untested
   ' ScopeTrackingRate RHS
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property (Let) TrackingRate" & MSG_NOT_IMPLEMENTED
End Property


Public Property Get TrackingRate() As DriveRates
 
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property (Get) TrackingRate" & MSG_NOT_IMPLEMENTED
                
                
End Property
Public Property Get UTCDate() As Date
    CheckConnected
    Dim dUTCDate As Date
    dUTCDate = CDate(Me.CommandString("GL"))                   ' Local time
    ' DO NOT USE CDate() here!!! Scope always returns MM/DD/YY even in Europe
    g_Tracing.Trace ("In get UTCDate Local Time" + CStr(dUTCDate))
    dUTCDate = dUTCDate + mdy_vb(Me.CommandString("GC"))        ' Calendar date
    g_Tracing.Trace ("In get UTCDate Calendar Date" + CStr(dUTCDate))
    dUTCDate = dUTCDate + (CDbl(Me.CommandString("GG")) / 24)   ' Now UTC
        g_Tracing.Trace ("Finishing Get UTCDate UTC" + CStr(dUTCDate))
   ' Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Get) UTCDate" & MSG_NOT_IMPLEMENTED
         UTCDate = dUTCDate

End Property
Public Property Get TrackingRates() As Object
   
'Dim tmp As DriveRates
'tmp = 999
'tmp = g_DriveRate
' g_Tracing.Trace ("Into get TrackingRates where it has type " + TypeName(g_DriveRates))
 'g_Tracing.Trace ("Into get TrackingRates where it gets value " + CStr(g_DriveRates.Item(1)))
 Set TrackingRates = g_DriveRates
    'Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
       '         "Property (Get) TrackingRates" & MSG_NOT_IMPLEMENTED
End Property

Public Property Let UTCDate(ByVal newVal As Date)
 Dim buf As String
    Dim tbuf As String
    Dim I As Integer
    Dim en As Boolean
    Dim UTCOffset As Integer
    
    CheckConnected
    
    UTCOffset = CInt(utc_offs() / 3600)
    buf = Format$(UTCOffset, "00")
    If UTCOffset >= 0 Then buf = "+" & buf          ' Must have explicit +
    If Not Me.CommandBool("SG " & buf) Then
        Err.Raise SCODE_BAD_UTCOFFSET, ERR_SOURCE, MSG_BAD_UTCOFFSET
    End If
    
        newVal = CDate(CDbl(newVal) - (UTCOffset / 24#))    ' Convert to local time
    If Not Me.CommandBool("SL " & Format$(newVal, "hh:nn:ss")) Then
        Err.Raise SCODE_BAD_DATETIME, ERR_SOURCE, MSG_BAD_DATETIME
    End If
   ' Make date string in scope format. Check the scope's
    ' date and if the same, skip the date set. Otherwise
    ' set it. We do things here so that various local date
    ' formats like 05.12.98 will work...
    '
    If mdy_vb(Me.CommandString("GC")) = CDate(Fix(newVal)) Then
        GoTo TS_OK                      ' That's it
    End If
    '
    ' OK, we're going to change the date...
    '
    m_Serial.Transmit ":SC " & vb_mdy(Fix(newVal)) & "#"
    '
    ' Wait for the OK boolean then bail if the date was bad
    '
    buf = m_Serial.ReceiveCounted(1)
    If Left$(buf, 1) = "0" Then     ' Illegal date
        Err.Raise SCODE_BAD_DATETIME, ERR_SOURCE, MSG_BAD_DATETIME
    End If
    '
    ' Wait for the "Updating..." then toss the blanks message
    '
    m_Serial.ReceiveTerminated "#"
    m_Serial.ReceiveTerminated "#"
    
TS_OK:
    '
    'Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Property (Let) UTCDate" & MSG_NOT_IMPLEMENTED

End Property
'
' Convert VB style date/time to MM/DD/YY for use as input to the scope
'
Private Function vb_mdy(dt As Date) As String
    Dim yy As Integer
    
    dt = CVDate(Fix(dt))
    yy = Year(dt) Mod 100
    vb_mdy = Format$(Month(dt), "00") & "/" & _
            Format$(Day(dt), "00") & "/" & _
            Format$(yy, "00")

End Function

Public Sub AbortSlew()
  
   
   '**  If its in a Move command and abort is called stop the appropriate direction
   Dim q_any As String
   If m_movingDEC Then
   q_any = "Q" & m_previousDEC
    m_movingDEC = False
    ElseIf m_movingRA Then
   q_any = "Q" & m_previousRA
   m_movingRA = False
   Else
   q_any = "Q"
   End If
    g_Tracing.Trace ("In AbortSlew with Command " & q_any)
   Me.CommandBlind q_any                             ' Halt the slew
   m_bSlewing = False
   g_Util.WaitForMilliseconds 1000                 ' Give the SS2K 3 seconds to react
    
End Sub
Public Function CanMoveAxis(ByVal Axis As TelescopeAxes) As Boolean
g_Tracing.Trace ("In CanMoveAxis with axis  " + CStr(Axis))
   ' Select Case Axis
   ' Case axisPrimary, axisSecondary
    '    CanMoveAxis = CanSetTracking
   ' Case Else
    '    CanMoveAxis = False
   ' End Select
   CanMoveAxis = True
 ' CanMoveAxis = False
End Function
Public Sub CommandBlind(ByVal Command As String)
 g_Tracing.Trace ("In CommandBlind with command " + Command)
   'Make sure we are connected
   CheckConnected
   
   m_Serial.ClearBuffers                           ' Clear remaining junk in buffer
   m_Serial.Transmit "#:" & Command & "#"
    
End Sub

Public Function CommandBool(ByVal Command As String) As Boolean
   Dim buf As String
    g_Tracing.Trace ("In CommandBool with command " + Command)
   'Make sure we are connected
   CheckConnected
   
   m_Serial.ClearBuffers                           ' Clear remaining junk in buffer
   m_Serial.Transmit "#:" & Command & "#"
   buf = m_Serial.ReceiveCounted(1)                ' Just a 1 or 0
   If buf = "1" Then
       CommandBool = True
   Else
       CommandBool = False
   End If

End Function

Public Function CommandString(ByVal Command As String) As String
   Dim buf As String
 g_Tracing.Trace ("In CommandString with command " + Command)
   'Make sure we are connected
   CheckConnected
    
   m_Serial.ClearBuffers                           ' Clear remaining junk in buffer
   m_Serial.Transmit "#:" & Command & "#"
   buf = m_Serial.ReceiveTerminated("#")
   g_Tracing.Trace ("Done CommandString with command " + Command + " Result = " + Left$(buf, Len(buf) - 1))
   If buf <> "" Then                   ' Overflow protection
       CommandString = Left$(buf, Len(buf) - 1)   ' Strip '#'
   Else
       CommandString = ""
   End If
   
End Function

Public Sub FindHome()

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Method FindHome" & MSG_NOT_IMPLEMENTED

End Sub

Public Sub Park()
    
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Method Park" & MSG_NOT_IMPLEMENTED

End Sub





Public Sub PulseGuide(ByVal Direction As GuideDirections, ByVal Duration As Long)

    g_Tracing.Trace ("In PulseGuide- Direction& Value = " & CStr(Direction) & ":" & CStr(Duration))
    m_iIsPulseGuiding = m_iIsPulseGuiding + 1 ' caters for 2 axis guiding
    Dim d As String
    
    Me.CommandBlind "RG"                ' Assure we're in Guide rate
    Select Case Direction
        Case guideNorth:    d = "n"
        Case guideSouth:    d = "s"
        Case guideEast:     d = "e"
        Case guideWest:     d = "w"
    End Select
    Me.CommandBlind "M" & d
    g_Util.WaitForMilliseconds Duration
    Me.CommandBlind "Q" & d
    g_Util.WaitForMilliseconds 200
    Me.CommandBlind "Q" & d
    m_iIsPulseGuiding = m_iIsPulseGuiding - 1
    If m_iIsPulseGuiding < 0 Then Err.Raise 999, ERR_SOURCE, "pulse guide count <0"
End Sub

Public Sub SetPark()

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Method SetPark" & MSG_NOT_IMPLEMENTED

End Sub

Public Sub SetupDialog()

   If m_Serial.Connected Then
      Err.Raise SCODE_SETUP_CONNECTED, ERR_SOURCE, MSG_SETUP_CONNECTED
   End If
   
   frmSetup.Show 1
   Unload frmSetup
    
End Sub

Public Sub SlewToCoordinates(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
     g_Tracing.Trace ("In SlewToCoordinates  RA is  " & g_Util.HoursToHMS(RightAscension) & "  DEC is " & g_Util.DegreesToDMS(Declination, ":", ":", ""))

   'Update target coordinates
   Me.TargetRightAscension = RightAscension
   Me.TargetDeclination = Declination
   
   'Call method for slewing to target.
   Me.SlewToTarget
    
End Sub

Public Sub SlewToCoordinatesAsync(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
       g_Tracing.Trace ("In SlewToCoordinatesAsync  RA is  " & g_Util.HoursToHMS(RightAscension) & "  DEC is " & g_Util.DegreesToDMS(Declination, ":", ":", ""))
   'Update target coordinates
   Me.TargetRightAscension = RightAscension
   Me.TargetDeclination = Declination
   
   'Call method for slewing to target asynchronous
   Me.SlewToTargetAsync
    
End Sub

Public Sub SlewToTarget()
g_Tracing.Trace ("In SlewToTarget ")
   Me.SlewToTargetAsync                        ' Start slew
   While Me.Slewing                            ' Wait for slew to complete
        g_Util.WaitForMilliseconds 200          ' Be responsive to AbortSlew()
   Wend

End Sub

Public Sub SlewToTargetAsync()
   Dim stat As Integer
   Dim msg As String
   Dim code As Long
   Dim buf As String
    g_Tracing.Trace ("In SlewToTargetAsync ")
   '
   ' This is where all the work is done for slewing. The other
   ' methods eventually call this one.
   '
   
   'Make sure any slewing in progress is terminated
   If Me.Slewing Then
      Me.AbortSlew
   End If
   
   'Send command to SS2K
   Me.CommandBlind "MS"
   
   'Adjust flags
   m_bSlewing = True                           ' Could be a race here!
   
   'Set initial end time for scope to have settled
   m_lAsyncEndTix = GetTickCount() + (CLng(g_Scope.SlewSettleTime) * 1000&)
   
   stat = CInt(m_Serial.ReceiveCounted(1))     ' Get only the start-status digit
   If stat <> 0 Then
      ' Failed to start slew
       Select Case stat
           Case 1:                             ' Below horizon
               msg = MSG_BELOW_HORIZ
               code = SCODE_BELOW_HORIZ
           Case 2:                             ' Below minimum elevation
                msg = MSG_BELOW_MIN_EL
               code = SCODE_BELOW_MIN_EL
           Case Else
               msg = MSG_SLEW_FAIL
               code = SCODE_SLEW_FAIL
       End Select
       m_bSlewing = False                      ' Clear slewing flag
       Err.Raise code, ERR_SOURCE, msg         ' FAILED
   End If
    
End Sub

Public Sub SyncToCoordinates(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
      g_Tracing.Trace ("In SyncToCoordinates  RA is  " & g_Util.HoursToHMS(RightAscension) & "  DEC is " & g_Util.DegreesToDMS(Declination, ":", ":", ""))
   'Update target coordinates
   Me.TargetRightAscension = RightAscension
   Me.TargetDeclination = Declination
   
   'Call method for synchronizing to target
   Me.SyncToTarget
    
End Sub

Public Sub SyncToTarget()
    Dim buf As String
    Dim tempRA As Double
    Dim tempDEC As Double
    
    

            DoSync
   
        ' 1st sync ok see if 1 star falignment forcing needed
            If g_OneStar.AStar Then
            ' force a sync far away
            
            g_Tracing.Trace ("In SyncToTarget  g_OneStar.Astar is  " & g_OneStar.AStar)
                tempRA = Me.TargetRightAscension ' Remember the real RA value
                tempDEC = Me.TargetDeclination
                If Me.TargetRightAscension < 12 Then
                    Me.TargetRightAscension = tempRA + 12
                Else
                    Me.TargetRightAscension = tempRA - 12
                End If ' of if >12
                Me.TargetDeclination = tempDEC ' probably of no value but ensures the sync uses the correct value
           DoSync 'Far away
            Me.TargetRightAscension = tempRA  '  Forces to the correct sync location
            Me.TargetDeclination = tempDEC  ' probably of no value but ensures the sync uses the correct value
            DoSync ' Sync back to the correct place
            End If ' of onestar
    

End Sub

Public Sub Unpark()

    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
         "Method Unpark" & MSG_NOT_IMPLEMENTED

End Sub

' ========================
' PRIVATE HELPER FUNCTIONS
' ========================

'
' CheckConnected() - Raise an error if the scope is not connected
'
Private Sub CheckConnected()

    If Not m_Serial.Connected Then _
        Err.Raise SCODE_NOT_CONNECTED, _
                    ERR_SOURCE, _
                    MSG_NOT_CONNECTED
End Sub
'
' SetLongFormat() - Set scope "long" format on or off
'
' The scope has only a toggle for the format, so we need to
' first test the format and then conditionally issue the
' toggle to set the desired format.
'
' Returns the format after the function is called.
'
Private Function SetLongFormat(b As Boolean) As Boolean
   Dim buf As String
   
   'Try retrieving RA. This will be 7 characters in short format and 8 in long
   buf = Me.CommandString("GR")            ' Get RA
    
   '
   ' Scope has some sort of race condition with "U", so
   ' we just wait a half sec. here and then do it.
   '
   g_Util.WaitForMilliseconds 500
   If ((b And (Len(buf) <= 7)) Or (Not b And (Len(buf) > 7))) Then
      Me.CommandBlind "U"                 ' change setting
   End If
    
   buf = Me.CommandString("GR")            ' Check final result
   If Len(buf) <= 7 Then
       SetLongFormat = False               ' Now/still in short format
   Else
       SetLongFormat = True                ' Now in long format
   End If
    
End Function


Public Property Get IsPulseGuiding() As Boolean
g_Tracing.Trace ("In IsPulseGuiding value " & m_iIsPulseGuiding)
IsPulseGuiding = m_iIsPulseGuiding > 0
End Property
Public Property Get InterfaceVersion() As Integer
    InterfaceVersion = 2
End Property



Private Sub m_Timer_Tick()
If m_bTracking Then
 m_Timer.Enabled = False
 Else
  Me.CommandBlind ("Q")
 End If
End Sub
Public Property Let SetTracing(ByVal newVal As Boolean)
         g_Profile.WriteValue g_ID, REGVAL_TRACE, newVal, REGSEC_TRACE

End Property
Public Property Get CanSetDeclinationRate() As Boolean

    CanSetDeclinationRate = False
End Property
Public Property Get CanSetRightAscensionRate() As Boolean

    CanSetRightAscensionRate = False
End Property

Public Property Get CanSetGuideRates() As Boolean
    CanSetGuideRates = False 'CanPulseGuide()
End Property

Public Property Get EquatorialSystem() As EquatorialCoordinateType

EquatorialSystem = equLocalTopocentric
    
End Property


Public Sub ScopeMoveAxis(ByVal Axis As TelescopeAxes, Rate As Double)
    ' This will command the scope to move in the specified directions
    'at the specified rates in Degrees per second.
    ' If the rate is not the same as the guiderate
    ' as specified in the ratescollection at index 1 then carry on regardless
    
    
    '  added on 2 June 2008.
    ' if singleaxis moves checked in setup then if a DEC move arrives quit any RA move.  This is a temp fix
    ' for a possible Metaguide issue
    
    g_Tracing.Trace ("In ScopeMoveAxis axis and rate = " & CStr(Axis) & " || " & CStr(Rate))
    CheckConnected   '********* uncomment after testing
    m_preMoveRA = False
    m_premoveDEC = False
     g_Tracing.Trace ("start Wait in ScopeMoveAxis")
    Wait ' for any command to finish
g_Tracing.Trace ("End Wait in ScopeMoveAxis")

    Dim s_d, s_r As String
    Dim s_ddefaultRA, s_ddefaultDEC As String
    s_ddefaultDEC = "n"
    s_ddefaultRA = "w"  '  if no previous move commands in this direction need a value else it will be a Q for all
    Dim d_rateDiff, d_gRate, d_rateError As Double
    d_gRate = GuideRateRightAscension
    d_rateDiff = Abs(SIDEREAL_RATE_DEG_SEC - Abs(Rate))
    d_rateError = Abs(d_rateDiff - Abs(d_gRate))

  
    Me.CommandBlind "RG"                ' Assure we're in Guide rate
    'If Rate = 0 Then r = "Q" Else r = "M" 'either stop or move commands
    
        s_r = "M"
    Select Case Axis
        Case axisPrimary:
        m_preMoveRA = True
         If (d_rateError > g_derrorMargin) And (d_rateDiff > ONE_PCT) Then
         g_Tracing.Trace ("Error in ScopeMoveAxis rate not acceptable ")
         Err.Raise SCODE_MOVE_RATE_ERROR, _
                    ERR_SOURCE, _
                    MSG_MOVE_RATE_ERROR
                    End If
        If d_rateDiff < ONE_PCT Then
        s_r = "Q"
        s_d = m_previousRA
        m_movingRA = False
        'If s_d = "" Then s_d = "w"
        Else   ' this is a move in RA east or west
                
                If Rate < SIDEREAL_RATE_DEG_SEC Then
                s_d = "e"
                
                Else
                s_d = "w" ' RA axis
                End If
                If ((m_previousRA <> s_d) And (m_movingRA)) Then
                ' its a reverse of direction without a stop being received.  Insert a stop command and 200msec delay
                g_Tracing.Trace ("Move Axis-> INSERT CommandBlind  Q" & m_previousRA)
                Me.CommandBlind "Q" & m_previousRA
                g_Util.WaitForMilliseconds 200
                Me.CommandBlind "Q" & m_previousRA
                m_movingRA = False
                End If
                m_previousRA = s_d
                m_movingRA = True
                
            End If
        
        Case axisSecondary:   ' *** DEC moves or quits
        If Rate = 0 Then
        s_r = "Q"
        s_d = m_previousDEC
        m_movingDEC = False
            If s_d = "" Then
            g_Tracing.Trace ("Move Axis-> CommandBlind-NOT Sent no previous MOVE in this axis")
            Exit Sub
            '**************************************
            '*  Need to send the appropriate Q command
            End If
        Else
                If Rate < 0 Then
                s_d = "s"
                Else
                s_d = "n" ' DEC axis
                End If
                If ((m_previousDEC <> s_d) And (m_movingDEC)) Then
                ' its a reverse of direction without a stop being received.  Insert a stop command and 200msec delay
                g_Tracing.Trace ("Move Axis-> INSERT CommandBlind  Q" & m_previousDEC)
                Me.CommandBlind "Q" & m_previousDEC
                g_Util.WaitForMilliseconds 200
                Me.CommandBlind "Q" & m_previousDEC
                m_movingDEC = False
                End If
  
                m_previousDEC = s_d
                m_movingDEC = True
                '*********** if a DEC move and single axis and currently performing RA move stop the RA move
                '*********** then carry on with the DEC move
                If ((m_movingRA) And (g_AxisRates.SingleAxis)) Then
                                        Me.CommandBlind "Q" & m_previousRA
                                        g_Util.WaitForMilliseconds 100
                                        Me.CommandBlind "Q" & m_previousRA
                                        m_movingRA = False
                        End If
                        
        End If  ' **********************of setup to perform a DEC move
         If (d_rateError > g_derrorMargin) And (Rate <> 0) Then
         g_Tracing.Trace ("Error in ScopeMoveAxis rate not acceptable ")
         Err.Raise SCODE_MOVE_RATE_ERROR, _
                    ERR_SOURCE, _
                    MSG_MOVE_RATE_ERROR
                    End If
            

    End Select
    
    g_Tracing.Trace ("Move Axis-> CommandBlind  " & s_r & s_d)
    Me.CommandBlind s_r & s_d
    
    If (Not m_movingRA) And (Not m_movingDEC) Then
    Me.Slewing = False
    Else
    Me.Slewing = True
    End If
    



Exit Sub



End Sub

'========================================================================
' ITelescope implementation

Private Sub ITelescope_AbortSlew()
    AbortSlew
End Sub

Private Property Get ITelescope_AlignmentMode() As AlignmentModes
    ITelescope_AlignmentMode = AlignmentMode
End Property

Private Property Get ITelescope_Altitude() As Double
    ITelescope_Altitude = Altitude
End Property

Private Property Get ITelescope_ApertureArea() As Double
ITelescope_ApertureArea = ApertureArea
End Property

Private Property Get ITelescope_ApertureDiameter() As Double
    ITelescope_ApertureDiameter = ApertureDiameter
End Property

Private Property Get ITelescope_AtHome() As Boolean
    ITelescope_AtHome = AtHome

End Property

Private Property Get ITelescope_AtPark() As Boolean
    ITelescope_AtPark = AtPark

End Property

Private Function ITelescope_AxisRates(ByVal Axis As TelescopeAxes) As IAxisRates
    '
    ' Note that this more or less "casts" our internal AxisRates
    ' object's interface to AxisRates.
    '
    Set ITelescope_AxisRates = AxisRates(Axis)
 
End Function

Private Property Get ITelescope_Azimuth() As Double
    ITelescope_Azimuth = Azimuth
End Property

Private Property Get ITelescope_CanFindHome() As Boolean
    ITelescope_CanFindHome = CanFindHome
End Property

Private Function ITelescope_CanMoveAxis(ByVal Axis As TelescopeAxes) As Boolean
    ITelescope_CanMoveAxis = CanMoveAxis(Axis)
End Function

Private Property Get ITelescope_CanPark() As Boolean
    ITelescope_CanPark = CanPark
End Property

Private Property Get ITelescope_CanPulseGuide() As Boolean
    ITelescope_CanPulseGuide = CanPulseGuide
End Property

Private Property Get ITelescope_CanSetDeclinationRate() As Boolean
    ITelescope_CanSetDeclinationRate = CanSetDeclinationRate
End Property

Private Property Get ITelescope_CanSetGuideRates() As Boolean
    ITelescope_CanSetGuideRates = CanSetGuideRates
End Property

Private Property Get ITelescope_CanSetPark() As Boolean
    ITelescope_CanSetPark = CanSetPark
End Property

Private Property Get ITelescope_CanSetPierSide() As Boolean
    ITelescope_CanSetPierSide = CanSetPierSide
 
End Property

Private Property Get ITelescope_CanSetRightAscensionRate() As Boolean
    ITelescope_CanSetRightAscensionRate = CanSetRightAscensionRate
End Property

Private Property Get ITelescope_CanSetTracking() As Boolean
    ITelescope_CanSetTracking = CanSetTracking
End Property

Private Property Get ITelescope_CanSlew() As Boolean
    ITelescope_CanSlew = CanSlew
End Property

Private Property Get ITelescope_CanSlewAltAz() As Boolean
    ITelescope_CanSlewAltAz = CanSlewAltAz
 
End Property

Private Property Get ITelescope_CanSlewAltAzAsync() As Boolean
    ITelescope_CanSlewAltAzAsync = CanSlewAltAzAsync
  
End Property

Private Property Get ITelescope_CanSlewAsync() As Boolean
    ITelescope_CanSlewAsync = CanSlewAsync
End Property

Private Property Get ITelescope_CanSync() As Boolean
    ITelescope_CanSync = CanSync
End Property

Private Property Get ITelescope_CanSyncAltAz() As Boolean
    ITelescope_CanSyncAltAz = CanSyncAltAz
End Property

Private Property Get ITelescope_CanUnpark() As Boolean
    ITelescope_CanUnpark = CanUnpark

End Property

Private Sub ITelescope_CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False)
    'CommandBlind Command, Raw
    CommandBlind Command
End Sub

Private Function ITelescope_CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean
   ' ITelescope_CommandBool = CommandBool(Command, Raw)
   ITelescope_CommandBool = CommandBool(Command)
End Function

Private Function ITelescope_CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String
   ' ITelescope_CommandString = CommandString(Command, Raw)
     ITelescope_CommandString = CommandString(Command)
End Function

Private Property Let ITelescope_Connected(ByVal RHS As Boolean)
    Connected = RHS
End Property

Private Property Get ITelescope_Connected() As Boolean
    ITelescope_Connected = Connected
End Property

Private Property Get ITelescope_Declination() As Double
    ITelescope_Declination = Declination
End Property

Private Property Let ITelescope_DeclinationRate(ByVal RHS As Double)
    DeclinationRate = RHS
End Property

Private Property Get ITelescope_DeclinationRate() As Double
    ITelescope_DeclinationRate = DeclinationRate
End Property

Private Property Get ITelescope_Description() As String
    ITelescope_Description = Description
End Property

Private Function ITelescope_DestinationSideOfPier(ByVal RightAscension As Double, ByVal Declination As Double) As PierSide
    ITelescope_DestinationSideOfPier = DestinationSideOfPier(RightAscension, Declination)
        
End Function

Private Property Let ITelescope_DoesRefraction(ByVal RHS As Boolean)
    DoesRefraction = RHS
 
End Property

Private Property Get ITelescope_DoesRefraction() As Boolean
    ITelescope_DoesRefraction = DoesRefraction
 
End Property

Private Property Get ITelescope_DriverInfo() As String
    ITelescope_DriverInfo = DriverInfo
End Property

Private Property Get ITelescope_DriverVersion() As String
    ITelescope_DriverVersion = DriverVersion
End Property

Private Property Get ITelescope_EquatorialSystem() As EquatorialCoordinateType
    ITelescope_EquatorialSystem = EquatorialSystem
End Property

Private Sub ITelescope_FindHome()
    FindHome
End Sub

Private Property Get ITelescope_FocalLength() As Double
    ITelescope_FocalLength = FocalLength
End Property

Private Property Let ITelescope_GuideRateDeclination(ByVal RHS As Double)
    GuideRateDeclination = RHS
       End Property

Private Property Get ITelescope_GuideRateDeclination() As Double
   ITelescope_GuideRateDeclination = GuideRateDeclination
     End Property

Private Property Let ITelescope_GuideRateRightAscension(ByVal RHS As Double)
    GuideRateRightAscension = RHS
     End Property

Private Property Get ITelescope_GuideRateRightAscension() As Double
    ITelescope_GuideRateRightAscension = GuideRateRightAscension
    End Property

Private Property Get ITelescope_InterfaceVersion() As Integer
    ITelescope_InterfaceVersion = InterfaceVersion
End Property

Private Property Get ITelescope_IsPulseGuiding() As Boolean
    ITelescope_IsPulseGuiding = IsPulseGuiding
End Property

Private Sub ITelescope_MoveAxis(ByVal Axis As TelescopeAxes, ByVal Rate As Double)
'g_Tracing.Trace ("In  Itelescope MoveAxis with Axis " & CStr(Axis) & " Rate " & CStr(Rate))
MoveAxis Axis, Rate
   
End Sub

Private Property Get ITelescope_Name() As String
    ITelescope_Name = Name
End Property

Private Sub ITelescope_Park()
    Park
End Sub

Private Sub ITelescope_PulseGuide(ByVal Direction As GuideDirections, ByVal Duration As Long)




    PulseGuide Direction, Duration
End Sub

Private Property Get ITelescope_RightAscension() As Double
    ITelescope_RightAscension = RightAscension
End Property

Private Property Let ITelescope_RightAscensionRate(ByVal RHS As Double)
    RightAscensionRate = RHS
End Property

Private Property Get ITelescope_RightAscensionRate() As Double
    ITelescope_RightAscensionRate = RightAscensionRate
End Property

Private Sub ITelescope_SetPark()
    SetPark
End Sub

Private Sub ITelescope_SetupDialog()
    SetupDialog
End Sub

Private Property Let ITelescope_SideOfPier(ByVal RHS As PierSide)
    SideOfPier = RHS
End Property

Private Property Get ITelescope_SideOfPier() As PierSide
    ITelescope_SideOfPier = SideOfPier
End Property

Private Property Get ITelescope_SiderealTime() As Double
    ITelescope_SiderealTime = SiderealTime
End Property

Private Property Let ITelescope_SiteElevation(ByVal RHS As Double)
    SiteElevation = RHS
End Property

Private Property Get ITelescope_SiteElevation() As Double
    ITelescope_SiteElevation = SiteElevation
End Property

Private Property Let ITelescope_SiteLatitude(ByVal RHS As Double)
    SiteLatitude = RHS
End Property

Private Property Get ITelescope_SiteLatitude() As Double
    ITelescope_SiteLatitude = SiteLatitude
End Property

Private Property Let ITelescope_SiteLongitude(ByVal RHS As Double)
    SiteLongitude = RHS
End Property

Private Property Get ITelescope_SiteLongitude() As Double
    ITelescope_SiteLongitude = SiteLongitude
End Property

Private Property Get ITelescope_Slewing() As Boolean
    ITelescope_Slewing = Slewing
End Property

Private Property Let ITelescope_SlewSettleTime(ByVal RHS As Integer)
    SlewSettleTime = RHS
End Property

Private Property Get ITelescope_SlewSettleTime() As Integer
    ITelescope_SlewSettleTime = SlewSettleTime
End Property

Private Sub ITelescope_SlewToAltAz(ByVal Azimuth As Double, ByVal Altitude As Double)
    SlewToAltAz Azimuth, Altitude
End Sub

Private Sub ITelescope_SlewToAltAzAsync(ByVal Azimuth As Double, ByVal Altitude As Double)
    SlewToAltAzAsync Azimuth, Altitude
End Sub

Private Sub ITelescope_SlewToCoordinates(ByVal RightAscension As Double, ByVal Declination As Double)
    SlewToCoordinates RightAscension, Declination
End Sub

Private Sub ITelescope_SlewToCoordinatesAsync(ByVal RightAscension As Double, ByVal Declination As Double)
    SlewToCoordinatesAsync RightAscension, Declination
End Sub

Private Sub ITelescope_SlewToTarget()
    SlewToTarget
End Sub

Private Sub ITelescope_SlewToTargetAsync()
    SlewToTargetAsync
End Sub

Private Sub ITelescope_SyncToAltAz(ByVal Azimuth As Double, ByVal Altitude As Double)
    SyncToAltAz Azimuth, Altitude
End Sub

Private Sub ITelescope_SyncToCoordinates(ByVal RightAscension As Double, ByVal Declination As Double)
    SyncToCoordinates RightAscension, Declination
End Sub

Private Sub ITelescope_SyncToTarget()
    SyncToTarget
End Sub

Private Property Let ITelescope_TargetDeclination(ByVal RHS As Double)
    TargetDeclination = RHS
End Property

Private Property Get ITelescope_TargetDeclination() As Double
    ITelescope_TargetDeclination = TargetDeclination
End Property

Private Property Let ITelescope_TargetRightAscension(ByVal RHS As Double)
    TargetRightAscension = RHS
End Property

Private Property Get ITelescope_TargetRightAscension() As Double
    ITelescope_TargetRightAscension = TargetRightAscension
End Property

Private Property Let ITelescope_Tracking(ByVal RHS As Boolean)
    Tracking = RHS
End Property

Private Property Get ITelescope_Tracking() As Boolean
    ITelescope_Tracking = Tracking
End Property

Private Property Let ITelescope_TrackingRate(ByVal RHS As DriveRates)
    TrackingRate = RHS
End Property

Private Property Get ITelescope_TrackingRate() As DriveRates
    ITelescope_TrackingRate = TrackingRate
End Property
Private Property Get ITelescope_TrackingRates() As ITrackingRates
    '
    ' Note that this more or less "casts" our internal TrackingRates
    ' object's interface to ITrackingRates.
    '
    Set ITelescope_TrackingRates = TrackingRates
End Property


Private Sub ITelescope_Unpark()
    Unpark
End Sub

Private Property Let ITelescope_UTCDate(ByVal RHS As Date)
    UTCDate = RHS
End Property

Private Property Get ITelescope_UTCDate() As Date
    ITelescope_UTCDate = UTCDate
End Property

Private Sub DoSync()
   Dim buf As String
    
   ' Allow old and new responses. Old responds "0" or "1", new
   ' responds "Coordinates matched.      ". So a "0" is a failure,
   ' as well as "".
   '
   If g_ROMVersion.Version > 205 Then
    buf = Me.CommandString("CM")
    If buf = "0" Or buf = "" Then
       Err.Raise SCODE_SYNC_FAILED, ERR_SOURCE, MSG_SYNC_FAILED
    End If
    Else
      If Not Me.CommandBool("CM") Then
    
       Err.Raise SCODE_SYNC_FAILED, ERR_SOURCE, MSG_SYNC_FAILED
      End If

    End If
            
End Sub
