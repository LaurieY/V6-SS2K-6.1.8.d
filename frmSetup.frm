VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ASCOM Vixen SS2K Driver Setup"
   ClientHeight    =   3390
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7350
   Begin VB.PictureBox picASCOM 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   135
      MouseIcon       =   "frmSetup.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmSetup.frx":0152
      ScaleHeight     =   840
      ScaleWidth      =   720
      TabIndex        =   2
      ToolTipText     =   "Click to go to the ASCOM web site"
      Top             =   2445
      Width           =   720
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2460
      TabIndex        =   1
      Top             =   2715
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   1200
      TabIndex        =   0
      Top             =   2715
      Width           =   1095
   End
   Begin TabDlg.SSTab ppSetup 
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   1
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   3836
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&Communication"
      TabPicture(0)   =   "frmSetup.frx":1016
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbSpeed"
      Tab(0).Control(1)=   "lbPort"
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(4)=   "Label1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Version"
      TabPicture(1)   =   "frmSetup.frx":1032
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "radVersion(210)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "radVersion(209)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "radVersion(208)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "radVersion(207)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "radVersion(206)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "radVersion(205)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "radVersion(204)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "radVersion(203)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "radVersion(202)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "radVersion(201)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "&Scope"
      TabPicture(2)   =   "frmSetup.frx":104E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "edtSlewSettleTime"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&About"
      TabPicture(3)   =   "frmSetup.frx":106A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblAbout"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Site&Elevation"
      TabPicture(4)   =   "frmSetup.frx":1086
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label3"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "edtSiteElevation"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "1-Star Alignment"
      TabPicture(5)   =   "frmSetup.frx":10A2
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label7"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cbox1Star"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin VB.CheckBox cbox1Star 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3840
         TabIndex        =   26
         ToolTipText     =   "Force 1-Star Alignment on Syncs"
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox lbPort 
         Height          =   315
         ItemData        =   "frmSetup.frx":10BE
         Left            =   -73800
         List            =   "frmSetup.frx":10F9
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   540
         Width           =   945
      End
      Begin VB.ComboBox lbSpeed 
         Height          =   315
         ItemData        =   "frmSetup.frx":1174
         Left            =   -73800
         List            =   "frmSetup.frx":11B7
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   900
         Width           =   945
      End
      Begin VB.TextBox edtSlewSettleTime 
         Height          =   285
         Left            =   -73440
         TabIndex        =   15
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox edtSiteElevation 
         Height          =   285
         Left            =   -73000
         TabIndex        =   14
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.01"
         Height          =   255
         Index           =   201
         Left            =   -74600
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.02"
         Height          =   255
         Index           =   202
         Left            =   -74600
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.03"
         Height          =   255
         Index           =   203
         Left            =   -74600
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.04"
         Height          =   255
         Index           =   204
         Left            =   -73600
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.05"
         Height          =   255
         Index           =   205
         Left            =   -73600
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.06"
         Height          =   255
         Index           =   206
         Left            =   -73600
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.07"
         Height          =   255
         Index           =   207
         Left            =   -72600
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.08"
         Height          =   255
         Index           =   208
         Left            =   -72600
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.09"
         Height          =   255
         Index           =   209
         Left            =   -72600
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton radVersion 
         Caption         =   "2.10"
         Height          =   255
         Index           =   210
         Left            =   -71600
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "1-Star Alignment d"
         Height          =   495
         Left            =   1680
         TabIndex        =   27
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Serial Port:"
         Height          =   225
         Left            =   -74760
         TabIndex        =   25
         Top             =   585
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "s"
         Height          =   255
         Left            =   -72840
         TabIndex        =   23
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Slew settle time:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "bps"
         Height          =   255
         Left            =   -72720
         TabIndex        =   21
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblAbout 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   20
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Site Elevation"
         Height          =   375
         Left            =   -74520
         TabIndex        =   19
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "Metres"
         Height          =   255
         Left            =   -71880
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'   FRMSETUP.FRM
'   =============
'
' Setup form for ASCOM Vixen SkySensor 2000-PCtelescope driver
'
' Written:                             22-Aug-00   Robert B. Denny <rdenny@dc3.com>
' Adapted to support SkySensor2000-PC  14-Dec-00   Arne Danielsen
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 22-Aug-00 rbd     Initial edit
' 17-Dec-00 ad      Expanded setup to fit SkySensor2000-PC
' 24-Jan-01 ad      Updated code to confirm with Beta 2 of the SDK
' 30-Dec-01 ad      Added SkySensor ROM v2.06 as option. No new
'                   functionality available with this option.
' 27-Jul-02 rbd     ASCOM logo and hot link
' 24-Nov-04 rbd     4.0.2 - COM Ports to 16
'---------------------------------------------------------------------

Option Explicit


Private Const SWP_NOSIZE            As Long = &H1
Private Const SWP_NOMOVE            As Long = &H2
Private Const SWP_NOZORDER          As Long = &H4
Private Const SWP_NOREDRAW          As Long = &H8
Private Const SWP_NOACTIVATE        As Long = &H10
Private Const SWP_FRAMECHANGED      As Long = &H20
Private Const SWP_SHOWWINDOW        As Long = &H40
Private Const SWP_HIDEWINDOW        As Long = &H80
Private Const SWP_NOCOPYBITS        As Long = &H100
Private Const SWP_NOOWNERZORDER     As Long = &H200
Private Const SWP_NOSENDCHANGING    As Long = &H400

Private Const SWP_DRAWFRAME         As Long = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION      As Long = SWP_NOOWNERZORDER

Private Const HWND_TOP              As Long = 0
Private Const HWND_BOTTOM           As Long = 1
Private Const HWND_TOPMOST          As Long = -1
Private Const HWND_NOTOPMOST        As Long = -2

Private Const SW_SHOWNORMAL         As Long = 1

Private Declare Function SetWindowPos Lib "user32.dll" ( _
                ByVal hWnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal X As Long, _
                ByVal y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal uFLags As Long) As Long

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
                ByVal hWnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
   Dim nIndex As Integer
   Dim strSpeed As String
   'Fill controls with settings
   
   'Communication settings
   lbPort.ListIndex = g_Communication.Port - 1  ' Select current port
   
   strSpeed = CStr(g_Communication.Speed)
   For nIndex = 0 To lbSpeed.ListCount
      If strSpeed = lbSpeed.List(nIndex) Then
         lbSpeed.ListIndex = nIndex
         Exit For
      End If
   Next nIndex
    
   'Version settings
   radVersion.Item(g_ROMVersion.Version) = True
   
   'Scope settings
   edtSlewSettleTime.Text = CStr(g_Scope.SlewSettleTime)
   edtSiteElevation.Text = CStr(g_SiteElevation.SiteElevation)
   'Copyright information
   lblAbout.Caption = App.FileDescription & " " & _
                App.Major & "." & App.Minor & "." & App.Revision
    If App.CompanyName <> "" Then _
        lblAbout.Caption = lblAbout.Caption & vbCrLf & App.CompanyName
        If g_OneStar.AStar = True Then
            cbox1Star.Value = 1
            Else
            cbox1Star.Value = 0
            End If
 
   ' Assure window pops up on top of others.
   '
   SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, (SWP_NOMOVE + SWP_NOSIZE)

End Sub

Private Sub cmdOK_Click()
   Dim nIndex As Integer
   
   If Validate = True Then
      'Retrieve settings from controls
      'Communication settings
      g_Communication.Port = lbPort.ListIndex + 1
      g_Communication.Speed = CInt(lbSpeed.List(lbSpeed.ListIndex))
      
      'Version settings
      For nIndex = radVersion.LBound To radVersion.UBound
        If radVersion.Item(nIndex) = True Then
           g_ROMVersion.Version = nIndex
           Exit For
        End If
      Next nIndex
      
      'Scope settings
      g_Scope.SlewSettleTime = CInt(edtSlewSettleTime.Text)
      g_SiteElevation.SiteElevation = CDbl(edtSiteElevation.Text)
      g_Communication.SaveSettings       'Store communcation settings to registry
      g_Scope.SaveSettings               'Store scope settings to registry
      g_ROMVersion.SaveSettings          'Store version settings to registry
      g_SiteElevation.SaveSettings
      
      
      If cbox1Star.Value = 1 Then
       g_OneStar.AStar = True
       Else
        g_OneStar.AStar = False
    End If
        '******************************
        '**************g_OneStar.AStar = True
        
        g_OneStar.SaveSettings
        
        
      Me.Hide
   End If
End Sub

Private Sub cmdCancel_Click()

    Me.Hide
    
End Sub


Private Sub picASCOM_Click()
    Dim z As Long

    z = ShellExecute(0, "Open", "http://ASCOM-Standards.org/", 0, 0, SW_SHOWNORMAL)
    If (z > 0) And (z <= 32) Then
        MsgBox _
            "It doesn't appear that you have a web browser installed " & _
            "on your system.", (vbOKOnly + vbExclamation + vbMsgBoxSetForeground), ERR_SOURCE
        Exit Sub
    End If
End Sub

Private Function Validate() As Boolean
   Dim nSlewSettleTime As Integer
   Validate = True
   nSlewSettleTime = CInt(edtSlewSettleTime.Text)
   
   If nSlewSettleTime < 0 Or nSlewSettleTime > 60 Then
      MsgBox ("Slew settle time must have a value between 0 and 60")
      ppSetup.Tab = 2
      edtSlewSettleTime.SetFocus
      Validate = False
   End If
End Function

