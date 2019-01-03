VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton moveaxis_file 
      Caption         =   "moveaxis from file"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton serial_file 
      Caption         =   "serial from file"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PulseGuide"
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "setup"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Scope
'Private m_Serial As DriverHelper.Serial






Private Sub Command1_Click()
    Dim ID As String
    Dim Chsr As DriverHelper.Chooser
    
    Set Chsr = New DriverHelper.Chooser
    Chsr.DeviceType = "Telescope"
    '  you can remember the scope ID and set it to give you the last one used
    ID = Chsr.Choose(ID)
    Set m_Scope = CreateObject(ID)
    ' connect to the scope
   m_Scope.Connected = True
''MyDate = "October 19, 1962"   ' Define date.

''MyShortDate = CDate(MyDate)   ' Convert to Date data type.

''MyTime = "4:35:47 AM"         ' Define time.

''MyShortTime = CDate(MyTime)   ' Convert to Date data type.
 '''     Dim fred As Double
 ''     fred = CDate(MyTime)
 '    'MySiderealTime = CDate(Me.CommandString("GS"))
 ''     SS2KSiderealTime = m_scope.SiderealTime
 

 'Dim foo2 As AxisRates
 Dim foo3 As Double
   'Set
   'If m_Scope.AxisRates(axisPrimary).Count > 0 Then
  ' foo3 = m_Scope.AxisRates(axisPrimary).Item(1).Minimum
      'foo3 = foo2.Item(1).Minimum
  ' Form2.Text1 = foo3


End Sub
Public Function CommandString(ByVal Command As String) As String
   Dim buf As String

   'Make sure we are connected
   CheckConnected
    
   m_Serial.ClearBuffers                           ' Clear remaining junk in buffer
   m_Serial.Transmit "#:" & Command & "#"
   buf = m_Serial.ReceiveTerminated("#")
   If buf <> "" Then                   ' Overflow protection
       CommandString = Left$(buf, Len(buf) - 1)   ' Strip '#'
   Else
       CommandString = ""
   End If
   
End Function
Private Sub CheckConnected()

    If Not m_Serial.Connected Then _
        Err.Raise SCODE_NOT_CONNECTED, _
                    ERR_SOURCE, _
                    MSG_NOT_CONNECTED
End Sub

Private Sub Command2_Click()
'Dim foo As AxisRates

    Set Chsr = New DriverHelper.Chooser
    Chsr.DeviceType = "Telescope"
    '  you can remember the scope ID and set it to give you the last one used
    'ID = Chsr.Choose(ID)
    ID = "SS2K.Telescope"
    Set m_Scope = CreateObject(ID)
 '   m_Scope.SetupDialog
    ' connect to the scope
   m_Scope.Connected = True

foo = m_Scope.PulseGuide(0, 200)

End Sub





Private Sub Form_Load()
   ' Dim ID As String
   ' Dim Chsr As DriverHelper.Chooser
    
   ' Set Chsr = New DriverHelper.Chooser
   ' Chsr.DeviceType = "Telescope"
    '  you can remember the scope ID and set it to give you the last one used
  '  ID = Chsr.Choose(ID)
   ' Set m_Scope = CreateObject(ID)

End Sub

Private Sub moveaxis_file_Click()

    Dim ID As String
    Dim Chsr As DriverHelper.Chooser
  Const ForReading = 1, ForWriting = 2, ForAppending = 3

    
    Set Chsr = New DriverHelper.Chooser
    Chsr.DeviceType = "Telescope"
    '  you can remember the scope ID and set it to give you the last one used
    'ID = Chsr.Choose(ID)
    ID = "SS2K.Telescope"
    Set m_Scope = CreateObject(ID)
    ' connect to the scope
   m_Scope.Connected = True


 'Dim foo2 As AxisRates
 Dim foo3 As Double

   Dim b_csgr As Boolean
Set fso = CreateObject("Scripting.FileSystemObject")
Set fR = fso.OpenTextFile("D:\\Ascom\\v5.0_fixes\\5.1.5d_moveaxis\\working_b_filtered_4.txt", ForReading)
Dim doString, tempString, startTime, fileTime, lineFileTime, fileSecs, lineFileSecs
 lineR = fR.ReadLine()
Dim d, h, m, s, tStart, rightNow, s_fileSecs0, s_fileSecs1, i_moveAxisPos
'rightNow = Now ' not needed

i_fileSecs0 = (3600 * Mid(lineR, 15, 2)) + (60 * Mid(lineR, 18, 2)) + Mid(lineR, 21, 2)
fR.Close
Set fR = fso.OpenTextFile("D:\\Ascom\\v5.0_fixes\\5.1.5d_moveaxis\\working_b_filtered_4.txt", ForReading)
'******************************

' ********** Run through the text file look for In ScopeMoveAxis line
' *********** then wait for the required time before sending the command
 While (Not fR.AtEndOfStream)
  lineR = fR.ReadLine()
i_fileSecs1 = (3600 * Mid(lineR, 15, 2)) + (60 * Mid(lineR, 18, 2)) + Mid(lineR, 21, 2)
 i_moveAxisPos = InStr(1, lineR, "In ScopeMoveAxis", vbTextCompare)
 If i_moveAxisPos > 0 Then
 i_moveDirection = Mid(lineR, i_moveAxisPos + 33, 1)
 s_moveAmount = Mid(lineR, i_moveAxisPos + 38)
 
 '*now calculate delay from the last line read held in s_fileSecs0
 i_delay = i_fileSecs1 - i_fileSecs0
 
 Pause (i_delay)
m_Scope.MoveAxis i_moveDirection, s_moveAmount
i_fileSecs0 = i_fileSecs1
 End If
 
 
Wend

'm_Scope.CommandBlind Me
End Sub

Private Sub serial_file_Click()
    Dim ID As String
    Dim Chsr As DriverHelper.Chooser
  Const ForReading = 1, ForWriting = 2, ForAppending = 3

    
    Set Chsr = New DriverHelper.Chooser
    Chsr.DeviceType = "Telescope"
    '  you can remember the scope ID and set it to give you the last one used
    'ID = Chsr.Choose(ID)
    ID = "SS2K.Telescope"
    Set m_Scope = CreateObject(ID)
    m_Scope.SetupDialog
    ' connect to the scope
   m_Scope.Connected = True


 'Dim foo2 As AxisRates
 Dim foo3 As Double

   Dim b_csgr As Boolean
Set fso = CreateObject("Scripting.FileSystemObject")
Set fR = fso.OpenTextFile("D:\\Ascom\\v5.0_fixes\\5.1.5d_moveaxis\\working_b_filtered_4.txt", ForReading)

Set m_Scope = Nothing
End Sub

' from http://www.visualbasic.happycodings.com/API_and_Miscellaneous/code23.html
' You can delay execution of your code for a specific time interval
' by using the Timer function. Increments such as .25 or .5 can be
' used as well.

' To use the Timer function to pause for a number of seconds,
' store the value of Timer in a variable. Then use a loop to wait
' until the Timer returns a specified number of seconds greater than
' the stored value. If the delay loop will execute when midnight
' passes, compensate by reducing the starting Timer value by the
' number of seconds in a day (24 hours * 60 minutes * 60 seconds).
' Calling DoEvents from within the loop allows events to be
' processed during the delay.

' Drop this sub in the appropriate form:

Sub Pause(ByVal nSecond As Single)
Dim t0 As Single
Dim dummy As Integer
        
        t0 = Timer
        
        Do While Timer - t0 < nSecond
                
                dummy = DoEvents()
                
                ' If we cross midnight, back up one day
                If Timer < t0 Then
                        t0 = t0 - 24 * 60 * 60 ' or t0 = t0 - 86400

                End If
        Loop

End Sub
