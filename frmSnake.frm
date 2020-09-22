VERSION 5.00
Begin VB.Form frmSnake 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake"
   ClientHeight    =   4140
   ClientLeft      =   2790
   ClientTop       =   1860
   ClientWidth     =   4680
   Icon            =   "frmSnake.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4680
   Begin VB.TextBox txtScore 
      Enabled         =   0   'False
      Height          =   345
      Left            =   3450
      TabIndex        =   3
      Text            =   "0"
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   300
      Left            =   570
      Top             =   3660
   End
   Begin VB.PictureBox picSnakeBox 
      Height          =   3420
      Left            =   150
      ScaleHeight     =   3360
      ScaleWidth      =   4320
      TabIndex        =   0
      Top             =   150
      Width           =   4380
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   60
         Left            =   1110
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   59
         Left            =   1050
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   58
         Left            =   1050
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   57
         Left            =   1020
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   56
         Left            =   1050
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   55
         Left            =   1020
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   54
         Left            =   1020
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   53
         Left            =   1050
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   52
         Left            =   1050
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   51
         Left            =   1020
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   50
         Left            =   1020
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   49
         Left            =   990
         Top             =   1020
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   48
         Left            =   990
         Top             =   1020
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   47
         Left            =   990
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   46
         Left            =   990
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   45
         Left            =   990
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   44
         Left            =   990
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   43
         Left            =   990
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   42
         Left            =   990
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   41
         Left            =   990
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   40
         Left            =   990
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   39
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   38
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   37
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   36
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   35
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   34
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   33
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   32
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   31
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   30
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   29
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   28
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   27
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   26
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   25
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   24
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   23
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   22
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   21
         Left            =   1020
         Top             =   990
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   20
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   19
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   18
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   17
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   16
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   15
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   14
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   13
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   12
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   11
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   10
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   9
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   8
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   7
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   6
         Left            =   1000
         Top             =   1000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   5
         Left            =   60
         Top             =   30
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   4
         Left            =   360
         Top             =   30
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   3
         Left            =   660
         Top             =   30
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00FF80FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   2
         Left            =   960
         Top             =   30
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillColor       =   &H00C0C000&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   1
         Left            =   1260
         Top             =   30
         Width           =   300
      End
      Begin VB.Shape shpLink 
         BorderColor     =   &H80000000&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   0
         Left            =   1560
         Top             =   30
         Width           =   300
      End
      Begin VB.Label lblFood 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2130
         TabIndex        =   1
         Top             =   1470
         Width           =   300
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Score :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2190
      TabIndex        =   4
      Top             =   3690
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Paused"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1530
      TabIndex        =   2
      Top             =   1290
      Width           =   1575
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play/Restart"
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "&Level"
         Begin VB.Menu mnu1 
            Caption         =   "1 - Easy"
         End
         Begin VB.Menu mn2 
            Caption         =   "2 "
         End
         Begin VB.Menu mnu3 
            Caption         =   "3"
         End
         Begin VB.Menu mn4 
            Caption         =   "4"
         End
         Begin VB.Menu mnu5 
            Caption         =   "5 - Difficult"
         End
      End
      Begin VB.Menu mnuPause 
         Caption         =   "P&ause"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuScore 
         Caption         =   "&Score"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpMe 
         Caption         =   "&Help"
      End
   End
End
Attribute VB_Name = "frmSnake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare all variables
'note better to use draw methods
'than controls as in this program
Dim BaseX As Integer ' left static variable
Dim BaseY As Integer ' top static variable
Dim counter As Integer

Dim Direction As Integer
Dim FrontDirection As Integer
'1= right
'2= left
'3= up
'4 = down
Dim NoOfBlocks As Integer
Dim response As Integer
Dim score As Integer

Private Sub cmdPlay_Click()
tmrSpeed.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
If KeyCode = vbKeyLeft Then
  If FrontDirection <> 1 Then
   FrontDirection = 2
   
  End If
   ElseIf KeyCode = vbKeyRight Then
    If FrontDirection <> 2 Then
     FrontDirection = 1
     
    End If
     ElseIf KeyCode = vbKeyUp Then
     If FrontDirection <> 4 Then
      FrontDirection = 3
      
     End If
     ElseIf KeyCode = vbKeyDown Then
     If FrontDirection <> 3 Then
        FrontDirection = 4
        
     End If
 End If
 
 End Sub

Private Sub Form_Load()
frmSnake.Show
initialise
End Sub

Private Sub mn2_Click()
response = MsgBox("Restart game on a different level?", vbOKCancel + vbQuestion, "Change Level?")
If response = vbOK Then
 Unload frmSnake
 Load frmSnake
   tmrSpeed.Interval = 450
End If

End Sub

Private Sub mn4_Click()
response = MsgBox("Restart game on a different level?", vbOKCancel + vbQuestion, "Change Level?")
If response = vbOK Then
 Unload frmSnake
 Load frmSnake
  initialise
  tmrSpeed.Interval = 200
End If

End Sub

Private Sub mnu1_Click()
response = MsgBox("Restart game on a different level?", vbOKCancel + vbQuestion, "Change Level?")
If response = vbOK Then
 Unload frmSnake
 Load frmSnake
  initialise
  tmrSpeed.Interval = 600
End If
End Sub

Private Sub mnu3_Click()
response = MsgBox("Restart game on a different level?", vbOKCancel + vbQuestion, "Change Level?")
If response = vbOK Then
 Unload frmSnake
 Load frmSnake
   tmrSpeed.Interval = 300
End If

End Sub

Private Sub mnu5_Click()

response = MsgBox("Restart game on a different level?", vbOKCancel + vbQuestion, "Change Level?")

If response = vbOK Then
 Unload frmSnake
  Load frmSnake
  tmrSpeed.Interval = 150
End If


End Sub

Private Sub mnuExit_Click()
Dim response As Integer
response = MsgBox("Are you sure that want to exit?", vbOKCancel + vbQuestion, "Exit Snake?")
If response = vbOK Then
 End
End If
End Sub

Private Sub mnuHelpMe_Click()
mnuPause.Checked = True
  tmrSpeed.Enabled = False
  picSnakeBox.Visible = False
frmSnake.Hide
Load frmHelp
End Sub

Private Sub mnuPause_Click()

If mnuPause.Checked = False Then
  mnuPause.Checked = True
  tmrSpeed.Enabled = False
  picSnakeBox.Visible = False
Else
 mnuPause.Checked = False
 tmrSpeed.Enabled = True
 picSnakeBox.Visible = True
End If

End Sub

Private Sub mnuPlay_Click()
Unload frmSnake
Load frmSnake
End Sub

Private Sub tmrSpeed_Timer()

BaseX = 300
BaseY = 300
Dim response As Integer
'need somthing to detect the place where change of direction has taken place

shpLink(0).Tag = FrontDirection
bumpself
For counter = 0 To (NoOfBlocks)
   
    Direction = Val(shpLink(counter).Tag)
   
  Select Case Direction 'move one block in the array
 
  Case Is = 1 'moving  right
    If shpLink(counter).Left < picSnakeBox.ScaleWidth - 360 Then
        
        shpLink(counter).Left = shpLink(counter).Left + BaseX
       
       Else
        'loose stuff goes here
       tmrSpeed.Enabled = False
       response = MsgBox("YOU LOOSE", vbOKOnly + vbExclamation, "LOSER")
       Exit For
     End If
  Case Is = 2 'moving left
     If shpLink(counter).Left > 60 Then
       shpLink(counter).Left = shpLink(counter).Left - BaseX
       
     Else
        'loose stuff goes here
       tmrSpeed.Enabled = False
       response = MsgBox("YOU LOOSE", vbOKOnly + vbExclamation, "LOSER")
       Exit For
     End If
     
  Case Is = 3 'moving up
      If shpLink(counter).Top > 60 Then
        shpLink(counter).Top = shpLink(counter).Top - BaseY
        
      Else
        'loose stuff goes here
       tmrSpeed.Enabled = False
       response = MsgBox("YOU LOOSE", vbOKOnly + vbExclamation, "LOSER")
       Exit For
      End If
  Case Is = 4 'moving down
    If shpLink(counter).Top < picSnakeBox.ScaleHeight - 360 Then
     shpLink(counter).Top = shpLink(counter).Top + BaseY
     
    Else
        'loose stuff goes here
       tmrSpeed.Enabled = False
       response = MsgBox("YOU LOOSE", vbOKOnly + vbExclamation, "LOSER")
       Exit For
    End If
 End Select

'when the food is eaten here
If shpLink(0).Top = lblFood.Top And shpLink(0).Left = lblFood.Left Then
  
  Beep
  Beep
  LocateFood 'relocates the food
  'add a block to the snake here
  
  shpLink(NoOfBlocks + 1).Top = shpLink(NoOfBlocks).Top
  shpLink(NoOfBlocks + 1).Left = shpLink(NoOfBlocks).Left
   
      
  NoOfBlocks = NoOfBlocks + 1
  shpLink(NoOfBlocks).Visible = True
  
 score = score + (4000 / tmrSpeed.Interval)
 txtScore.Text = Str(score)

End If

Next counter

For counter = (NoOfBlocks) To 2 Step -1
  shpLink(counter).Tag = shpLink(counter - 1).Tag
Next

shpLink(1).Tag = FrontDirection

End Sub

Private Sub initialise()
tmrSpeed.Enabled = True

'set initial directions

FrontDirection = 1
score = 0
For counter = 0 To 5
 shpLink(counter).Tag = 1
Next counter

NoOfBlocks = 5 'including 0 = 6

Call LocateFood
Randomize

For counter = 1 To 60
 shpLink(counter).FillColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Next counter

mnuPause.Checked = False

End Sub
 

Private Sub LocateFood()
Dim RowNumber As Integer
Dim ColomnNumber As Integer
Dim tryflag As Integer
Dim rowxvalue As Integer
Dim rowyvalue As Integer
Dim counter2 As Integer

Do
 RowNumber = Rnd * 10
 ColomnNumber = Rnd * 13
 rowxvalue = RowNumber * 300 + 30
 rowyvalue = RowNumber * 300 + 60

 For counter2 = 0 To NoOfBlocks
  If rowxvalue = shpLink(counter2).Top And rowyvalue = shpLink(counter2).Left Then
    tryflag = 1
   Exit For
  Else: tryflag = 0
  End If
 Next counter2
Loop While tryflag = 1

lblFood.Top = RowNumber * 300 + 30
lblFood.Left = ColomnNumber * 300 + 60

End Sub


Public Sub bumpself()
'checks that next move is not bumping into into itself
Dim response  As Integer
Dim counter1 As Integer

 Select Case FrontDirection
   Case Is = 1 ' moving right
    
    For counter1 = 1 To NoOfBlocks
      If (shpLink(0).Top = shpLink(counter1).Top) And (shpLink(0).Left + 300 = shpLink(counter1).Left) Then
        Beep
        response = MsgBox("You Loose", vbOKOnly + vbExclamation, "LOOSER")
        tmrSpeed.Enabled = False
        Exit For
        Exit For
      End If
    Next counter1
   
   Case Is = 2 ' moving left
     For counter1 = 1 To NoOfBlocks
      If (shpLink(0).Top = shpLink(counter1).Top) And (shpLink(0).Left - 300 = shpLink(counter1).Left) Then
        Beep
        response = MsgBox("You Loose", vbOKOnly + vbExclamation, "LOOSER")
        tmrSpeed.Enabled = False
        Exit For
        Exit For
      End If
    Next counter1
   Case Is = 3 ' moving up
     For counter1 = 1 To NoOfBlocks
      If (shpLink(0).Top - 300 = shpLink(counter1).Top) And (shpLink(0).Left = shpLink(counter1).Left) Then
        Beep
        response = MsgBox("You Loose", vbOKOnly + vbExclamation, "LOOSER")
        tmrSpeed.Enabled = False
        Exit For
        Exit For
      End If
    Next counter1
   Case Is = 4 ' moving down
    For counter1 = 1 To NoOfBlocks
      If (shpLink(0).Top + 300 = shpLink(counter1).Top) And (shpLink(0).Left = shpLink(counter1).Left) Then
        Beep
        response = MsgBox("You Loose", vbOKOnly + vbExclamation, "LOOSER")
        tmrSpeed.Enabled = False
        Exit For
        Exit For
      End If
    Next counter1
 End Select
For counter1 = 1 To NoOfBlocks
 If (shpLink(0).Top = shpLink(counter1).Top) And (shpLink(0).Left = shpLink(counter1).Left) Then
  Beep
  response = MsgBox("You Loose", vbOKOnly + vbExclamation, "LOOSER")
  tmrSpeed.Enabled = False
  Exit For
  Exit For
 End If
Next counter1
End Sub
