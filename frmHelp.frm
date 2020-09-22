VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake Help"
   ClientHeight    =   4590
   ClientLeft      =   2580
   ClientTop       =   1485
   ClientWidth     =   4680
   ControlBox      =   0   'False
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2355
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmHelp.frx":0000
      Top             =   960
      Width           =   4065
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Ok"
      Height          =   555
      Left            =   1830
      TabIndex        =   0
      Top             =   3600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "How to play snake :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   960
      TabIndex        =   1
      Top             =   300
      Width           =   2775
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
Unload frmHelp
frmSnake.Show
End Sub

Private Sub Form_Load()
frmHelp.Show
End Sub
