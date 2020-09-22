VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake"
   ClientHeight    =   4245
   ClientLeft      =   2625
   ClientTop       =   1560
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00008000&
      Height          =   1515
      Left            =   870
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmintro.frx":0000
      Top             =   1320
      Width           =   3075
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play Snake"
      Height          =   615
      Left            =   1650
      TabIndex        =   0
      Top             =   3270
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "SNAKE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   645
      Left            =   660
      TabIndex        =   1
      Top             =   540
      Width           =   3435
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlay_Click()
Unload frmIntro
Load frmSnake
End Sub
