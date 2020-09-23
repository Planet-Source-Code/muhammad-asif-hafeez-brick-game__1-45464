VERSION 5.00
Begin VB.Form Help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BRICK_GAME Help"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton okbttn 
      Caption         =   "Ok!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2047
      TabIndex        =   8
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BRICK_GAME is a very simple game to play:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "To move the pieces left and right, use the arrow keys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "To rotate the pieces, press either the up arrow key or return"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   4380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "To make the pieces fall right to the bottom, press space bar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   4290
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "To pause and unpause the game, press either the F3 key"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   4125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To make the pieces fall faster, press the down arrow key"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "faster or slower, or use the slider."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   2460
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "To change the speed press F5 or F6 to make the blocks fall"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   4230
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call okbttn_Click
End Sub

Private Sub Label3_Click()

End Sub

Private Sub okbttn_Click()
Unload Help
End Sub
