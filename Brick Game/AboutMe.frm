VERSION 5.00
Begin VB.Form AboutMe 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2655
   ClientLeft      =   2310
   ClientTop       =   1620
   ClientWidth     =   5040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1832.528
   ScaleMode       =   0  'User
   ScaleWidth      =   4732.821
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "AboutMe.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "2000-CE-195."
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
      Left            =   1200
      TabIndex        =   8
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "E-mail:"
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
      Left            =   960
      TabIndex        =   7
      Top             =   1680
      Width           =   570
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Please e-mail me any comments or questions!"
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
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   3900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "asif_ssuet@yahoo.com or mhafeez@ssuet.edu.pk"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   3630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sir Syed University of Engineering and Technology."
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
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   3675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Muhammad Asif Hafeez."
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
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Programming and Graphics: "
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
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BRICK_GAME ver 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1980
   End
End
Attribute VB_Name = "AboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
If Game_End = False Then
    Brick.tmrFall.Enabled = True
End If
Unload AboutMe
End Sub
