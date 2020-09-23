VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Brick 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BRICK GAME"
   ClientHeight    =   5010
   ClientLeft      =   3825
   ClientTop       =   1545
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Brick.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnimate_Ending_3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6960
      Top             =   3000
   End
   Begin VB.Timer tmrAnimate_Ending_2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6960
      Top             =   3480
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4800
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
      Begin VB.CommandButton bttnLevelDown 
         Height          =   495
         Left            =   1320
         Picture         =   "Brick.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton bttnLevelUp 
         Height          =   495
         Left            =   1320
         Picture         =   "Brick.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin MSComctlLib.Slider levelSlider 
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         LargeChange     =   1
         Max             =   9
      End
      Begin VB.Label lblLevelCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LEVEL:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   105
      End
   End
   Begin VB.CheckBox muteCheck 
      Caption         =   "MUTE SOUND"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame speedFrame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
      Begin MSComctlLib.Slider speedSlider 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   9
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblSpeedCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SPEED:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame scoreFrame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.Label lblLines 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   960
         Width           =   105
      End
      Begin VB.Label lblLineCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LINES:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   510
      End
      Begin VB.Label lblScoreCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SCORE:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblScore 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   105
      End
   End
   Begin VB.Timer tmrAnimate_Ending 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6960
      Top             =   3960
   End
   Begin VB.PictureBox brickBoard 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   2040
      ScaleHeight     =   4800
      ScaleWidth      =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2700
      Begin VB.Image Blank_Image 
         Height          =   390
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image Block_Image 
         Height          =   390
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image Block_Image 
         Height          =   390
         Index           =   2
         Left            =   840
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image Block_Image 
         Height          =   390
         Index           =   3
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image Block_Image 
         Height          =   390
         Index           =   4
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image Grid_Image 
         Height          =   390
         Index           =   0
         Left            =   840
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   390
      End
   End
   Begin VB.Frame nextBrickFrame 
      Caption         =   "NEXT BRICK"
      Height          =   1455
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.Image Next_Picture 
         Height          =   240
         Index           =   4
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Next_Picture 
         Height          =   240
         Index           =   3
         Left            =   960
         Stretch         =   -1  'True
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Next_Picture 
         Height          =   240
         Index           =   2
         Left            =   600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Next_Picture 
         Height          =   240
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Timer tmrFall 
      Enabled         =   0   'False
      Left            =   6960
      Top             =   4440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6840
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   6840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image BlockGraphic 
      Height          =   240
      Index           =   3
      Left            =   7200
      Picture         =   "Brick.frx":0B8E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   240
   End
   Begin VB.Image BlockGraphic 
      Height          =   240
      Index           =   5
      Left            =   7200
      Picture         =   "Brick.frx":1130
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image BlockGraphic 
      Height          =   240
      Index           =   4
      Left            =   7200
      Picture         =   "Brick.frx":170A
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image BlockGraphic 
      Height          =   240
      Index           =   1
      Left            =   7200
      Picture         =   "Brick.frx":1CDD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image BlockGraphic 
      Height          =   240
      Index           =   2
      Left            =   7200
      Picture         =   "Brick.frx":2476
      Stretch         =   -1  'True
      Top             =   360
      Width           =   240
   End
   Begin VB.Image BlockGraphic 
      Height          =   240
      Index           =   6
      Left            =   7200
      Picture         =   "Brick.frx":28F2
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image BlockGraphic 
      Height          =   240
      Index           =   7
      Left            =   7200
      Picture         =   "Brick.frx":2E96
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   240
   End
   Begin VB.Menu filemnu 
      Caption         =   "&File"
      Begin VB.Menu newmnu 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu pausemnu 
         Caption         =   "&Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu exitmnu 
         Caption         =   "E&xit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu optionsmnu 
      Caption         =   "&Options"
      Begin VB.Menu speedUmnu 
         Caption         =   "Speed &Up"
         Shortcut        =   {F5}
      End
      Begin VB.Menu speedDmnu 
         Caption         =   "Speed &Down"
         Shortcut        =   {F6}
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu sndmnu 
         Caption         =   "&Sound"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu helpmnu 
      Caption         =   "&Help"
      Begin VB.Menu instmnu 
         Caption         =   "&Instructions"
         Shortcut        =   {F1}
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu aboutmnu 
         Caption         =   "&About"
         Shortcut        =   {F8}
      End
   End
End
Attribute VB_Name = "Brick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Grid_Filled(1 To Number_of_Columns, 1 To Number_of_Rows) As Boolean

Private Type Type_Block
    x_pos As Integer
    y_pos As Integer
End Type

Private Type Type_Group_of_Blocks
    Block_Type As Integer
    Rotation As Integer
    Block(1 To 4) As Type_Block
End Type

Private Tetris_Blocks As Type_Group_of_Blocks
Private Number_of_Blocks As Integer
Private Next_Block As Integer

Option Explicit

Private Sub aboutmnu_Click()
tmrFall.Enabled = False
AboutMe.Show vbModal
End Sub

Private Sub brickBoard_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If tmrFall.Enabled = True Then
    If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then
        If Rotate_Blocks(Tetris_Blocks.Block_Type) = True Then
            If muteCheck.Value = 0 Then
                Call sndPlaySound(App.Path & "\snd2.wav", &H1)
            End If
            Call Move_Blocks
        End If
    ElseIf KeyCode = vbKeyLeft Then
        Call Shift_Blocks_Left(Tetris_Blocks.Block_Type)
        Call Move_Blocks
    ElseIf KeyCode = vbKeyRight Then
        Call Shift_Blocks_Right(Tetris_Blocks.Block_Type)
        Call Move_Blocks
    ElseIf KeyCode = vbKeySpace Then
            Call Shoot_Blocks_Down
    ElseIf KeyCode = vbKeyDown Then
        tmrFall.Enabled = False
        tmrFall.Interval = 1
        tmrFall.Enabled = True
        tmrFall.Interval = Speed_Value(Speed)
    End If
End If
End Sub

Public Sub Shoot_Blocks_Down()
On Error Resume Next
Dim i As Integer
Do While Check_For_Collision = False
        For i = 1 To 4
            Tetris_Blocks.Block(i).y_pos = Tetris_Blocks.Block(i).y_pos + 1
        Next i
        Call Move_Blocks
        'Call Check_For_Row
Loop

If muteCheck.Value = 0 Then
    Call sndPlaySound(App.Path & "\snd1.wav", &H1)
End If

If Tetris_Blocks.Block(1).y_pos > 0 And Tetris_Blocks.Block(2).y_pos > 0 And Tetris_Blocks.Block(3).y_pos > 0 And Tetris_Blocks.Block(4).y_pos > 0 Then
    Grid_Filled(Tetris_Blocks.Block(1).x_pos, Tetris_Blocks.Block(1).y_pos) = True
    Grid_Filled(Tetris_Blocks.Block(2).x_pos, Tetris_Blocks.Block(2).y_pos) = True
    Grid_Filled(Tetris_Blocks.Block(3).x_pos, Tetris_Blocks.Block(3).y_pos) = True
    Grid_Filled(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos) = True
    Grid_Image(findGridPos(Tetris_Blocks.Block(1).x_pos, Tetris_Blocks.Block(1).y_pos)).Picture = BlockGraphic(Tetris_Blocks.Block_Type)
    Grid_Image(findGridPos(Tetris_Blocks.Block(2).x_pos, Tetris_Blocks.Block(2).y_pos)).Picture = BlockGraphic(Tetris_Blocks.Block_Type)
    Grid_Image(findGridPos(Tetris_Blocks.Block(3).x_pos, Tetris_Blocks.Block(3).y_pos)).Picture = BlockGraphic(Tetris_Blocks.Block_Type)
    Grid_Image(findGridPos(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos)).Picture = BlockGraphic(Tetris_Blocks.Block_Type)
    Call createBrick
    Exit Sub
End If
End Sub

Private Sub bttnLevelDown_Click()
If Level = 1 Then
    Level = 0
    Call newmnu_Click
ElseIf Level > 1 Then
    Level = Level - 1
    Call newmnu_Click
End If
brickBoard.SetFocus
End Sub

Private Sub bttnLevelUp_Click()
If Level = 0 Then
    Score = 0
    Level = 1
    Lines = 0
    Call newmnu_Click
ElseIf Level < Number_of_Levels Then
    Level = Level + 1
    Call newmnu_Click
End If
brickBoard.SetFocus
End Sub

Private Sub exitmnu_Click()
Unload Brick
Unload AboutMe
Unload Help
End
End Sub

Private Sub Form_Load()
Call Load_Grid_Images
pausemnu.Enabled = False
Score = 0
Points_Value = 100

Speed_Value(1) = 700
Speed_Value(2) = 500
Speed_Value(3) = 300
Speed_Value(4) = 200
Speed_Value(5) = 150
Speed_Value(6) = 100
Speed_Value(7) = 75
Speed_Value(8) = 40
Speed_Value(9) = 10
Level = 0
Speed = 1
sndmnu.Checked = True
End Sub

Public Sub Load_Grid_Images()
On Error Resume Next
Dim i As Integer

For i = 1 To Number_of_Columns * Number_of_Rows
    Load Grid_Image(i)
    Grid_Image(i).Stretch = True
    Grid_Image(i).Visible = True
    Grid_Image(i).Width = Block_Length
    Grid_Image(i).Height = Block_Length
    Grid_Image(i).left = ((i - 1) Mod Number_of_Columns) * Block_Length
    Grid_Image(i).top = (Int((i - 1) / Number_of_Columns)) * Block_Length
Next i
End Sub

Private Sub instmnu_Click()
If tmrFall.Enabled = True Then
    Call pausemnu_Click
End If
'tmrFall.Enabled = False
Help.Show vbModal
End Sub

'Private Sub levelSlider_Change()
'If levelSlider.Value > Level Then
'    If Level = 0 Then
'        Score = 0
'        Lines = 0
'        Level = 1
'        Call newmnu_Click
'    ElseIf Level < Number_of_Levels Then
'        Lines = 0
'        Level = Level + 1
'        Call newmnu_Click
'    End If
'ElseIf levelSlider.Value < Level Then
'    If Level = 1 Then
'        Score = 0
'        Lines = 0
'        Level = 0
'        Call newmnu_Click
'    ElseIf Level > 1 Then
'        Lines = 0
'        Level = Level - 1
'        Call newmnu_Click
'    End If
'End If
'lblLevel.Caption = levelSlider.Value
'brickBoard.SetFocus
'End Sub

Private Sub muteCheck_Click()
brickBoard.SetFocus
End Sub

Private Sub New_Game()
Dim Random_Number As Integer
Dim i, j, k As Integer
Randomize
Random_Number = Int(Rnd * 7) + 1
If Random_Number > 4 And Random_Number <> 7 Then
    Random_Number = Int(Rnd * 7) + 1
End If
Game_End = False
Next_Block = Random_Number
pausemnu.Enabled = True
Call createBrick
tmrFall.Enabled = False
tmrFall.Interval = 1000
tmrFall.Enabled = True
tmrFall.Interval = Speed_Value(Speed)

Score = 0
Lines = 0
Points_Value = 100
lblScore.Caption = 0
lblLines.Caption = 0
Level = 0
lblLevel.Caption = 0
tmrAnimate_Ending.Enabled = False
For i = 1 To Number_of_Columns * Number_of_Rows
    Grid_Image(i).Stretch = True
    Grid_Image(i).Visible = True
    Grid_Image(i).Width = Block_Length
    Grid_Image(i).Height = Block_Length
    Grid_Image(i).Picture = Blank_Image.Picture
    Grid_Image(i).left = ((i - 1) Mod Number_of_Columns) * Block_Length
    Grid_Image(i).top = (Int((i - 1) / Number_of_Columns)) * Block_Length
Next i

For j = 1 To Number_of_Columns
    For k = 1 To Number_of_Rows
        Grid_Filled(j, k) = False
    Next k
Next j
Call createNextBrick
End Sub

Private Sub newmnu_Click()
If Level = 0 Then
    If Score > 0 Then
        If tmrFall.Enabled = True Then
            Call pausemnu_Click
        End If
        'frmHighScores.Show vbModal
    End If
    Call New_Game
Else
    Call New_Levels_Game
End If
End Sub

Private Sub pausemnu_Click()
If tmrFall.Enabled = True Then
    tmrFall.Enabled = False
    pausemnu.Caption = "&Continue"
    Exit Sub
Else
    tmrFall.Enabled = True
    pausemnu.Caption = "&Pause"
    Exit Sub
End If
End Sub

Private Sub sndmnu_Click()
If sndmnu.Checked = True Then
    sndmnu.Checked = False
    muteCheck.Value = 1
    Exit Sub
ElseIf sndmnu.Checked = False Then
    sndmnu.Checked = True
    muteCheck.Value = 0
    Exit Sub
End If
End Sub

Private Sub speedDmnu_Click()
If Speed > 1 Then
    Speed = Speed - 1
    speedSlider.Value = speedSlider.Value - 1
    lblSpeed.Caption = Speed
End If
End Sub

Private Sub speedSlider_Change()
Speed = speedSlider.Value
brickBoard.SetFocus
lblSpeed.Caption = speedSlider.Value
End Sub

Private Sub speedUmnu_Click()
If Speed <> 9 Then
    Speed = Speed + 1
    speedSlider.Value = speedSlider.Value + 1
    lblSpeed.Caption = Speed
End If
End Sub

Private Sub tmrAnimate_Ending_2_Timer()
Static Row As Integer
Static Direction As Integer
Dim i As Integer
tmrFall.Enabled = False
If Row = 0 Then
    Row = Number_of_Rows
End If
Row = Row - 1

If Row < 1 Then
    tmrAnimate_Ending_2.Enabled = False
    Row = 0
    tmrAnimate_Ending_3.Enabled = True
    Exit Sub
End If
For i = Number_of_Columns * Row To Number_of_Columns * (Row - 1) + 1 Step -1
    Grid_Image(i).Picture = Grid_Image(Number_of_Columns * (Number_of_Rows - 1) + ((Grid_Image(i).left + Block_Length) / Block_Length)).Picture
Next i
Block_Image(1).Visible = False
Block_Image(2).Visible = False
Block_Image(3).Visible = False
Block_Image(4).Visible = False
End Sub

Private Sub tmrAnimate_Ending_3_Timer()
Static Row As Integer
Dim i As Integer
tmrFall.Enabled = False
'If Row = 0 Then Row = Number_of_Rows + 1
Row = Row + 1
If Row > Number_of_Rows Then
    tmrAnimate_Ending_3.Enabled = False
    Row = 0
    Call newmnu_Click
    Exit Sub
End If
For i = Number_of_Columns * (Row - 1) To Number_of_Columns * Row
    Grid_Image(i).Picture = Blank_Image.Picture
Next i
Block_Image(1).Visible = False
Block_Image(2).Visible = False
Block_Image(3).Visible = False
Block_Image(4).Visible = False
End Sub

Private Sub tmrFall_Timer()
On Error Resume Next
Dim i As Integer
lblSpeed.Caption = speedSlider.Value
'lblLevel.Caption = levelSlider.Value
If Check_For_Collision = True Then
    If muteCheck.Value = 0 Then
        Call sndPlaySound(App.Path & "\snd1.wav", &H1)
    End If
    If Tetris_Blocks.Block(1).y_pos > 0 And Tetris_Blocks.Block(2).y_pos > 0 And Tetris_Blocks.Block(3).y_pos > 0 And Tetris_Blocks.Block(4).y_pos > 0 Then
        Grid_Filled(Tetris_Blocks.Block(1).x_pos, Tetris_Blocks.Block(1).y_pos) = True
        Grid_Filled(Tetris_Blocks.Block(2).x_pos, Tetris_Blocks.Block(2).y_pos) = True
        Grid_Filled(Tetris_Blocks.Block(3).x_pos, Tetris_Blocks.Block(3).y_pos) = True
        Grid_Filled(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos) = True
        Grid_Image(findGridPos(Tetris_Blocks.Block(1).x_pos, Tetris_Blocks.Block(1).y_pos)).Picture = BlockGraphic(Tetris_Blocks.Block_Type)
        Grid_Image(findGridPos(Tetris_Blocks.Block(2).x_pos, Tetris_Blocks.Block(2).y_pos)).Picture = BlockGraphic(Tetris_Blocks.Block_Type)
        Grid_Image(findGridPos(Tetris_Blocks.Block(3).x_pos, Tetris_Blocks.Block(3).y_pos)).Picture = BlockGraphic(Tetris_Blocks.Block_Type)
        Grid_Image(findGridPos(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos)).Picture = BlockGraphic(Tetris_Blocks.Block_Type)
        Call createBrick
    Else
        tmrFall.Enabled = False
        pausemnu.Enabled = False
        tmrAnimate_Ending.Enabled = True
    End If
End If


pausemnu.Caption = "&Pause"
For i = 1 To 4
    Tetris_Blocks.Block(i).y_pos = Tetris_Blocks.Block(i).y_pos + 1
Next i

If muteCheck.Value = 0 Then
    sndmnu.Checked = True
Else
    sndmnu.Checked = False
End If

Call Move_Blocks
Call Check_For_Row
If tmrFall.Interval <> 1 Then tmrFall.Interval = Speed_Value(Speed)
End Sub

Public Function Check_For_Collision() As Boolean
On Error Resume Next
Dim i As Integer
For i = 1 To 4
    If Tetris_Blocks.Block(i).y_pos >= 0 Then
        If Tetris_Blocks.Block(i).y_pos = Number_of_Rows Then
            Check_For_Collision = True
            Exit Function
        ElseIf Grid_Filled(Tetris_Blocks.Block(i).x_pos, Tetris_Blocks.Block(i).y_pos + 1) = True Then
            Check_For_Collision = True
            Exit Function
        End If
    End If
Next i
Check_For_Collision = False
End Function

Public Sub Check_For_Row()
Dim Check_Row As Boolean
On Error Resume Next
Dim i, j, Filled_Row As Integer
For i = 1 To Number_of_Rows
    For j = 1 To Number_of_Columns
        If Grid_Filled(j, i) = False Then
            Exit For
        Else
            If j = Number_of_Columns Then
                If i = Number_of_Rows And Level > 0 Then
                    Level = Level + 1
                    tmrFall.Enabled = False
                    tmrAnimate_Ending_2.Enabled = True
                    Check_Row = True
                Else
                    Lines = Lines + 1
                    Call Update_Score
                    Filled_Row = i
                    Call Copy_Rows_Down(Filled_Row)
                    Check_Row = True
                End If
            End If
        End If
    Next j
Next i
If Check_Row = True And muteCheck.Value = 0 Then
    Call sndPlaySound(App.Path & "\snd4.wav", &H1)
End If
End Sub

Public Sub Copy_Rows_Down(From_Filled_Row As Integer)
On Error Resume Next
Dim i, j, Column, Row As Integer
If From_Filled_Row > 2 Then
    For i = 1 To Number_of_Columns
        For j = From_Filled_Row To 2 Step -1
            Column = i
            Row = j
            Grid_Filled(Column, Row) = Grid_Filled(Column, Row - 1)
            Grid_Image((Row - 1) * Number_of_Columns + Column).Picture = Grid_Image((Row - 2) * Number_of_Columns + Column).Picture
        Next j
    Next i
ElseIf From_Filled_Row = 2 Then
    For i = 1 To Number_of_Columns
        Column = i
        Row = 2
        Grid_Filled(Column, Row) = Grid_Filled(Column, Row - 1)
        Grid_Image((Row - 1) * Number_of_Columns + Column).Picture = Grid_Image((Row - 2) * Number_of_Columns + Column).Picture
    Next i
End If
End Sub

Public Sub Update_Score()
Score = Score + Points_Value
If Points_Value = 100 Then
    Points_Value = 150
ElseIf Points_Value = 150 Then
    Points_Value = 250
ElseIf Points_Value = 250 Then
    Points_Value = 500
Else
    Points_Value = 1000
End If
lblScore.Caption = Score
lblLines.Caption = Lines
End Sub

Public Sub Move_Blocks()
Dim i As Integer

For i = 1 To 4
    Block_Image(i).left = Tetris_Blocks.Block(i).x_pos * Block_Length - Block_Length
    Block_Image(i).top = Tetris_Blocks.Block(i).y_pos * Block_Length - Block_Length
Next i

End Sub

Public Function findGridPos(X As Integer, Y As Integer) As Integer
findGridPos = (Y - 1) * Number_of_Columns + X
End Function

Private Sub tmrAnimate_Ending_Timer()
Static Row As Integer
Dim i As Integer
'If Row = 0 Then Row = Number_of_Rows + 1
tmrFall.Enabled = False
Game_End = True
Row = Row + 1
If Row > Number_of_Rows Then
    tmrAnimate_Ending.Enabled = False
    Row = 0
    Exit Sub
End If
For i = Number_of_Columns * (Row - 1) To Number_of_Columns * Row
    Grid_Image(i).Picture = Blank_Image.Picture
Next i
Block_Image(1).Visible = False
Block_Image(2).Visible = False
Block_Image(3).Visible = False
Block_Image(4).Visible = False
End Sub

Public Sub createBrick()
On Error Resume Next
Dim Random_Number As Integer
Dim Piece As Integer
Dim i As Integer

Piece = Next_Block

Select Case Piece
    Case Is = 1
        Tetris_Blocks.Block_Type = 1
        Tetris_Blocks.Rotation = 1
        
        Tetris_Blocks.Block(1).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(1).y_pos = -3
        
        Tetris_Blocks.Block(2).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(2).y_pos = -2
        
        Tetris_Blocks.Block(3).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(3).y_pos = -1

        Tetris_Blocks.Block(4).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(4).y_pos = 0
        
        Number_of_Blocks = Number_of_Blocks + 4

        For i = 1 To 4
            Block_Image(i).Width = Block_Length
            Block_Image(i).Height = Block_Length
            Block_Image(i).Stretch = True
            Block_Image(i).Visible = True
            Block_Image(i).Picture = BlockGraphic(1)
            Block_Image(i).left = Tetris_Blocks.Block(i).x_pos * Block_Length - Block_Length
            Block_Image(i).top = Tetris_Blocks.Block(i).y_pos * Block_Length - Block_Length
        Next i
    Case Is = 2
        Tetris_Blocks.Block_Type = 2
        Tetris_Blocks.Rotation = 1
        
        Tetris_Blocks.Block(1).x_pos = Middle_Start_Up_X - 1
        Tetris_Blocks.Block(1).y_pos = -1
        
        Tetris_Blocks.Block(2).x_pos = Middle_Start_Up_X - 1
        Tetris_Blocks.Block(2).y_pos = 0
        
        Tetris_Blocks.Block(3).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(3).y_pos = 0

        Tetris_Blocks.Block(4).x_pos = Middle_Start_Up_X + 1
        Tetris_Blocks.Block(4).y_pos = 0
        
        Number_of_Blocks = Number_of_Blocks + 4

        For i = 1 To 4
            Block_Image(i).Width = Block_Length
            Block_Image(i).Height = Block_Length
            Block_Image(i).Stretch = True
            Block_Image(i).Visible = True
            Block_Image(i).Picture = BlockGraphic(2)
            Block_Image(i).left = Tetris_Blocks.Block(i).x_pos * Block_Length - Block_Length
            Block_Image(i).top = Tetris_Blocks.Block(i).y_pos * Block_Length - Block_Length
        Next i
    Case Is = 3
        Tetris_Blocks.Block_Type = 3
        Tetris_Blocks.Rotation = 1
        
        Tetris_Blocks.Block(1).x_pos = Middle_Start_Up_X - 1
        Tetris_Blocks.Block(1).y_pos = 0
        
        Tetris_Blocks.Block(2).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(2).y_pos = 0
        
        Tetris_Blocks.Block(3).x_pos = Middle_Start_Up_X + 1
        Tetris_Blocks.Block(3).y_pos = 0

        Tetris_Blocks.Block(4).x_pos = Middle_Start_Up_X + 1
        Tetris_Blocks.Block(4).y_pos = -1
        
        Number_of_Blocks = Number_of_Blocks + 4

        For i = 1 To 4
            Block_Image(i).Width = Block_Length
            Block_Image(i).Height = Block_Length
            Block_Image(i).Stretch = True
            Block_Image(i).Visible = True
            Block_Image(i).Picture = BlockGraphic(3)
            Block_Image(i).left = Tetris_Blocks.Block(i).x_pos * Block_Length - Block_Length
            Block_Image(i).top = Tetris_Blocks.Block(i).y_pos * Block_Length - Block_Length
        Next i
    Case Is = 4
        Tetris_Blocks.Block_Type = 4
        Tetris_Blocks.Rotation = 1
        
        Tetris_Blocks.Block(1).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(1).y_pos = 0
        
        Tetris_Blocks.Block(2).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(2).y_pos = -1
        
        Tetris_Blocks.Block(3).x_pos = Middle_Start_Up_X + 1
        Tetris_Blocks.Block(3).y_pos = -2

        Tetris_Blocks.Block(4).x_pos = Middle_Start_Up_X + 1
        Tetris_Blocks.Block(4).y_pos = -1
        
        Number_of_Blocks = Number_of_Blocks + 4

        For i = 1 To 4
            Block_Image(i).Width = Block_Length
            Block_Image(i).Height = Block_Length
            Block_Image(i).Stretch = True
            Block_Image(i).Visible = True
            Block_Image(i).Picture = BlockGraphic(4)
            Block_Image(i).left = Tetris_Blocks.Block(i).x_pos * Block_Length - Block_Length
            Block_Image(i).top = Tetris_Blocks.Block(i).y_pos * Block_Length - Block_Length
        Next i
    Case Is = 5
        Tetris_Blocks.Block_Type = 5
        Tetris_Blocks.Rotation = 1
        
        Tetris_Blocks.Block(1).x_pos = Middle_Start_Up_X - 1
        Tetris_Blocks.Block(1).y_pos = -2
        
        Tetris_Blocks.Block(2).x_pos = Middle_Start_Up_X - 1
        Tetris_Blocks.Block(2).y_pos = -1
        
        Tetris_Blocks.Block(3).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(3).y_pos = 0

        Tetris_Blocks.Block(4).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(4).y_pos = -1
        
        Number_of_Blocks = Number_of_Blocks + 4

        For i = 1 To 4
            Block_Image(i).Width = Block_Length
            Block_Image(i).Height = Block_Length
            Block_Image(i).Stretch = True
            Block_Image(i).Visible = True
            Block_Image(i).Picture = BlockGraphic(5)
            Block_Image(i).left = Tetris_Blocks.Block(i).x_pos * Block_Length - Block_Length
            Block_Image(i).top = Tetris_Blocks.Block(i).y_pos * Block_Length - Block_Length
        Next i
    Case Is = 6
        Tetris_Blocks.Block_Type = 6
        Tetris_Blocks.Rotation = 1
        
        Tetris_Blocks.Block(1).x_pos = Middle_Start_Up_X - 1
        Tetris_Blocks.Block(1).y_pos = 0
        
        Tetris_Blocks.Block(2).x_pos = Middle_Start_Up_X + 1
        Tetris_Blocks.Block(2).y_pos = 0
        
        Tetris_Blocks.Block(3).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(3).y_pos = -1

        Tetris_Blocks.Block(4).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(4).y_pos = 0
        
        Number_of_Blocks = Number_of_Blocks + 4

        For i = 1 To 4
            Block_Image(i).Width = Block_Length
            Block_Image(i).Height = Block_Length
            Block_Image(i).Stretch = True
            Block_Image(i).Visible = True
            Block_Image(i).Picture = BlockGraphic(6)
            Block_Image(i).left = Tetris_Blocks.Block(i).x_pos * Block_Length - Block_Length
            Block_Image(i).top = Tetris_Blocks.Block(i).y_pos * Block_Length - Block_Length
        Next i
    Case Is = 7
        Tetris_Blocks.Block_Type = 7
        Tetris_Blocks.Rotation = 0
        
        Tetris_Blocks.Block(1).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(1).y_pos = -1
        
        Tetris_Blocks.Block(2).x_pos = Middle_Start_Up_X + 1
        Tetris_Blocks.Block(2).y_pos = -1
        
        Tetris_Blocks.Block(3).x_pos = Middle_Start_Up_X
        Tetris_Blocks.Block(3).y_pos = 0

        Tetris_Blocks.Block(4).x_pos = Middle_Start_Up_X + 1
        Tetris_Blocks.Block(4).y_pos = 0
        
        Number_of_Blocks = Number_of_Blocks + 4

        For i = 1 To 4
            Block_Image(i).Width = Block_Length
            Block_Image(i).Height = Block_Length
            Block_Image(i).Stretch = True
            Block_Image(i).Visible = True
            Block_Image(i).Picture = BlockGraphic(7)
            Block_Image(i).left = Tetris_Blocks.Block(i).x_pos * Block_Length - Block_Length
            Block_Image(i).top = Tetris_Blocks.Block(i).y_pos * Block_Length - Block_Length
        Next i
End Select

Randomize
Random_Number = Int(Rnd * 7) + 1
If Random_Number = Piece Then
    Random_Number = Int(Rnd * 7) + 1
End If
Next_Block = Random_Number
Call createNextBrick
End Sub

Private Function Rotate_Blocks(Type_of_Block As Integer) As Boolean
On Error Resume Next
Select Case Type_of_Block
    Case Is = 1
        If Tetris_Blocks.Rotation = 1 Then
            If Tetris_Blocks.Block(1).x_pos > 2 And Tetris_Blocks.Block(1).x_pos < Number_of_Columns Then
                Tetris_Blocks.Rotation = 2
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 2
                Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos + 1
                Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
                Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos - 1
                Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
                Tetris_Blocks.Block(4).y_pos = Tetris_Blocks.Block(4).y_pos - 2
                Rotate_Blocks = True
                Exit Function
            End If
        ElseIf Tetris_Blocks.Rotation = 2 Then
            If Tetris_Blocks.Block(3).x_pos < Number_of_Rows - 1 Then
                If Grid_Filled(Tetris_Blocks.Block(3).x_pos, Tetris_Blocks.Block(3).y_pos + 1) = False And Grid_Filled(Tetris_Blocks.Block(3).x_pos, Tetris_Blocks.Block(3).y_pos + 2) = False And Grid_Filled(Tetris_Blocks.Block(3).x_pos, Tetris_Blocks.Block(3).y_pos - 1) = False Then
                    Tetris_Blocks.Rotation = 1
                    Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 2
                    Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos - 1
                    Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
                    Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos + 1
                    Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
                    Tetris_Blocks.Block(4).y_pos = Tetris_Blocks.Block(4).y_pos + 2
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        Else
            End
        End If
    Case Is = 2
        If Tetris_Blocks.Rotation = 1 Then
            If Grid_Filled(Tetris_Blocks.Block(2).x_pos, Tetris_Blocks.Block(2).y_pos + 1) = False Then
                Tetris_Blocks.Rotation = 2
                Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
                Tetris_Blocks.Block(2).y_pos = Tetris_Blocks.Block(2).y_pos - 1
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
                Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 2
                Tetris_Blocks.Block(4).y_pos = Tetris_Blocks.Block(4).y_pos + 1
                Rotate_Blocks = True
                Exit Function
            End If
        ElseIf Tetris_Blocks.Rotation = 2 Then
            If Can_Shift_Blocks_Right(Type_of_Block) = True Then
                If Grid_Filled(Tetris_Blocks.Block(2).x_pos + 1, Tetris_Blocks.Block(2).y_pos) = False And Grid_Filled(Tetris_Blocks.Block(2).x_pos + 1, Tetris_Blocks.Block(2).y_pos + 1) = False Then
                    Tetris_Blocks.Rotation = 3
                    Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 2
                    Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos - 1
                    Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 2
                    Tetris_Blocks.Block(4).y_pos = Tetris_Blocks.Block(4).y_pos - 1
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        ElseIf Tetris_Blocks.Rotation = 3 Then
            If Grid_Filled(Tetris_Blocks.Block(1).x_pos, Tetris_Blocks.Block(1).y_pos + 2) = False And Grid_Filled(Tetris_Blocks.Block(2).x_pos, Tetris_Blocks.Block(2).y_pos + 2) = False Then
                Tetris_Blocks.Rotation = 4
                Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos + 2
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
                Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos + 1
                Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
                Tetris_Blocks.Block(4).y_pos = Tetris_Blocks.Block(4).y_pos + 1
                Rotate_Blocks = True
                Exit Function
            End If
        ElseIf Tetris_Blocks.Rotation = 4 Then
            If Can_Shift_Blocks_Right(Type_of_Block) = True Then
                If Grid_Filled(Tetris_Blocks.Block(3).x_pos + 1, Tetris_Blocks.Block(3).y_pos) = False And Grid_Filled(Tetris_Blocks.Block(3).x_pos - 1, Tetris_Blocks.Block(3).y_pos) = False Then
                    Tetris_Blocks.Rotation = 1
                    Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos - 2
                    Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
                    Tetris_Blocks.Block(2).y_pos = Tetris_Blocks.Block(2).y_pos + 1
                    Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
                    Tetris_Blocks.Block(4).y_pos = Tetris_Blocks.Block(4).y_pos - 1
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        End If
    Case Is = 3
        If Tetris_Blocks.Rotation = 1 Then
            If Grid_Filled(Tetris_Blocks.Block(3).x_pos, Tetris_Blocks.Block(3).y_pos + 1) = False And Grid_Filled(Tetris_Blocks.Block(2).x_pos, Tetris_Blocks.Block(2).y_pos + 1) = False Then
                Tetris_Blocks.Rotation = 2
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
                Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos + 1
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
                Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos - 1
                Tetris_Blocks.Block(4).y_pos = Tetris_Blocks.Block(4).y_pos + 2
                Rotate_Blocks = True
                Exit Function
            End If
        ElseIf Tetris_Blocks.Rotation = 2 Then
            If Can_Shift_Blocks_Left(Type_of_Block) = True Then
                If Grid_Filled(Tetris_Blocks.Block(3).x_pos - 1, Tetris_Blocks.Block(3).y_pos) = False And Grid_Filled(Tetris_Blocks.Block(2).x_pos - 1, Tetris_Blocks.Block(2).y_pos) = False Then
                    Tetris_Blocks.Rotation = 3
                    Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
                    Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos - 1
                    Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
                    Tetris_Blocks.Block(2).y_pos = Tetris_Blocks.Block(2).y_pos - 1
                    Tetris_Blocks.Block(4).y_pos = Tetris_Blocks.Block(4).y_pos - 2
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        ElseIf Tetris_Blocks.Rotation = 3 Then
            If Grid_Filled(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos - 1) = False And Grid_Filled(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos + 2) = False Then
                Tetris_Blocks.Rotation = 4
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 2
                Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos + 1
                Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 2
                Tetris_Blocks.Block(2).y_pos = Tetris_Blocks.Block(2).y_pos + 1
                Rotate_Blocks = True
                Exit Function
            End If
        ElseIf Tetris_Blocks.Rotation = 4 Then
            If Can_Shift_Blocks_Left(Type_of_Block) = True Then
                If Grid_Filled(Tetris_Blocks.Block(2).x_pos - 1, Tetris_Blocks.Block(2).y_pos) = False And Grid_Filled(Tetris_Blocks.Block(2).x_pos - 2, Tetris_Blocks.Block(2).y_pos) = False Then
                    Tetris_Blocks.Rotation = 1
                    Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 2
                    Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos - 1
                    Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
                    Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
                    Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos + 1
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        End If
    Case Is = 4
        If Tetris_Blocks.Rotation = 1 Then
            If Can_Shift_Blocks_Left(Type_of_Block) = True Then
                If Grid_Filled(Tetris_Blocks.Block(3).x_pos - 2, Tetris_Blocks.Block(3).y_pos) = False Then
                    Tetris_Blocks.Rotation = 2
                    Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
                    Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos - 2
                    Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        ElseIf Tetris_Blocks.Rotation = 2 Then
            If Grid_Filled(Tetris_Blocks.Block(2).x_pos, Tetris_Blocks.Block(2).y_pos + 1) = False Then
                Tetris_Blocks.Rotation = 1
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
                Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos + 2
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
                Rotate_Blocks = True
                Exit Function
            End If
        End If
    Case Is = 5
        If Tetris_Blocks.Rotation = 1 Then
            If Can_Shift_Blocks_Right(Type_of_Block) = True Then
                If Grid_Filled(Tetris_Blocks.Block(1).x_pos + 2, Tetris_Blocks.Block(1).y_pos) = False Then
                    Tetris_Blocks.Rotation = 2
                    Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
                    Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
                    Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos - 2
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        ElseIf Tetris_Blocks.Rotation = 2 Then
            If Grid_Filled(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos + 1) = False Then
                Tetris_Blocks.Rotation = 1
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
                Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos + 2
                Rotate_Blocks = True
                Exit Function
            End If
        End If
    Case Is = 6
        If Tetris_Blocks.Rotation = 1 Then
            If Grid_Filled(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos + 1) = False Then
                Tetris_Blocks.Rotation = 2
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
                Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos + 1
                Rotate_Blocks = True
                Exit Function
            End If
        ElseIf Tetris_Blocks.Rotation = 2 Then
            If Can_Shift_Blocks_Left(Type_of_Block) = True Then
                If Grid_Filled(Tetris_Blocks.Block(4).x_pos - 1, Tetris_Blocks.Block(4).y_pos) = False Then
                    Tetris_Blocks.Rotation = 3
                    Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
                    Tetris_Blocks.Block(1).y_pos = Tetris_Blocks.Block(1).y_pos - 1
                    Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos + 2
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        ElseIf Tetris_Blocks.Rotation = 3 Then
            If Grid_Filled(Tetris_Blocks.Block(4).x_pos, Tetris_Blocks.Block(4).y_pos - 1) = False Then
                Tetris_Blocks.Rotation = 4
                Tetris_Blocks.Block(3).y_pos = Tetris_Blocks.Block(3).y_pos - 2
                Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
                Tetris_Blocks.Block(2).y_pos = Tetris_Blocks.Block(2).y_pos + 1
                Rotate_Blocks = True
                Exit Function
            End If
        ElseIf Tetris_Blocks.Rotation = 4 Then
            If Can_Shift_Blocks_Right(Type_of_Block) = True Then
                If Grid_Filled(Tetris_Blocks.Block(4).x_pos + 1, Tetris_Blocks.Block(4).y_pos) = False Then
                    Tetris_Blocks.Rotation = 1
                    Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
                    Tetris_Blocks.Block(2).y_pos = Tetris_Blocks.Block(2).y_pos - 1
                    Rotate_Blocks = True
                    Exit Function
                End If
            End If
        End If
    Case Is = 7
        'Do not Rotate
End Select
Rotate_Blocks = False
End Function

Public Sub Shift_Blocks_Left(Block_Type As Integer)
On Error Resume Next
Dim i As Integer
For i = 1 To 4
    If Tetris_Blocks.Block(i).y_pos > 0 Then
        If Tetris_Blocks.Block(i).x_pos > 1 Then
            If Grid_Filled(Tetris_Blocks.Block(i).x_pos - 1, Tetris_Blocks.Block(i).y_pos) = True Then
                Exit Sub
            End If
        End If
    End If
Next i
Select Case Block_Type
    Case Is = 1
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
        End If
    Case Is = 2
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
        End If
    Case Is = 3
        If Tetris_Blocks.Rotation = 4 Then
            If Tetris_Blocks.Block(3).x_pos > 1 Then
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
                Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
                Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
            End If
        Else
            If Tetris_Blocks.Block(1).x_pos > 1 Then
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
                Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
                Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
            End If
        End If
    Case Is = 4
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
        End If
    Case Is = 5
        If Tetris_Blocks.Block(2).x_pos > 1 Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
        End If
    Case Is = 6
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
        End If
    Case Is = 7
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos - 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos - 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos - 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos - 1
        End If
End Select
End Sub

Public Sub Shift_Blocks_Right(Block_Type As Integer)
On Error Resume Next
Dim i As Integer
For i = 1 To 4
    If Tetris_Blocks.Block(i).y_pos > 0 Then
        If Tetris_Blocks.Block(i).x_pos < Number_of_Columns Then
            If Grid_Filled(Tetris_Blocks.Block(i).x_pos + 1, Tetris_Blocks.Block(i).y_pos) = True Then
                Exit Sub
            End If
        End If
    End If
Next i
Select Case Block_Type
    Case Is = 1
        If Tetris_Blocks.Block(4).x_pos < Number_of_Columns Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
        End If
    Case Is = 2
        If Tetris_Blocks.Rotation = 2 Then
            If Tetris_Blocks.Block(2).x_pos < Number_of_Columns Then
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
                Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
                Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
            End If
        Else
            If Tetris_Blocks.Block(4).x_pos < Number_of_Columns Then
                Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
                Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
                Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
                Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
            End If
        End If
    Case Is = 3
        If Tetris_Blocks.Block(4).x_pos < Number_of_Columns Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
        End If
    Case Is = 4
        If Tetris_Blocks.Block(4).x_pos < Number_of_Columns Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
        End If
    Case Is = 5
        If Tetris_Blocks.Block(3).x_pos < Number_of_Columns Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
        End If
    Case Is = 6
        If Tetris_Blocks.Block(2).x_pos < Number_of_Columns Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
        End If
    Case Is = 7
        If Tetris_Blocks.Block(2).x_pos < Number_of_Columns Then
            Tetris_Blocks.Block(1).x_pos = Tetris_Blocks.Block(1).x_pos + 1
            Tetris_Blocks.Block(2).x_pos = Tetris_Blocks.Block(2).x_pos + 1
            Tetris_Blocks.Block(3).x_pos = Tetris_Blocks.Block(3).x_pos + 1
            Tetris_Blocks.Block(4).x_pos = Tetris_Blocks.Block(4).x_pos + 1
        End If
End Select
End Sub

Public Function Can_Shift_Blocks_Left(Block_Type As Integer) As Boolean
Select Case Block_Type
    Case Is = 1
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Can_Shift_Blocks_Left = True
            Exit Function
        End If
    Case Is = 2
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Can_Shift_Blocks_Left = True
            Exit Function
        End If
    Case Is = 3
        If Tetris_Blocks.Rotation = 4 Then
            If Tetris_Blocks.Block(3).x_pos > 1 Then
                Can_Shift_Blocks_Left = True
                Exit Function
            End If
        Else
            If Tetris_Blocks.Block(1).x_pos > 1 Then
                Can_Shift_Blocks_Left = True
                Exit Function
            End If
        End If
    Case Is = 4
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Can_Shift_Blocks_Left = True
            Exit Function
        End If
    Case Is = 5
        If Tetris_Blocks.Block(2).x_pos > 1 Then
            Can_Shift_Blocks_Left = True
            Exit Function
        End If
    Case Is = 6
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Can_Shift_Blocks_Left = True
            Exit Function
        End If
    Case Is = 7
        If Tetris_Blocks.Block(1).x_pos > 1 Then
            Can_Shift_Blocks_Left = True
            Exit Function
        End If
End Select
Can_Shift_Blocks_Left = False
End Function

Public Function Can_Shift_Blocks_Right(Block_Type As Integer) As Boolean
Select Case Block_Type
    Case Is = 1
        If Tetris_Blocks.Block(4).x_pos < Number_of_Columns Then
            Can_Shift_Blocks_Right = True
            Exit Function
        End If
    Case Is = 2
        If Tetris_Blocks.Rotation = 2 Then
            If Tetris_Blocks.Block(2).x_pos < Number_of_Columns Then
                Can_Shift_Blocks_Right = True
                Exit Function
            End If
        Else
            If Tetris_Blocks.Block(4).x_pos < Number_of_Columns Then
                Can_Shift_Blocks_Right = True
                Exit Function
            End If
        End If
    Case Is = 3
        If Tetris_Blocks.Block(4).x_pos < Number_of_Columns Then
            Can_Shift_Blocks_Right = True
            Exit Function
        End If
    Case Is = 4
        If Tetris_Blocks.Block(4).x_pos < Number_of_Columns Then
            Can_Shift_Blocks_Right = True
            Exit Function
        End If
    Case Is = 5
        If Tetris_Blocks.Block(3).x_pos < Number_of_Columns Then
            Can_Shift_Blocks_Right = True
            Exit Function
        End If
    Case Is = 6
        If Tetris_Blocks.Block(2).x_pos < Number_of_Columns Then
            Can_Shift_Blocks_Right = True
            Exit Function
        End If
    Case Is = 7
        If Tetris_Blocks.Block(2).x_pos < Number_of_Columns Then
            Can_Shift_Blocks_Right = True
            Exit Function
        End If
End Select
Can_Shift_Blocks_Right = False
End Function

Public Sub createNextBrick()
Dim i, left, top As Integer

Select Case Next_Block
    Case Is = 1
        left = 480
        For i = 1 To 4
            Next_Picture(i).Picture = BlockGraphic(Next_Block)
            Next_Picture(i).left = left
            Next_Picture(i).top = 720
            left = left + 240
        Next
    Case Is = 2
        For i = 1 To 4
            Next_Picture(i).Picture = BlockGraphic(Next_Block)
        Next
        Next_Picture(1).left = 600
        Next_Picture(1).top = 600
        Next_Picture(2).left = 600
        Next_Picture(2).top = 840
        Next_Picture(3).left = 840
        Next_Picture(3).top = 840
        Next_Picture(4).left = 1080
        Next_Picture(4).top = 840
    Case Is = 3
        For i = 1 To 4
            Next_Picture(i).Picture = BlockGraphic(Next_Block)
        Next
        Next_Picture(1).left = 960
        Next_Picture(1).top = 600
        Next_Picture(2).left = 480
        Next_Picture(2).top = 840
        Next_Picture(3).left = 720
        Next_Picture(3).top = 840
        Next_Picture(4).left = 960
        Next_Picture(4).top = 840
    Case Is = 4
        For i = 1 To 4
            Next_Picture(i).Picture = BlockGraphic(Next_Block)
        Next
        Next_Picture(1).left = 840
        Next_Picture(1).top = 480
        Next_Picture(2).left = 600
        Next_Picture(2).top = 720
        Next_Picture(3).left = 840
        Next_Picture(3).top = 720
        Next_Picture(4).left = 600
        Next_Picture(4).top = 960
    Case Is = 5
        For i = 1 To 4
            Next_Picture(i).Picture = BlockGraphic(Next_Block)
        Next
        Next_Picture(1).left = 600
        Next_Picture(1).top = 480
        Next_Picture(2).left = 600
        Next_Picture(2).top = 720
        Next_Picture(3).left = 840
        Next_Picture(3).top = 720
        Next_Picture(4).left = 840
        Next_Picture(4).top = 960
    Case Is = 6
        For i = 1 To 4
            Next_Picture(i).Picture = BlockGraphic(Next_Block)
        Next
        Next_Picture(1).left = 840
        Next_Picture(1).top = 480
        Next_Picture(2).left = 600
        Next_Picture(2).top = 720
        Next_Picture(3).left = 840
        Next_Picture(3).top = 720
        Next_Picture(4).left = 1080
        Next_Picture(4).top = 720
    Case Is = 7
        For i = 1 To 4
            Next_Picture(i).Picture = BlockGraphic(Next_Block)
        Next
        Next_Picture(1).left = 600
        Next_Picture(1).top = 600
        Next_Picture(2).left = 840
        Next_Picture(2).top = 600
        Next_Picture(3).left = 600
        Next_Picture(3).top = 840
        Next_Picture(4).left = 840
        Next_Picture(4).top = 840
End Select
End Sub

Public Sub New_Levels_Game()
Dim Random_Number As Integer
Dim i, j, k As Integer
Randomize
Random_Number = Int(Rnd * 7) + 1
If Random_Number > 4 And Random_Number <> 7 Then
    Random_Number = Int(Rnd * 7) + 1
End If
Next_Block = Random_Number
Call createBrick
tmrFall.Enabled = False
tmrFall.Interval = 1000
tmrFall.Enabled = True
tmrFall.Interval = Speed_Value(Speed)
If Level > Number_of_Levels Then
    Level = Number_of_Levels
    MsgBox "Congradulations ! You have completed all levels !", vbOKOnly, Brick.Caption
    tmrFall.Enabled = False
    pausemnu.Enabled = False
    Exit Sub
End If
Points_Value = 100
lblScore.Caption = Score
lblLevel.Caption = Level
lblLines.Caption = Lines
pausemnu.Caption = "&Pause"
tmrAnimate_Ending.Enabled = False
pausemnu.Enabled = True

For i = 1 To Number_of_Columns * Number_of_Rows
    Grid_Image(i).Stretch = True
    Grid_Image(i).Visible = True
    Grid_Image(i).Width = Block_Length
    Grid_Image(i).Height = Block_Length
    Grid_Image(i).Picture = Blank_Image.Picture
    Grid_Image(i).left = ((i - 1) Mod Number_of_Columns) * Block_Length
    Grid_Image(i).top = (Int((i - 1) / Number_of_Columns)) * Block_Length
Next i

For j = 1 To Number_of_Columns
    For k = 1 To Number_of_Rows
        Grid_Filled(j, k) = False
    Next k
Next j

Select Case Level
    Case Is = 1
        Grid_Image(195).Picture = BlockGraphic(1).Picture
        Grid_Filled(8, 18) = True
        Grid_Image(201).Picture = BlockGraphic(5).Picture
        Grid_Filled(3, 19) = True
        Grid_Image(204).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 19) = True
        Grid_Image(209).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 19) = True
        Grid_Image(210).Picture = BlockGraphic(7).Picture
        Grid_Filled(1, 20) = True
        Grid_Image(212).Picture = BlockGraphic(5).Picture
        Grid_Filled(3, 20) = True
        Grid_Image(214).Picture = BlockGraphic(4).Picture
        Grid_Filled(5, 20) = True
        Grid_Image(215).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 20) = True
        Grid_Image(216).Picture = BlockGraphic(4).Picture
        Grid_Filled(7, 20) = True
        Grid_Image(219).Picture = BlockGraphic(6).Picture
        Grid_Filled(10, 20) = True
    Case Is = 2
        Grid_Image(180).Picture = BlockGraphic(1).Picture
        Grid_Filled(4, 17) = True
        Grid_Image(187).Picture = BlockGraphic(7).Picture
        Grid_Filled(11, 17) = True
        Grid_Image(193).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 18) = True
        Grid_Image(196).Picture = BlockGraphic(7).Picture
        Grid_Filled(9, 18) = True
        Grid_Image(201).Picture = BlockGraphic(3).Picture
        Grid_Filled(3, 19) = True
        Grid_Image(203).Picture = BlockGraphic(4).Picture
        Grid_Filled(5, 19) = True
        Grid_Image(204).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 19) = True
        Grid_Image(205).Picture = BlockGraphic(4).Picture
        Grid_Filled(7, 19) = True
        Grid_Image(209).Picture = BlockGraphic(5).Picture
        Grid_Filled(11, 19) = True
        Grid_Image(211).Picture = BlockGraphic(3).Picture
        Grid_Filled(2, 20) = True
        Grid_Image(212).Picture = BlockGraphic(6).Picture
        Grid_Filled(3, 20) = True
        Grid_Image(214).Picture = BlockGraphic(4).Picture
        Grid_Filled(5, 20) = True
        Grid_Image(215).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 20) = True
        Grid_Image(216).Picture = BlockGraphic(2).Picture
        Grid_Filled(7, 20) = True
        Grid_Image(217).Picture = BlockGraphic(2).Picture
        Grid_Filled(8, 20) = True
        Grid_Image(218).Picture = BlockGraphic(2).Picture
        Grid_Filled(9, 20) = True
        Grid_Image(220).Picture = BlockGraphic(5).Picture
        Grid_Filled(11, 20) = True
    Case Is = 3
        Grid_Image(162).Picture = BlockGraphic(4).Picture
        Grid_Filled(8, 15) = True
        Grid_Image(174).Picture = BlockGraphic(4).Picture
        Grid_Filled(9, 16) = True
        Grid_Image(188).Picture = BlockGraphic(2).Picture
        Grid_Filled(1, 18) = True
        Grid_Image(197).Picture = BlockGraphic(7).Picture
        Grid_Filled(10, 18) = True
        Grid_Image(199).Picture = BlockGraphic(2).Picture
        Grid_Filled(1, 19) = True
        Grid_Image(200).Picture = BlockGraphic(2).Picture
        Grid_Filled(2, 19) = True
        Grid_Image(203).Picture = BlockGraphic(3).Picture
        Grid_Filled(5, 19) = True
        Grid_Image(207).Picture = BlockGraphic(7).Picture
        Grid_Filled(9, 19) = True
        Grid_Image(210).Picture = BlockGraphic(2).Picture
        Grid_Filled(1, 20) = True
        Grid_Image(211).Picture = BlockGraphic(2).Picture
        Grid_Filled(2, 20) = True
        Grid_Image(212).Picture = BlockGraphic(2).Picture
        Grid_Filled(3, 20) = True
        Grid_Image(213).Picture = BlockGraphic(6).Picture
        Grid_Filled(4, 20) = True
        Grid_Image(214).Picture = BlockGraphic(6).Picture
        Grid_Filled(5, 20) = True
        Grid_Image(215).Picture = BlockGraphic(3).Picture
        Grid_Filled(6, 20) = True
    Case Is = 4
        Grid_Image(177).Picture = BlockGraphic(1).Picture
        Grid_Filled(1, 17) = True
        Grid_Image(180).Picture = BlockGraphic(3).Picture
        Grid_Filled(4, 17) = True
        Grid_Image(181).Picture = BlockGraphic(3).Picture
        Grid_Filled(5, 17) = True
        Grid_Image(186).Picture = BlockGraphic(5).Picture
        Grid_Filled(10, 17) = True
        Grid_Image(187).Picture = BlockGraphic(4).Picture
        Grid_Filled(11, 17) = True
        Grid_Image(190).Picture = BlockGraphic(3).Picture
        Grid_Filled(3, 18) = True
        Grid_Image(191).Picture = BlockGraphic(3).Picture
        Grid_Filled(4, 18) = True
        Grid_Image(192).Picture = BlockGraphic(3).Picture
        Grid_Filled(5, 18) = True
        Grid_Image(193).Picture = BlockGraphic(3).Picture
        Grid_Filled(6, 18) = True
        Grid_Image(196).Picture = BlockGraphic(5).Picture
        Grid_Filled(9, 18) = True
        Grid_Image(197).Picture = BlockGraphic(4).Picture
        Grid_Filled(10, 18) = True
        Grid_Image(198).Picture = BlockGraphic(4).Picture
        Grid_Filled(11, 18) = True
        Grid_Image(202).Picture = BlockGraphic(3).Picture
        Grid_Filled(4, 19) = True
        Grid_Image(206).Picture = BlockGraphic(5).Picture
        Grid_Filled(8, 19) = True
        Grid_Image(207).Picture = BlockGraphic(6).Picture
        Grid_Filled(9, 19) = True
        Grid_Image(208).Picture = BlockGraphic(6).Picture
        Grid_Filled(10, 19) = True
        Grid_Image(210).Picture = BlockGraphic(2).Picture
        Grid_Filled(1, 20) = True
        Grid_Image(215).Picture = BlockGraphic(7).Picture
        Grid_Filled(6, 20) = True
        Grid_Image(216).Picture = BlockGraphic(7).Picture
        Grid_Filled(7, 20) = True
        Grid_Image(217).Picture = BlockGraphic(7).Picture
        Grid_Filled(8, 20) = True
        Grid_Image(218).Picture = BlockGraphic(6).Picture
        Grid_Filled(9, 20) = True
        Grid_Image(219).Picture = BlockGraphic(6).Picture
        Grid_Filled(10, 20) = True
        Grid_Image(220).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 20) = True
    Case Is = 5
        Grid_Image(179).Picture = BlockGraphic(1).Picture
        Grid_Filled(3, 17) = True
        Grid_Image(184).Picture = BlockGraphic(4).Picture
        Grid_Filled(8, 17) = True
        Grid_Image(187).Picture = BlockGraphic(2).Picture
        Grid_Filled(11, 17) = True
        Grid_Image(193).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 18) = True
        Grid_Image(198).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 18) = True
        Grid_Image(199).Picture = BlockGraphic(3).Picture
        Grid_Filled(1, 19) = True
        Grid_Image(203).Picture = BlockGraphic(5).Picture
        Grid_Filled(5, 19) = True
        Grid_Image(205).Picture = BlockGraphic(3).Picture
        Grid_Filled(7, 19) = True
        Grid_Image(208).Picture = BlockGraphic(3).Picture
        Grid_Filled(10, 19) = True
        Grid_Image(209).Picture = BlockGraphic(4).Picture
        Grid_Filled(11, 19) = True
        Grid_Image(210).Picture = BlockGraphic(3).Picture
        Grid_Filled(1, 20) = True
        Grid_Image(211).Picture = BlockGraphic(3).Picture
        Grid_Filled(2, 20) = True
        Grid_Image(213).Picture = BlockGraphic(6).Picture
        Grid_Filled(4, 20) = True
        Grid_Image(217).Picture = BlockGraphic(7).Picture
        Grid_Filled(8, 20) = True
        Grid_Image(220).Picture = BlockGraphic(5).Picture
        Grid_Filled(11, 20) = True
    Case Is = 6
        Grid_Image(157).Picture = BlockGraphic(3).Picture
        Grid_Filled(3, 15) = True
        Grid_Image(167).Picture = BlockGraphic(3).Picture
        Grid_Filled(2, 16) = True
        Grid_Image(168).Picture = BlockGraphic(1).Picture
        Grid_Filled(3, 16) = True
        Grid_Image(174).Picture = BlockGraphic(5).Picture
        Grid_Filled(9, 16) = True
        Grid_Image(176).Picture = BlockGraphic(7).Picture
        Grid_Filled(11, 16) = True
        Grid_Image(177).Picture = BlockGraphic(3).Picture
        Grid_Filled(1, 17) = True
        Grid_Image(178).Picture = BlockGraphic(4).Picture
        Grid_Filled(2, 17) = True
        Grid_Image(179).Picture = BlockGraphic(1).Picture
        Grid_Filled(3, 17) = True
        Grid_Image(184).Picture = BlockGraphic(5).Picture
        Grid_Filled(8, 17) = True
        Grid_Image(185).Picture = BlockGraphic(4).Picture
        Grid_Filled(9, 17) = True
        Grid_Image(186).Picture = BlockGraphic(7).Picture
        Grid_Filled(10, 17) = True
        Grid_Image(187).Picture = BlockGraphic(7).Picture
        Grid_Filled(11, 17) = True
        Grid_Image(190).Picture = BlockGraphic(1).Picture
        Grid_Filled(3, 18) = True
        Grid_Image(191).Picture = BlockGraphic(1).Picture
        Grid_Filled(4, 18) = True
        Grid_Image(195).Picture = BlockGraphic(5).Picture
        Grid_Filled(8, 18) = True
        Grid_Image(196).Picture = BlockGraphic(6).Picture
        Grid_Filled(9, 18) = True
        Grid_Image(197).Picture = BlockGraphic(6).Picture
        Grid_Filled(10, 18) = True
        Grid_Image(199).Picture = BlockGraphic(2).Picture
        Grid_Filled(1, 19) = True
        Grid_Image(201).Picture = BlockGraphic(4).Picture
        Grid_Filled(3, 19) = True
        Grid_Image(208).Picture = BlockGraphic(6).Picture
        Grid_Filled(10, 19) = True
        Grid_Image(211).Picture = BlockGraphic(2).Picture
        Grid_Filled(2, 20) = True
        Grid_Image(213).Picture = BlockGraphic(4).Picture
        Grid_Filled(4, 20) = True
        Grid_Image(214).Picture = BlockGraphic(4).Picture
        Grid_Filled(5, 20) = True
        Grid_Image(218).Picture = BlockGraphic(7).Picture
        Grid_Filled(9, 20) = True
    Case Is = 7
        Grid_Image(146).Picture = BlockGraphic(7).Picture
        Grid_Filled(3, 14) = True
        Grid_Image(149).Picture = BlockGraphic(7).Picture
        Grid_Filled(6, 14) = True
        Grid_Image(152).Picture = BlockGraphic(7).Picture
        Grid_Filled(9, 14) = True
        Grid_Image(176).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 16) = True
        Grid_Image(182).Picture = BlockGraphic(1).Picture
        Grid_Filled(6, 17) = True
        Grid_Image(183).Picture = BlockGraphic(1).Picture
        Grid_Filled(7, 17) = True
        Grid_Image(184).Picture = BlockGraphic(1).Picture
        Grid_Filled(8, 17) = True
        Grid_Image(185).Picture = BlockGraphic(2).Picture
        Grid_Filled(9, 17) = True
        Grid_Image(186).Picture = BlockGraphic(6).Picture
        Grid_Filled(10, 17) = True
        Grid_Image(187).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 17) = True
        Grid_Image(194).Picture = BlockGraphic(5).Picture
        Grid_Filled(7, 18) = True
        Grid_Image(196).Picture = BlockGraphic(4).Picture
        Grid_Filled(9, 18) = True
        Grid_Image(198).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 18) = True
        Grid_Image(199).Picture = BlockGraphic(3).Picture
        Grid_Filled(1, 19) = True
        Grid_Image(200).Picture = BlockGraphic(3).Picture
        Grid_Filled(2, 19) = True
        Grid_Image(203).Picture = BlockGraphic(5).Picture
        Grid_Filled(5, 19) = True
        Grid_Image(204).Picture = BlockGraphic(5).Picture
        Grid_Filled(6, 19) = True
        Grid_Image(206).Picture = BlockGraphic(4).Picture
        Grid_Filled(8, 19) = True
        Grid_Image(207).Picture = BlockGraphic(4).Picture
        Grid_Filled(9, 19) = True
        Grid_Image(209).Picture = BlockGraphic(3).Picture
        Grid_Filled(11, 19) = True
        Grid_Image(211).Picture = BlockGraphic(3).Picture
        Grid_Filled(2, 20) = True
        Grid_Image(212).Picture = BlockGraphic(5).Picture
        Grid_Filled(3, 20) = True
        Grid_Image(213).Picture = BlockGraphic(5).Picture
        Grid_Filled(4, 20) = True
        Grid_Image(214).Picture = BlockGraphic(5).Picture
        Grid_Filled(5, 20) = True
        Grid_Image(217).Picture = BlockGraphic(4).Picture
        Grid_Filled(8, 20) = True
        Grid_Image(218).Picture = BlockGraphic(3).Picture
        Grid_Filled(9, 20) = True
        Grid_Image(219).Picture = BlockGraphic(3).Picture
        Grid_Filled(10, 20) = True
        Grid_Image(220).Picture = BlockGraphic(3).Picture
        Grid_Filled(11, 20) = True
    Case Is = 8
        Grid_Image(133).Picture = BlockGraphic(1).Picture
        Grid_Filled(1, 13) = True
        Grid_Image(134).Picture = BlockGraphic(1).Picture
        Grid_Filled(2, 13) = True
        Grid_Image(135).Picture = BlockGraphic(1).Picture
        Grid_Filled(3, 13) = True
        Grid_Image(138).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 13) = True
        Grid_Image(141).Picture = BlockGraphic(1).Picture
        Grid_Filled(9, 13) = True
        Grid_Image(142).Picture = BlockGraphic(1).Picture
        Grid_Filled(10, 13) = True
        Grid_Image(143).Picture = BlockGraphic(1).Picture
        Grid_Filled(11, 13) = True
        Grid_Image(149).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 14) = True
        Grid_Image(150).Picture = BlockGraphic(4).Picture
        Grid_Filled(7, 14) = True
        Grid_Image(159).Picture = BlockGraphic(6).Picture
        Grid_Filled(5, 15) = True
        Grid_Image(161).Picture = BlockGraphic(4).Picture
        Grid_Filled(7, 15) = True
        Grid_Image(162).Picture = BlockGraphic(5).Picture
        Grid_Filled(8, 15) = True
        Grid_Image(169).Picture = BlockGraphic(6).Picture
        Grid_Filled(4, 16) = True
        Grid_Image(173).Picture = BlockGraphic(5).Picture
        Grid_Filled(8, 16) = True
        Grid_Image(179).Picture = BlockGraphic(6).Picture
        Grid_Filled(3, 17) = True
        Grid_Image(181).Picture = BlockGraphic(2).Picture
        Grid_Filled(5, 17) = True
        Grid_Image(184).Picture = BlockGraphic(7).Picture
        Grid_Filled(8, 17) = True
        Grid_Image(185).Picture = BlockGraphic(7).Picture
        Grid_Filled(9, 17) = True
        Grid_Image(189).Picture = BlockGraphic(6).Picture
        Grid_Filled(2, 18) = True
        Grid_Image(191).Picture = BlockGraphic(4).Picture
        Grid_Filled(4, 18) = True
        Grid_Image(192).Picture = BlockGraphic(2).Picture
        Grid_Filled(5, 18) = True
        Grid_Image(193).Picture = BlockGraphic(2).Picture
        Grid_Filled(6, 18) = True
        Grid_Image(199).Picture = BlockGraphic(3).Picture
        Grid_Filled(1, 19) = True
        Grid_Image(203).Picture = BlockGraphic(2).Picture
        Grid_Filled(5, 19) = True
        Grid_Image(205).Picture = BlockGraphic(3).Picture
        Grid_Filled(7, 19) = True
        Grid_Image(209).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 19) = True
        Grid_Image(210).Picture = BlockGraphic(3).Picture
        Grid_Filled(1, 20) = True
        Grid_Image(211).Picture = BlockGraphic(3).Picture
        Grid_Filled(2, 20) = True
        Grid_Image(215).Picture = BlockGraphic(1).Picture
        Grid_Filled(6, 20) = True
        Grid_Image(217).Picture = BlockGraphic(3).Picture
        Grid_Filled(8, 20) = True
        Grid_Image(218).Picture = BlockGraphic(6).Picture
        Grid_Filled(9, 20) = True
        Grid_Image(219).Picture = BlockGraphic(6).Picture
        Grid_Filled(10, 20) = True
        Grid_Image(220).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 20) = True
    Case Is = 9
        Grid_Image(127).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 12) = True
        Grid_Image(137).Picture = BlockGraphic(4).Picture
        Grid_Filled(5, 13) = True
        Grid_Image(138).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 13) = True
        Grid_Image(139).Picture = BlockGraphic(4).Picture
        Grid_Filled(7, 13) = True
        Grid_Image(149).Picture = BlockGraphic(4).Picture
        Grid_Filled(6, 14) = True
        Grid_Image(158).Picture = BlockGraphic(1).Picture
        Grid_Filled(4, 15) = True
        Grid_Image(162).Picture = BlockGraphic(1).Picture
        Grid_Filled(8, 15) = True
        Grid_Image(164).Picture = BlockGraphic(5).Picture
        Grid_Filled(10, 15) = True
        Grid_Image(167).Picture = BlockGraphic(2).Picture
        Grid_Filled(2, 16) = True
        Grid_Image(168).Picture = BlockGraphic(1).Picture
        Grid_Filled(3, 16) = True
        Grid_Image(174).Picture = BlockGraphic(1).Picture
        Grid_Filled(9, 16) = True
        Grid_Image(175).Picture = BlockGraphic(5).Picture
        Grid_Filled(10, 16) = True
        Grid_Image(176).Picture = BlockGraphic(5).Picture
        Grid_Filled(11, 16) = True
        Grid_Image(177).Picture = BlockGraphic(2).Picture
        Grid_Filled(1, 17) = True
        Grid_Image(178).Picture = BlockGraphic(2).Picture
        Grid_Filled(2, 17) = True
        Grid_Image(179).Picture = BlockGraphic(4).Picture
        Grid_Filled(3, 17) = True
        Grid_Image(185).Picture = BlockGraphic(7).Picture
        Grid_Filled(9, 17) = True
        Grid_Image(186).Picture = BlockGraphic(7).Picture
        Grid_Filled(10, 17) = True
        Grid_Image(187).Picture = BlockGraphic(5).Picture
        Grid_Filled(11, 17) = True
        Grid_Image(189).Picture = BlockGraphic(4).Picture
        Grid_Filled(2, 18) = True
        Grid_Image(190).Picture = BlockGraphic(4).Picture
        Grid_Filled(3, 18) = True
        Grid_Image(191).Picture = BlockGraphic(3).Picture
        Grid_Filled(4, 18) = True
        Grid_Image(197).Picture = BlockGraphic(7).Picture
        Grid_Filled(10, 18) = True
        Grid_Image(198).Picture = BlockGraphic(7).Picture
        Grid_Filled(11, 18) = True
        Grid_Image(199).Picture = BlockGraphic(3).Picture
        Grid_Filled(1, 19) = True
        Grid_Image(201).Picture = BlockGraphic(2).Picture
        Grid_Filled(3, 19) = True
        Grid_Image(204).Picture = BlockGraphic(2).Picture
        Grid_Filled(6, 19) = True
        Grid_Image(205).Picture = BlockGraphic(3).Picture
        Grid_Filled(7, 19) = True
        Grid_Image(209).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 19) = True
        Grid_Image(210).Picture = BlockGraphic(3).Picture
        Grid_Filled(1, 20) = True
        Grid_Image(211).Picture = BlockGraphic(3).Picture
        Grid_Filled(2, 20) = True
        Grid_Image(216).Picture = BlockGraphic(6).Picture
        Grid_Filled(7, 20) = True
        Grid_Image(217).Picture = BlockGraphic(7).Picture
        Grid_Filled(8, 20) = True
        Grid_Image(219).Picture = BlockGraphic(6).Picture
        Grid_Filled(10, 20) = True
        Grid_Image(220).Picture = BlockGraphic(6).Picture
        Grid_Filled(11, 20) = True
End Select
End Sub

