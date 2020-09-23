Attribute VB_Name = "Declarations"
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const Block_Length = 240
Public Const Number_of_Columns = 11
Public Const Number_of_Rows = 20
Public Next_Flag As Boolean

Public Const Middle_Start_Up_X = 6
Public Const Number_of_Levels = 9
Public Speed As Integer
Public Speed_Value(1 To 9)  As Integer

Public Score As Double
Public Lines As Double
Public Level As Integer
Public Game_End As Boolean
Public Marquee_Label_Left As Integer

Public Points_Value As Integer

