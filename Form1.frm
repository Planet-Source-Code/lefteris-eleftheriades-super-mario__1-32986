VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7080
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox TimeUPpic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   1050
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   376
      TabIndex        =   11
      Top             =   2310
      Visible         =   0   'False
      Width           =   5640
   End
   Begin VB.PictureBox FlagPole 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   840
      Index           =   1
      Left            =   3900
      Picture         =   "Form1.frx":160C2
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   10
      Top             =   1050
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox FlagPole 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   840
      Index           =   0
      Left            =   3450
      Picture         =   "Form1.frx":17364
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   9
      Top             =   1050
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox AntPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   3450
      Picture         =   "Form1.frx":18606
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   8
      Top             =   540
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox AntPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   3450
      Picture         =   "Form1.frx":187B8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4140
      Top             =   3300
   End
   Begin VB.PictureBox Buffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   210
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   5
      Top             =   3750
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   240
      Picture         =   "Form1.frx":1A3A2
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox DigitalNosPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   2100
      Picture         =   "Form1.frx":1ADDC
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox DigitalNosMaskPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   2100
      Picture         =   "Form1.frx":2129E
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5250
      Top             =   3930
   End
   Begin VB.PictureBox MarioMaskPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   90
      Picture         =   "Form1.frx":21768
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox MarioPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   90
      Picture         =   "Form1.frx":21F6E
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   0
      Picture         =   "Form1.frx":2B748
      ScaleHeight     =   395
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   737
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   11055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim SuperMario As New CaracterObject
Dim Ant(5) As New CaracterObject
Dim TimeStatus As New DigitalCounter
Dim FrameStatus As New DigitalCounter
Dim TimerVal&
Dim StageX&, MarioY&
Dim ObjectMainPosX As Single
Dim FrameOT As Integer
Dim ExitFlag As Boolean
Const AscShift = 16
Const AscCtrl = 17
Const AscAlt = 18
Dim Secs&
Dim UpMode As Boolean
Dim JumpMode As Boolean
Dim Speed As Long
Dim Facing&
Dim I&, C&
Dim F As Long
Dim FF As Boolean
Dim Fast As Long
Dim AntFrame As Long
Dim AntX(10) As Long
Dim FlagRR As Boolean
Dim AFF As Boolean
Dim isInDieMode As Boolean
Dim ExitGame As Boolean
Private Sub Form_Load()

  SuperMario.SpriteDataFile = App.Path & "\Caracters.Spr"
  TimeStatus.SpriteDataFile = App.Path & "\Digital.Spr"
  FrameStatus.SpriteDataFile = App.Path & "\Digital.Spr"
  For I = 0 To 5
    Ant(I).SpriteDataFile = App.Path & "\AntN.Spr"
  Next I
  MarioY = 348
  Buffer.Width = Me.Width / 15
  Buffer.Height = Me.Height / 15
  Me.Show
  FrameOT = 5
  TimerVal = 101
  Secs& = 1
  Fast = 0
  AntX(0) = 500
  DoEvents
  DoEvents
  If MsgBox("Do you want this game to run through a Timer or a loop?" & vbCrLf & "Yes = Timer", vbYesNo Or vbDefaultButton2) = vbYes Then
    Timer2.Enabled = True
    GoTo Skip
  End If
    ''''''''''''''
    'Begin Loop
    Do Until ExitFlag
    'newDoEvents
    DoEvents '15 F.P.S
    'If KeyPressedCode <> 0 Then
     
     Buffer.Cls
     If StageX <= -Picture3.Width Then StageX = 0
     If StageX > 0 Then StageX = -Picture3.Width
     
     If ObjectMainPosX > 0 Then
        StageX = 0
        ObjectMainPosX = 0
     End If
     BitBlt Buffer.hdc, StageX, 0, Picture3.Width, Picture3.Height, Picture3.hdc, 0, 0, SRCCOPY
     BitBlt Buffer.hdc, StageX + Picture3.Width, 0, Picture3.Width, Picture3.Height, Picture3.hdc, 0, 0, SRCCOPY
     Buffer.Refresh
     newDoEvents
     'If we make the form > 640
     'BitBlt Buffer.hdc, StageX + (Picture3.Width * 2), 0, Picture3.Width, Picture3.Height, Picture3.hdc, 0, 0, SRCCOPY
     AFF = Not AFF
     If AFF Then
       AntFrame = AntFrame + 1
       If AntFrame >= 4 Then AntFrame = 1
       If AntX(0) < 500 Then FlagRR = True
       If AntX(0) > 700 Then FlagRR = False
       If FlagRR Then
         AntX(0) = AntX(0) + 5
       Else
         AntX(0) = AntX(0) - 5
       End If
     End If
     Ant(0).Draw AntFrame, AntX(0) + ObjectMainPosX, 365, Buffer.hdc, AntPic(0).hdc, AntPic(1).hdc
     
     For I = 0 To 42
         For C = 0 To 2
             DrawTile ObjectMainPosX + ((Tile.Width - 1) * I), Int((Tile.Height - 1) * C)
         Next C
     Next I
     F = F + 1
     
     FrameStatus.Value = Round(F / Secs&)
     FrameStatus.Draw (Me.Width / 15) - 30, 3, DigitalNosPic, DigitalNosMaskPic, Buffer, RedDC
     
     TimeStatus.Draw 3, 3, DigitalNosPic, DigitalNosMaskPic, Buffer, BlueDC
         
         If (AntX(0) < (-ObjectMainPosX + (Me.Width / 30))) And (AntX(0) > (-ObjectMainPosX + (Me.Width / 30)) - 20) And MarioY > 328 Then
            isInDieMode = True
         End If
         
         If isInDieMode Then
             MarioY = MarioY + 15
             JumpMode = False
             'SuperMario.Draw Facing, Me.Width / 30, MarioY, Buffer.hdc, MarioPic.hdc, MarioMaskPic.hdc
             If MarioY > 450 Then
               isInDieMode = False
               MarioY = 348
               ObjectMainPosX = 0
               StageX = 0
             End If
         End If
         If IsKeyDown(AscCtrl) Then
           'Fast
           Speed = 20
         Else
           'Normal
           Speed = 10
         End If
       If Not isInDieMode Then
         If IsKeyDown(vbKeyLeft) Then
             StageX = StageX + Speed
             ObjectMainPosX = ObjectMainPosX + Speed
             FF = Not FF
             If FF Then
               If FrameOT = 1 Then
                  FrameOT = 2
               Else
                  FrameOT = 1
               End If
             End If
             Facing = 3
         End If
         
         If IsKeyDown(vbKeyRight) Then
             StageX = StageX - Speed
             ObjectMainPosX = ObjectMainPosX - Speed
             FF = Not FF
             If FF Then
               If FrameOT = 5 Then
                  FrameOT = 6
               Else
                  FrameOT = 5
               End If
             End If
             Facing = 7
         End If
         If IsKeyDown(AscAlt) And MarioY >= 348 Then
          UpMode = True
          JumpMode = True
         End If
         End If
         If IsKeyDown(vbKeyEscape) Then
             ExitFlag = True
             ExitGame = True
             Exit Do
         End If
 
     MaskedBilt Buffer.hdc, FlagPole(0).hdc, FlagPole(1).hdc, ObjectMainPosX + (Me.Width / 30), 340, FlagPole(0).Width, FlagPole(0).Height, 0, 0

     If UpMode Then MarioY = MarioY - Speed
     If MarioY < 260 Then UpMode = False
     If (Not UpMode) And MarioY < 348 Then MarioY = MarioY + Speed
     If MarioY >= 348 Then JumpMode = False
     If JumpMode Then
        SuperMario.Draw Facing, Me.Width / 30, MarioY, Buffer.hdc, MarioPic.hdc, MarioMaskPic.hdc
     Else
        SuperMario.Draw FrameOT, Me.Width / 30, MarioY, Buffer.hdc, MarioPic.hdc, MarioMaskPic.hdc
     End If
     
     'DoEvents
     Buffer.Refresh
     Me.Cls
     BitBlt Me.hdc, 0, 0, Me.Width / 15, Me.Height / 15, Buffer.hdc, 0, 0, SRCCOPY
     Me.Refresh
     
 ' End If
  newDoEvents
  Loop
    'End loop
    ''''''''''''''
  If ExitGame Then End
Skip:
End Sub

Public Sub newDoEvents()
       Dim InputState&
       InputState& = GetInputState()
       If InputState& <> 0 Then DoEvents
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Resize.ReleaseCapture
  Resize.SendMessage Me.hwnd, Resize.WM_NCLBUTTONDOWN, Resize.HTMOVE, 0
End Sub

Private Sub Form_Resize()
  Buffer.Width = Me.Width / 15
  Buffer.Height = Me.Height / 15
End Sub

Function AsyncKeyHasBeenPressed(ByVal Key&) As Boolean
  Dim KS As Integer
  'Returns if the specified key
  'is the last one pressed and it is still
  'pressed.
  'If you combine two or more keys,
  'Returns the fist key pressed
  
  KS = GetAsyncKeyState(Key&)
  If KS = -32767 Then
     AsyncKeyHasBeenPressed = True
  Else
     AsyncKeyHasBeenPressed = False
  End If
End Function

Function IsKeyDown(ByVal Key&) As Boolean
  Dim KS As Integer
  'Returns if a key is press at the moment
  'Great for reading key combinations
  'E.g. Diagonal Up/Right
  'if GetKeyState(vbkeyUp) and GetKeyState(vbkeyRight) then ...
  
  KS = GetKeyState(Key&)
  If KS < 0 Then
     IsKeyDown = True
  Else
     IsKeyDown = False
  End If
End Function

Private Sub Timer1_Timer()
  TimerVal = TimerVal - 1
  If TimerVal < 0 Then
     TimerVal = 0
     ExitGame = False
     ExitFlag = True
     Me.Cls
     Me.PaintPicture TimeUPpic.Picture, ((Me.Width / 15) - TimeUPpic.Width) / 2, ((Me.Height / 15) - TimeUPpic.Height) / 2
  End If
  TimeStatus.Value = LeadingZeros(TimerVal, 3)
  Secs& = Secs& + 1
End Sub

Function LeadingZeros(Value As Variant, Zeros As Long) As String
 Dim OutV As String
 Dim Done As Boolean
 OutV = Trim(Str(Value))
 Done = False
 Do
  If Len(OutV) < Zeros Then
     OutV = "0" & OutV
  Else
     Done = True
  End If
 Loop Until Done
 LeadingZeros = OutV
End Function

Sub DrawTile(ByVal X&, ByVal Y&)
    BitBlt Buffer.hdc, X&, Picture3.Height + Y&, Tile.Width, Tile.Height, Tile.hdc, 0, 0, SRCCOPY
End Sub

Private Sub Timer2_Timer()
    newDoEvents
    'DoEvents 15 F.P.S
    'If KeyPressedCode <> 0 Then
     
     Buffer.Cls
     If StageX <= -Picture3.Width Then StageX = 0
     If StageX > 0 Then StageX = -Picture3.Width
     
     If ObjectMainPosX > 0 Then
        StageX = 0
        ObjectMainPosX = 0
     End If
     BitBlt Buffer.hdc, StageX, 0, Picture3.Width, Picture3.Height, Picture3.hdc, 0, 0, SRCCOPY
     BitBlt Buffer.hdc, StageX + Picture3.Width, 0, Picture3.Width, Picture3.Height, Picture3.hdc, 0, 0, SRCCOPY
     Buffer.Refresh
     newDoEvents
     'If we make the form > 640
     'BitBlt Buffer.hdc, StageX + (Picture3.Width * 2), 0, Picture3.Width, Picture3.Height, Picture3.hdc, 0, 0, SRCCOPY
     
     For I = 0 To 42
         For C = 0 To 0
             DrawTile ObjectMainPosX + ((Tile.Width - 1) * I), Int((Tile.Height - 1) * C)
         Next C
     Next I
     F = F + 1
     
     FrameStatus.Value = Round(F / Secs&)
     FrameStatus.Draw (Me.Width / 15) - 30, 3, DigitalNosPic, DigitalNosMaskPic, Buffer, RedDC
     
     TimeStatus.Draw 3, 3, DigitalNosPic, DigitalNosMaskPic, Buffer, BlueDC
         If IsKeyDown(AscCtrl) Then
           'Fast
           Speed = 20
         Else
           'Normal
           Speed = 10
         End If
     
         If IsKeyDown(vbKeyLeft) Then
             StageX = StageX + Speed
             ObjectMainPosX = ObjectMainPosX + Speed
             FF = Not FF
             If FF Then
               If FrameOT = 1 Then
                  FrameOT = 2
               Else
                  FrameOT = 1
               End If
             End If
             Facing = 3
         End If
         
         If IsKeyDown(vbKeyRight) Then
             StageX = StageX - Speed
             ObjectMainPosX = ObjectMainPosX - Speed
             FF = Not FF
             If FF Then
               If FrameOT = 5 Then
                  FrameOT = 6
               Else
                  FrameOT = 5
               End If
             End If
             Facing = 7
         End If
         If IsKeyDown(vbKeyEscape) Then
             ExitFlag = True
             End
         End If

     If IsKeyDown(AscAlt) And MarioY >= 348 Then
        UpMode = True
        JumpMode = True
     End If

    
     If UpMode Then MarioY = MarioY - (Speed / 2)
     If MarioY < 260 Then UpMode = False
     If (Not UpMode) And MarioY < 348 Then MarioY = MarioY + (Speed / 2)
     If MarioY >= 348 Then JumpMode = False
     If JumpMode Then
        SuperMario.Draw Facing, Me.Width / 30, MarioY, Buffer.hdc, MarioPic.hdc, MarioMaskPic.hdc
     Else
        SuperMario.Draw FrameOT, Me.Width / 30, MarioY, Buffer.hdc, MarioPic.hdc, MarioMaskPic.hdc
     End If
     'DoEvents
     Buffer.Refresh
     Me.Cls
     BitBlt Me.hdc, 0, 0, Me.Width / 15, Me.Height / 15, Buffer.hdc, 0, 0, SRCCOPY
     Me.Refresh
 ' End If
  newDoEvents
End Sub

Sub MaskedBilt(Buffer&, Source&, Mask&, X&, Y&, Width&, Height&, SrcX&, SrcY&)
    BitBlt Buffer, X, Y, Width, Height, Mask, SrcX, SrcY, SRCAND
    BitBlt Buffer, X, Y, Width, Height, Source, SrcX, SrcY, SRCINVERT
End Sub
