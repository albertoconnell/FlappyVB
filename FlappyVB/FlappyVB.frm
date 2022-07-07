VERSION 5.00
Begin VB.Form FlappyVB 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flappy VB6"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6750
   Icon            =   "FlappyVB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   16
      Left            =   6360
      Top             =   0
   End
   Begin VB.Frame PauseFrame 
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   1380
      TabIndex        =   0
      Top             =   1920
      Width           =   3975
      Begin VB.CommandButton AIGame 
         Caption         =   "COM"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Start 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   1
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label FinalScore 
         Caption         =   "Score: "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   3360
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "FLAPPY VB6"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label PipeLabel 
      BackColor       =   &H80000007&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape ScoreBox 
      Height          =   1695
      Index           =   1
      Left            =   5880
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape ScoreBox 
      Height          =   1695
      Index           =   0
      Left            =   5880
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label ScoreLabel 
      BackColor       =   &H80000007&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape PipeA 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   3960
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Shape PipeBT 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Shape PipeBT 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2535
      Index           =   2
      Left            =   1080
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape PipeBT 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Index           =   0
      Left            =   240
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape PipeB 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Shape PipeB 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   1080
      Top             =   6720
      Width           =   375
   End
   Begin VB.Shape PipeB 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Index           =   0
      Left            =   240
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Shape PipeAT 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   4920
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape PipeAT 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   3960
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Shape PipeAT 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Index           =   0
      Left            =   4080
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape PipeA 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   4920
      Top             =   6720
      Width           =   375
   End
   Begin VB.Shape PipeA 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Index           =   0
      Left            =   4080
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Flappy 
      BackColor       =   &H000000FF&
      Caption         =   "VB6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Cloud 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   5
      Left            =   5040
      Shape           =   2  'Oval
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Cloud 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   4
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   975
   End
   Begin VB.Shape Cloud 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   3
      Left            =   4800
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Cloud 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   2
      Left            =   840
      Shape           =   2  'Oval
      Top             =   720
      Width           =   1695
   End
   Begin VB.Shape Cloud 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   960
      Shape           =   2  'Oval
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Cloud 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   360
      Shape           =   2  'Oval
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   0
      Top             =   8040
      Width           =   6735
   End
End
Attribute VB_Name = "FlappyVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gameover, hit, pause, AR, BR As Boolean
Dim Score, Speed, FX, FY, AT, A, B, BT, C, AI, Pipes As Integer

Private Sub AIGame_Click()
    AI = True
    Call Start_Click
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        If Flappy.Top > 1000 Then
            Flappy.Top = Flappy.Top - 1000
            'Beep
        End If
    End If
End Sub

Private Sub Form_Load()
    Score = 0
    Gameover = False
    hit = False
    pause = True
    Speed = 50
    AR = False
    BR = False
    AI = False
    AI = False
    Flappy.Left = FlappyVB.Width / 3
    FlappyVB.Top = Screen.Height / 4
    FlappyVB.Left = Screen.Width / 6
End Sub
Private Sub CorrectHeight()
    Dim ADist, BDist As Double
    ADist = PipeA(1).Left - Flappy.Left
    BDist = PipeB(1).Left - Flappy.Left
    If Flappy.Top > FlappyVB.Height - 2250 Then
        Call Form_KeyPress(32)
    End If
    If ADist < 2500 And ADist > -1500 Then
        If Flappy.Top + 500 > A Then
            Call Form_KeyPress(32)
        End If
    End If
    If BDist < 2500 And BDist > -1500 Then
        If Flappy.Top + 500 > B Then
            Call Form_KeyPress(32)
        End If
    End If
End Sub



Private Sub Start_Click()
    'Show hidden controls
    PauseFrame.Visible = False
    ScoreLabel.Visible = True
    PipeLabel.Visible = True
    'Reset Booleans
    pause = False
    Gameover = False
    'Reset Score
    Score = 0
    Pipes = 0
    'Update Labels
    ScoreLabel.Caption = "Score:" + CStr(Score)
    PipeLabel.Caption = "Pipes: " + CStr(Pipes)
    'Position Label
    Flappy.Top = 1000
    Randomize
    'Create offset for gap between pipes
    Dim offset As Integer
    offset = Rnd * 3000 + 100
    'Initialize Pipe position
    For i = 0 To 2
        PipeAT(i).Left = PipeA(i).Left + 5000 + offset
        PipeBT(i).Left = PipeBT(i).Left + 5000 + Screen.Width + offset
        
        PipeA(i).Left = PipeA(i).Left + 5000 + offset
        PipeB(i).Left = PipeB(i).Left + 5000 + Screen.Width + offset
    Next
End Sub

Private Sub Timer1_Timer()
    If Not pause Then
        FY = Flappy.Top + 250
        If Flappy.Top < FlappyVB.Height - 2000 Then
            Flappy.Top = Flappy.Top + 75
        Else
        'Game over logic for falling on ground
            Gameover = True
        End If
        For i = 0 To Cloud.Count - 1
            Cloud(i).Left = Cloud(i).Left - 2
            If Cloud(i).Left < -500 Then
                Cloud(i).Left = FlappyVB.Width
            End If
        Next
        If Gameover Then
            pause = True
            PauseFrame.Visible = True
            FinalScore.Visible = True
            FinalScore.Caption = "Score: " & CStr(Score)
            AI = False
        End If
        Call MovePipe
        Call CheckClearance
        Call hitScore
        Pipes = Score / 4
        PipeLabel = "Pipes: " + CStr(Pipes)
        If AI Then
            Call CorrectHeight
        End If
    End If
End Sub
Private Function rotate(item As Variant) As Boolean
    If item.Left < -1000 Then
        item.Left = FlappyVB.Width + 2000
        rotate = True
    Else
        rotate = False
    End If
End Function
Private Sub MovePipe()
    For i = 0 To 2
        PipeA(i).Left = PipeA(i).Left - Speed
        PipeAT(i).Left = PipeAT(i).Left - Speed
        PipeB(i).Left = PipeB(i).Left - Speed
        PipeBT(i).Left = PipeBT(i).Left - Speed
        AR = rotate(PipeA(i))
        CR = rotate(PipeAT(i))
        BR = rotate(PipeB(i))
        BR = rotate(PipeBT(i))
        If AR Then
            Call NewSize(i, PipeA(i))
        End If
        If BR Then
            Call NewSize(i, PipeB(i))
        End If
    Next
    Call CorrectPipes
End Sub
Private Sub CorrectPipes()
    'A column of pipes correction
    PipeA(1).Top = PipeA(0).Top - 400
    PipeA(2).Top = PipeA(0).Top
    PipeA(2).Height = PipeA(0).Height
    PipeAT(0).Height = PipeA(1).Top - 2500
    PipeAT(1).Top = PipeAT(0).Height
    PipeAT(2).Height = PipeAT(0).Height
    ScoreBox(0).Left = PipeAT(1).Left + 1500
    ScoreBox(0).Top = PipeAT(1).Top + 500
    ScoreBox(0).Height = 2000
    'B column of pipes correction
    PipeB(1).Top = PipeB(0).Top - 400
    PipeB(2).Top = PipeB(0).Top
    PipeB(2).Height = PipeB(0).Height
    PipeBT(0).Height = PipeB(1).Top - 2500
    PipeBT(1).Top = PipeBT(0).Height
    PipeBT(2).Height = PipeBT(0).Height
    ScoreBox(1).Left = PipeBT(1).Left + 1500
    ScoreBox(1).Top = PipeBT(1).Top + 500
    ScoreBox(1).Height = 2000
End Sub
Private Sub NewSize(index As Variant, item As Variant)
    'Only adjust the size of the first element
    If index = 0 Then
        'Create temporary variables
        Dim num, change As Integer
        Randomize
        num = Rnd * 4000 + 1000
        If item.Height > num Then
            change = item.Height - num
            item.Height = num
            item.Top = item.Top + change
        Else
            change = num - item.Height
            item.Height = num
            item.Top = item.Top - change
        End If
        
    End If
End Sub
Private Sub CheckClearance()
    A = PipeA(1).Top
    AT = PipeAT(1).Top - 500
    B = PipeB(1).Top
    BT = PipeBT(1).Top - 500
    For i = 0 To 2
        If Flappy.Left > PipeA(1).Left And Flappy.Left < (PipeA(1).Left + 1500) Then
            If Flappy.Top > A Or Flappy.Top < AT Then
                Gameover = True
            End If
        End If
        If Flappy.Left > PipeB(1).Left And Flappy.Left < (PipeB(1).Left + 1500) Then
            If Flappy.Top > B Or Flappy.Top < BT Then
                Gameover = True
            End If
        End If
    Next
End Sub
Private Sub hitScore()
    For i = 0 To 1
        If Flappy.Left >= ScoreBox(i).Left And Flappy.Left <= ScoreBox(i).Left + 200 Then
            Score = Score + 1
            ScoreLabel.Caption = "Score: " + CStr(Score)
        End If
    Next
End Sub
