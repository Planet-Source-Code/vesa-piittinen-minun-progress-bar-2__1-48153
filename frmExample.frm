VERSION 5.00
Object = "{342EF3A4-3372-4702-8BB2-7391C01823D1}#27.0#0"; "MProgress2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minun Progress Bar 2 example"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "Speed comparison:"
      Height          =   975
      Left            =   720
      TabIndex        =   11
      Top             =   3600
      Width           =   8055
      Begin VB.CommandButton butCompare 
         Caption         =   "Test!"
         Height          =   255
         Left            =   5640
         TabIndex        =   15
         Top             =   600
         Width           =   2295
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   10000
         Scrolling       =   1
      End
      Begin MinunProgressBar2.MProgress MProgress1 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoPercent       =   -1  'True
      End
      Begin VB.Label Comparison 
         Alignment       =   2  'Center
         Caption         =   "Result appears here"
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   240
         Width           =   2295
      End
   End
   Begin MinunProgressBar2.MProgress Bar 
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PercentAlign    =   3
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1815
      ScaleWidth      =   8655
      TabIndex        =   1
      Top             =   600
      Width           =   8655
      Begin MinunProgressBar2.MProgress Bar 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   661
         BorderStyle     =   0
         Fade            =   1
         FadeBG1         =   128
         FadeBG2         =   192
         FadeFG1         =   128
         FadeFG2         =   255
         FadeStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoPercent       =   -1  'True
      End
      Begin VB.Label Txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "My Cool Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Txt 
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000040&
         X1              =   7560
         X2              =   3480
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000040&
         X1              =   840
         X2              =   4920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000040&
         X1              =   7560
         X2              =   7560
         Y1              =   1200
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000040&
         X1              =   840
         X2              =   840
         Y1              =   600
         Y2              =   240
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000040&
         Height          =   615
         Left            =   120
         Top             =   600
         Width           =   8415
      End
      Begin VB.Label Txt 
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Index           =   2
         Left            =   4920
         TabIndex        =   5
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "My Cool Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Index           =   3
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   3015
      End
   End
   Begin MinunProgressBar2.MProgress Bar 
      Align           =   1  'Align Top
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   873
      BorderStyle     =   0
      Fade            =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NoPercent       =   -1  'True
      PercentAlign    =   5
      PercentBefore   =   " Minun Progress Bar 2 is the best progress bar on PSC"
   End
   Begin MinunProgressBar2.MProgress Bar 
      Height          =   2415
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   4260
      BorderStyle     =   1
      Direction       =   1
      Fade            =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PercentAlign    =   1
      Reverse         =   -1  'True
      Vertical        =   -1  'True
   End
   Begin MinunProgressBar2.MProgress Bar 
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   9
      Top             =   3000
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   873
      BackColor       =   12640511
      BorderStyle     =   3
      Fade            =   1
      FadeBG1         =   8388608
      FadeBG2         =   16744576
      FadeFG1         =   16576
      FadeFG2         =   12640511
      FadeStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16761024
      PercentAfter    =   "% and I'm happy 'bout it"
      PercentBefore   =   "It's going"
   End
   Begin VB.Label Txt 
      Alignment       =   1  'Right Justify
      Caption         =   "If you like it, please rate it :)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   10
      Top             =   4680
      Width           =   8055
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Bar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bar_MouseMove Index, Button, Shift, X, Y
End Sub

Private Sub Bar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Select Case Bar(Index).Direction
        Case munDirRight
            Bar(Index).Value = X
        Case munDirLeft
            Bar(Index).Value = Bar(Index).Width - X
        Case munDirDown
            Bar(Index).Value = Y
        Case munDirUp
            Bar(Index).Value = Bar(Index).Height - Y
    End Select
End Sub

Private Sub butCompare_Click()
    Dim Start1 As Long, Start2 As Long, End1 As Long, End2 As Long
    Dim A As Integer
    Start1 = GetTickCount
    For A = 0 To 10000
        MProgress1.Value = A
    Next A
    End1 = GetTickCount
    Start2 = GetTickCount
    For A = 0 To 10000
        ProgressBar1.Value = A
    Next A
    End2 = GetTickCount
    Comparison = (End1 - Start1) & "ms vs " & (End2 - Start2) & "ms"
End Sub

Private Sub Form_Load()
    Dim A As Integer
    For A = 0 To Bar.Count - 1
        Select Case Bar(A).Direction
            Case munDirLeft, munDirRight
                Bar(A).Max = Bar(A).Width
            Case Else
                Bar(A).Max = Bar(A).Height
        End Select
    Next A
End Sub
