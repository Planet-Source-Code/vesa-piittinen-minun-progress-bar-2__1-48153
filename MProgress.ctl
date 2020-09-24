VERSION 5.00
Begin VB.UserControl MProgress 
   Alignable       =   -1  'True
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ScaleHeight     =   1095
   ScaleWidth      =   4815
   ToolboxBitmap   =   "MProgress.ctx":0000
   Begin VB.PictureBox DFG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   3
      Top             =   720
      Width           =   4695
      Visible         =   0   'False
   End
   Begin VB.PictureBox DBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   2
      Top             =   360
      Width           =   4695
      Visible         =   0   'False
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox FG 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   15
      End
   End
End
Attribute VB_Name = "MProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Minun Progress Bar version 2.0

'My First ActiveX (and also the fastest VB progress bar on PlanetSourceCode!)
'(well, Minun Progress Bar Deluxe is a bit faster, but it looks different)

'Last update on 31th of August 2003 (after 9 months pause)
'(project started in 14th November 2002)
'Author: Merri <vesa@merri.net>

'If you make changes, change UserControl name, project name...everything!
'Just to make sure your modified control doesn't mess up with the original :)
'It'll be best for both. If you find a "critical" bug, please mail me.
'You can also mail me if you find something that could be done better/faster

'The control is freeware, but it'd be nice to know if you spread your own
'version that is based on this :) It'd be also nice to know if you use the
'control in your project(s).

'I know there are some minor weird things that make it more confusing...
'it's because I've made changes in properties names and replaced Reverse
'and Vertical properties with a new Direction property. There might be
'also values not used at all... it's just too long a time I last worked
'with the project so I've lost track on what's in use and what is not...

'Sorry, no more commenting - observe it yourself

Option Explicit
Public Enum munBlend
    munBlendNone
    munBlendWeak
    munBlendAvarage
    munBlendStrong
    munBlendCosSin1
    munBlendCosSin2
    munBlendCosSin3
End Enum
Public Enum munBorderStyle
    munBStyleNone
    munBStyleSunken
    munBStyleSunkenOuter
    munBStyleRaised
    munBStyleRaisedInner
    munBStyleBump
    munBStyleEtched
End Enum
Public Enum munBorderWidth
    munBorderNone
    munBorderSingle
    munBorderDouble
End Enum
Public Enum munDirection
    munDirRight
    munDirUp
    munDirLeft
    munDirDown
End Enum
Public Enum munFade
    munFadeNone
    munFadeHorizontal
    munFadeVertical
End Enum
Public Enum munFadeStyle
    munFadeStill
    munFadeMoving
End Enum
Public Enum munPercentAlign
    munAlignCenter
    munAlignBarOut
    munAlignBarIn
    munAlignCenterOut
    munAlignCenterIn
    munAlignLeft
    munAlignRight
End Enum
'Constant Declarations:
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
'Default Property Values:
Const m_def_BackColor = &H8000000E
Const m_def_Blend = 0
Const m_def_BorderStyle = 6
Const m_def_Custom = 0
Const m_def_CustomText = "Done!"
Const m_def_Direction = 0
Const m_def_Fade = 2
Const m_def_FadeBG1 = &H80000010
Const m_def_FadeBG2 = &H80000014
Const m_def_FadeFG1 = &H8000000E
Const m_def_FadeFG2 = &H8000000D
Const m_def_FadeStyle = 1
Const m_def_ForeColor = &H80000012
Const m_def_Interval = 1
Const m_def_ManualRefresh = False
Const m_def_Max = 10000
Const m_def_Min = 0
Const m_def_NoPercent = False
Const m_def_Percent = 0
Const m_def_PercentAfter = "%"
Const m_def_PercentAlign = 0
Const m_def_PercentBefore = ""
Const m_def_Reverse = False
Const m_def_ScaleMode = vbTwips
Const m_def_Vertical = False
'Property Variables:
Dim m_BackColor As Long
Dim m_Blend As munBlend
Dim m_BorderStyle As munBorderStyle
Dim m_Custom As Boolean
Dim m_CustomText As String
Dim m_Direction2 As munDirection
Dim m_Fade As Byte
Dim m_FadeBG1 As Long
Dim m_FadeBG2 As Long
Dim m_FadeFG1 As Long
Dim m_FadeFG2 As Long
Dim m_FadeStyle As Byte
Dim m_Font As Font
Dim m_ForeColor As Long
Dim m_Interval As Integer
Dim m_ManualRefresh As Boolean
Dim m_Max As Currency
Dim m_Min As Currency
Dim m_NoPercent As Boolean
Dim m_Percent As Byte
Dim m_PercentAfter As String
Dim m_PercentAlign As Byte
Dim m_PercentBefore As String
Dim m_Reverse As Boolean
Dim m_ScaleMode As Integer
Dim m_Value As Currency
Dim m_Vertical As Boolean
'Internal Variables:
Dim m_Down As Byte
Dim m_BGborder As munBorderStyle
Dim m_FGborder As munBorderStyle
Dim m_ScaleWidth As Integer
Dim m_ScaleHeight As Integer
Dim m_TextWidth As Integer
Dim m_TextHeight As Integer
Dim m_TextNoChange As Boolean
Dim m_Scroll As Integer
Dim m_OldPercent As Byte
Dim Text As String
Dim Temp As Currency, TempMax As Currency
'API Declarations:
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetWindowParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Event Declarations:
Event BarResize()
Attribute BarResize.VB_Description = "Occurs when the progress bar of a control has resized."
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses the mouse button."
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event Resize()
Attribute Resize.VB_Description = "Occurs when the size of a control has changed."
Public Property Get Alignment() As munPercentAlign
    Alignment = m_PercentAlign
End Property
Public Property Let Alignment(ByVal New_PercentAlign As munPercentAlign)
    m_PercentAlign = New_PercentAlign
    PropertyChanged "PercentAlign"
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets object's background color."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    FG.ForeColor = New_BackColor
    BG.BackColor = New_BackColor
    PropertyChanged "BackColor"
    Draw
End Property
Public Property Get Blend() As munBlend
    Blend = m_Blend
End Property
Public Property Let Blend(ByVal New_Blend As munBlend)
    m_Blend = New_Blend
    PropertyChanged "Blend"
    Draw
    If Not m_ManualRefresh Then Refresh
End Property
Public Property Get BorderStyle() As munBorderStyle
Attribute BorderStyle.VB_Description = "Sets object's borderstyle."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle2 As munBorderStyle)
    Dim New_BorderStyle As munBorderStyle
    If New_BorderStyle2 > 6 Then
        New_BorderStyle = 6
    Else
        New_BorderStyle = New_BorderStyle2
    End If
    m_BorderStyle = New_BorderStyle
    EdgeUnSubClass UserControl.hwnd
    m_BGborder = EdgeSubClass(UserControl.hwnd, New_BorderStyle)
    PropertyChanged "BorderStyle"
    m_Down = Abs(m_BorderStyle = munBStyleNone Or m_BorderStyle = munBStyleSunken Or m_BorderStyle = munBStyleSunkenOuter)
    UserControl_Resize
    UserControl.Refresh
End Property
Public Property Get Caption() As String
    Caption = m_PercentBefore
End Property
Public Property Let Caption(ByVal New_PercentBefore As String)
    m_PercentBefore = New_PercentBefore
    PropertyChanged "PercentBefore"
    TextRefresh
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get Caption2() As String
    Caption2 = m_PercentAfter
End Property
Public Property Let Caption2(ByVal New_PercentAfter As String)
    m_PercentAfter = New_PercentAfter
    PropertyChanged "PercentAfter"
    TextRefresh
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get CustomCaption() As String
Attribute CustomCaption.VB_Description = "Sets object's custom ""Done!"" text."
Attribute CustomCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CustomCaption = m_CustomText
End Property
Public Property Let CustomCaption(ByVal New_CustomText As String)
    m_CustomText = New_CustomText
    PropertyChanged "CustomText"
    TextRefresh
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get Direction() As munDirection
    Direction = m_Direction2
End Property
Public Property Let Direction(ByVal New_Direction As munDirection)
    m_Direction2 = New_Direction
    On Error Resume Next
    Select Case m_Direction2
        Case 0
            m_Reverse = False
            m_Vertical = False
        Case 1
            m_Reverse = True
            m_Vertical = True
        Case 2
            m_Reverse = True
            m_Vertical = False
        Case 3
            m_Reverse = False
            m_Vertical = True
    End Select
    PropertyChanged "Direction"
    PropertyChanged "Reverse"
    PropertyChanged "Vertical"
    If Not m_ManualRefresh Then Refresh: DrawFade
End Property
Public Property Get DisablePercent() As Boolean
    DisablePercent = m_NoPercent
End Property
Public Property Let DisablePercent(ByVal New_NoPercent As Boolean)
    m_NoPercent = New_NoPercent
    PropertyChanged "NoPercent"
    If Not m_ManualRefresh Then Text = "": DrawFade
End Property
Private Sub Draw()
    Dim A As Integer, Col(5) As Single, Cols(2) As Single, Inc(2) As Single
    Dim Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long
    BG.BackColor = BG.BackColor
    FG.BackColor = FG.BackColor
    If m_FadeBG1 < 0 Or m_FadeBG1 > RGB(255, 255, 255) Then
        DBG.BackColor = m_FadeBG1
        Color1 = GetPixel(DBG.hdc, 1, 1)
    Else
        Color1 = m_FadeBG1
    End If
    If m_FadeBG2 < 0 Or m_FadeBG2 > RGB(255, 255, 255) Then
        DBG.BackColor = m_FadeBG2
        Color2 = GetPixel(DBG.hdc, 1, 1)
    Else
        Color2 = m_FadeBG2
    End If
    If m_FadeFG1 < 0 Or m_FadeFG1 > RGB(255, 255, 255) Then
        DBG.BackColor = m_FadeFG1
        Color3 = GetPixel(DBG.hdc, 1, 1)
    Else
        Color3 = m_FadeFG1
    End If
    If m_FadeFG2 < 0 Or m_FadeFG2 > RGB(255, 255, 255) Then
        DBG.BackColor = m_FadeFG2
        Color4 = GetPixel(DBG.hdc, 1, 1)
    Else
        Color4 = m_FadeFG2
    End If
    If m_Fade = 1 Then
        'background color
        Col(0) = Color1 \ 65536
        Col(1) = (Color1 Mod 65536) \ 256
        Col(2) = Color1 Mod 256
        Col(3) = Color2 \ 65536
        Col(4) = (Color2 Mod 65536) \ 256
        Col(5) = Color2 Mod 256
        Cols(0) = Col(0)
        Cols(1) = Col(1)
        Cols(2) = Col(2)
        Inc(0) = (Col(3) - Col(0)) / DBG.ScaleWidth
        Inc(1) = (Col(4) - Col(1)) / DBG.ScaleWidth
        Inc(2) = (Col(5) - Col(2)) / DBG.ScaleWidth
        For A = 0 To DBG.ScaleWidth
            Call SetPixel(DBG.hdc, A, 0, RGB(CByte(Cols(2)), CByte(Cols(1)), CByte(Cols(0))))
            GoSub GetTheColor
            If Cols(0) < 0 Then Cols(0) = 0
            If Cols(1) < 0 Then Cols(1) = 0
            If Cols(2) < 0 Then Cols(2) = 0
            If Cols(0) > 255 Then Cols(0) = 255
            If Cols(1) > 255 Then Cols(1) = 255
            If Cols(2) > 255 Then Cols(2) = 255
        Next A
        StretchBlt DBG.hdc, 0, 0, DBG.ScaleWidth, DBG.ScaleHeight, DBG.hdc, 0, 0, DBG.ScaleWidth, 1, vbSrcCopy
        'foreground color
        Col(0) = Color3 \ 65536
        Col(1) = (Color3 Mod 65536) \ 256
        Col(2) = Color3 Mod 256
        Col(3) = Color4 \ 65536
        Col(4) = (Color4 Mod 65536) \ 256
        Col(5) = Color4 Mod 256
        Cols(0) = Col(0)
        Cols(1) = Col(1)
        Cols(2) = Col(2)
        Inc(0) = (Col(3) - Col(0)) / DFG.ScaleWidth
        Inc(1) = (Col(4) - Col(1)) / DFG.ScaleWidth
        Inc(2) = (Col(5) - Col(2)) / DFG.ScaleWidth
        For A = 0 To DFG.ScaleWidth
            Call SetPixel(DFG.hdc, A, 0, RGB(CByte(Cols(2)), CByte(Cols(1)), CByte(Cols(0))))
            GoSub GetTheColor
            If Cols(0) < 0 Then Cols(0) = 0
            If Cols(1) < 0 Then Cols(1) = 0
            If Cols(2) < 0 Then Cols(2) = 0
            If Cols(0) > 255 Then Cols(0) = 255
            If Cols(1) > 255 Then Cols(1) = 255
            If Cols(2) > 255 Then Cols(2) = 255
        Next A
        StretchBlt DFG.hdc, 0, 0, DFG.ScaleWidth, DFG.ScaleHeight, DFG.hdc, 0, 0, DFG.ScaleWidth, 1, vbSrcCopy
    ElseIf m_Fade = 2 Then
        'background color
        Col(0) = Color1 \ 65536
        Col(1) = (Color1 Mod 65536) \ 256
        Col(2) = Color1 Mod 256
        Col(3) = Color2 \ 65536
        Col(4) = (Color2 Mod 65536) \ 256
        Col(5) = Color2 Mod 256
        Cols(0) = Col(0)
        Cols(1) = Col(1)
        Cols(2) = Col(2)
        Inc(0) = (Col(3) - Col(0)) / DBG.ScaleHeight
        Inc(1) = (Col(4) - Col(1)) / DBG.ScaleHeight
        Inc(2) = (Col(5) - Col(2)) / DBG.ScaleHeight
        For A = 0 To DBG.ScaleHeight
            Call SetPixel(DBG.hdc, 0, A, RGB(CByte(Cols(2)), CByte(Cols(1)), CByte(Cols(0))))
            GoSub GetTheColor
            If Cols(0) < 0 Then Cols(0) = 0
            If Cols(1) < 0 Then Cols(1) = 0
            If Cols(2) < 0 Then Cols(2) = 0
            If Cols(0) > 255 Then Cols(0) = 255
            If Cols(1) > 255 Then Cols(1) = 255
            If Cols(2) > 255 Then Cols(2) = 255
        Next A
        StretchBlt DBG.hdc, 0, 0, DBG.ScaleWidth, DBG.ScaleHeight, DBG.hdc, 0, 0, 1, DBG.ScaleHeight, vbSrcCopy
        'foreground color
        Col(0) = Color3 \ 65536
        Col(1) = (Color3 Mod 65536) \ 256
        Col(2) = Color3 Mod 256
        Col(3) = Color4 \ 65536
        Col(4) = (Color4 Mod 65536) \ 256
        Col(5) = Color4 Mod 256
        Cols(0) = Col(0)
        Cols(1) = Col(1)
        Cols(2) = Col(2)
        Inc(0) = (Col(3) - Col(0)) / DFG.ScaleHeight
        Inc(1) = (Col(4) - Col(1)) / DFG.ScaleHeight
        Inc(2) = (Col(5) - Col(2)) / DFG.ScaleHeight
        For A = 0 To DFG.ScaleHeight
            Call SetPixel(DFG.hdc, 0, A, RGB(CByte(Cols(2)), CByte(Cols(1)), CByte(Cols(0))))
            GoSub GetTheColor
            If Cols(0) < 0 Then Cols(0) = 0
            If Cols(1) < 0 Then Cols(1) = 0
            If Cols(2) < 0 Then Cols(2) = 0
            If Cols(0) > 255 Then Cols(0) = 255
            If Cols(1) > 255 Then Cols(1) = 255
            If Cols(2) > 255 Then Cols(2) = 255
        Next A
        StretchBlt DFG.hdc, 0, 0, DFG.ScaleWidth, DFG.ScaleHeight, DFG.hdc, 0, 0, 1, DFG.ScaleHeight, vbSrcCopy
    End If
    If Not m_ManualRefresh Then DrawFade
    Exit Sub
GetTheColor:
    Select Case m_Blend
        Case 0
            Cols(0) = Cols(0) + Inc(0)
            Cols(1) = Cols(1) + Inc(1)
            Cols(2) = Cols(2) + Inc(2)
        Case 1
            Cols(0) = Cols(0) * 0.25 + (Cols(0) + Inc(0)) * 0.75
            Cols(1) = Cols(1) * 0.25 + (Cols(1) + Inc(1)) * 0.75
            Cols(2) = Cols(2) * 0.25 + (Cols(2) + Inc(2)) * 0.75
        Case 2
            Cols(0) = Cols(0) / 2 + (Cols(0) + Inc(0)) / 2
            Cols(1) = Cols(1) / 2 + (Cols(1) + Inc(1)) / 2
            Cols(2) = Cols(2) / 2 + (Cols(2) + Inc(2)) / 2
        Case 3
            Cols(0) = Cols(0) * 0.75 + (Cols(0) + Inc(0)) * 0.25
            Cols(1) = Cols(1) * 0.75 + (Cols(1) + Inc(1)) * 0.25
            Cols(2) = Cols(2) * 0.75 + (Cols(2) + Inc(2)) * 0.25
        Case 4
            Cols(0) = Cols(0) + Inc(0) + Cos(Col(0)) + Sin(Col(3))
            Cols(1) = Cols(1) + Inc(1) + Cos(Col(1)) + Sin(Col(4))
            Cols(2) = Cols(2) + Inc(2) + Cos(Col(2)) + Sin(Col(5))
        Case 5
            Cols(0) = Cols(0) + Inc(0) + Cos(Col(0)) - Sin(Col(3))
            Cols(1) = Cols(1) + Inc(1) + Cos(Col(1)) - Sin(Col(4))
            Cols(2) = Cols(2) + Inc(2) + Cos(Col(2)) - Sin(Col(5))
        Case 6
            Cols(0) = Cols(0) + Inc(0) - Cos(Col(0)) + Sin(Col(3))
            Cols(1) = Cols(1) + Inc(1) - Cos(Col(1)) + Sin(Col(4))
            Cols(2) = Cols(2) + Inc(2) - Cos(Col(2)) + Sin(Col(5))
    End Select
    Return
End Sub
Private Sub DrawFade()
    Dim A As Integer, PrB As Boolean, PrF As Boolean
    TextRefresh
    If m_Fade > 0 Then
        Select Case m_Direction2
            Case 0
                A = FG.ScaleWidth + FG.Left
                Select Case m_FadeStyle
                    Case 1
                        BitBlt BG.hdc, A, 0, BG.ScaleWidth - A, BG.ScaleHeight, DBG.hdc, 0, 0, vbSrcCopy
                        BitBlt FG.hdc, 1, 0, FG.ScaleWidth - 1, BG.ScaleHeight, DFG.hdc, BG.ScaleWidth - FG.ScaleWidth + 1, 0, vbSrcCopy
                    Case Else
                        BitBlt BG.hdc, A, 0, BG.ScaleWidth - A, BG.ScaleHeight, DBG.hdc, A, 0, vbSrcCopy
                        BitBlt FG.hdc, 1, 0, FG.ScaleWidth - 1, BG.ScaleHeight, DFG.hdc, 0, 0, vbSrcCopy
                End Select
            Case 1
                A = FG.Top
                Select Case m_FadeStyle
                    Case 1
                        BitBlt BG.hdc, 0, 0, BG.ScaleWidth, A, DBG.hdc, 0, BG.ScaleHeight - A, vbSrcCopy
                        BitBlt FG.hdc, 0, 0, BG.ScaleWidth, FG.ScaleHeight, DFG.hdc, 0, 0, vbSrcCopy
                    Case Else
                        BitBlt BG.hdc, 0, 0, BG.ScaleWidth, A, DBG.hdc, 0, 0, vbSrcCopy
                        BitBlt FG.hdc, 0, 0, BG.ScaleWidth, FG.ScaleHeight, DFG.hdc, 0, A, vbSrcCopy
                End Select
                m_TextNoChange = False
            Case 2
                A = FG.Left
                Select Case m_FadeStyle
                    Case 1
                        BitBlt BG.hdc, 0, 0, A, BG.ScaleHeight, DBG.hdc, BG.ScaleWidth - A, 0, vbSrcCopy
                        BitBlt FG.hdc, 0, 0, FG.ScaleWidth, BG.ScaleHeight, DFG.hdc, 0, 0, vbSrcCopy
                    Case Else
                        BitBlt BG.hdc, 0, 0, A, BG.ScaleHeight, DBG.hdc, 0, 0, vbSrcCopy
                        BitBlt FG.hdc, 0, 0, FG.ScaleWidth, BG.ScaleHeight, DFG.hdc, A, 0, vbSrcCopy
                End Select
                m_TextNoChange = False
            Case 3
                A = FG.ScaleHeight + FG.Top
                Select Case m_FadeStyle
                    Case 1
                        BitBlt BG.hdc, 0, A, BG.ScaleWidth, BG.ScaleHeight - A, DBG.hdc, 0, 0, vbSrcCopy
                        BitBlt FG.hdc, 0, 1, BG.ScaleWidth, FG.ScaleHeight - 1, DFG.hdc, 0, BG.ScaleHeight - FG.ScaleHeight + 1, vbSrcCopy
                    Case Else
                        BitBlt BG.hdc, 0, A, BG.ScaleWidth, BG.ScaleHeight - A, DBG.hdc, 0, A, vbSrcCopy
                        BitBlt FG.hdc, 0, 1, BG.ScaleWidth, FG.ScaleHeight - 1, DFG.hdc, 0, 0, vbSrcCopy
                End Select
        End Select
    End If
    If Text <> "" Then
        Select Case m_Direction2
            Case 0
                BG.CurrentY = (BG.ScaleHeight - m_TextHeight) / 2 + m_Down
                FG.CurrentY = BG.CurrentY - FG.Top
                Select Case m_PercentAlign
                    Case 1
                        BG.CurrentX = FG.Width + m_Down
                        PrB = True
                    Case 2
                        FG.CurrentX = FG.ScaleWidth - m_TextWidth + m_Down
                        PrF = True
                    Case 3
                        BG.CurrentX = FG.Width + ((BG.ScaleWidth - FG.Width) - m_TextWidth) / 2 + m_Down
                        PrB = True
                    Case 4
                        FG.CurrentX = (FG.ScaleWidth - m_TextWidth) / 2 + m_Down
                        PrF = True
                    Case 5
                        BG.CurrentX = m_Down
                        FG.CurrentX = BG.CurrentX - FG.Left
                        PrB = True
                        PrF = True
                    Case 6
                        BG.CurrentX = BG.ScaleWidth - m_TextWidth + m_Down
                        FG.CurrentX = BG.CurrentX - FG.Left
                        PrB = True
                        PrF = True
                    Case Else
                        BG.CurrentX = (BG.ScaleWidth - m_TextWidth) / 2 + m_Down
                        FG.CurrentX = BG.CurrentX - FG.Left
                        PrB = True
                        PrF = True
                End Select
            Case 1
                BG.CurrentX = (BG.ScaleWidth - m_TextWidth) / 2 + m_Down
                FG.CurrentX = BG.CurrentX - FG.Left
                Select Case m_PercentAlign
                    Case 1
                        BG.CurrentY = FG.Top - m_TextHeight + m_Down
                        PrB = True
                    Case 2
                        FG.CurrentY = m_Down
                        PrF = True
                    Case 3
                        BG.CurrentY = (FG.Top - m_TextHeight) / 2 + m_Down
                        PrB = True
                    Case 4
                        FG.CurrentY = (FG.ScaleHeight - m_TextHeight) / 2 + m_Down
                        PrF = True
                    Case 5
                        BG.CurrentY = m_Down
                        FG.CurrentY = BG.CurrentY - FG.Top
                        PrB = True
                        PrF = True
                    Case 6
                        BG.CurrentY = BG.ScaleHeight - m_TextHeight + m_Down
                        FG.CurrentY = BG.CurrentY - FG.Top
                        PrB = True
                        PrF = True
                    Case Else
                        BG.CurrentY = (BG.ScaleHeight - m_TextHeight) / 2 + m_Down
                        FG.CurrentY = BG.CurrentY - FG.Top
                        PrB = True
                        PrF = True
                End Select
            Case 2
                BG.CurrentY = (BG.ScaleHeight - m_TextHeight) / 2 + m_Down
                FG.CurrentY = BG.CurrentY - FG.Top
                Select Case m_PercentAlign
                    Case 1
                        BG.CurrentX = FG.Left - m_TextWidth + m_Down
                        PrB = True
                    Case 2
                        FG.CurrentX = m_Down
                        PrF = True
                    Case 3
                        BG.CurrentX = (FG.Left - m_TextWidth) / 2 + m_Down
                        PrB = True
                    Case 4
                        FG.CurrentX = (FG.ScaleWidth - m_TextWidth) / 2 + m_Down
                        PrF = True
                    Case 5
                        BG.CurrentX = m_Down
                        FG.CurrentX = BG.CurrentX - FG.Left
                        PrB = True
                        PrF = True
                    Case 6
                        BG.CurrentX = BG.ScaleWidth - m_TextWidth + m_Down
                        FG.CurrentX = BG.CurrentX - FG.Left
                        PrB = True
                        PrF = True
                    Case Else
                        BG.CurrentX = (BG.ScaleWidth - m_TextWidth) / 2 + m_Down
                        FG.CurrentX = BG.CurrentX - FG.Left
                        PrB = True
                        PrF = True
                End Select
            Case 3
                BG.CurrentX = (BG.ScaleWidth - m_TextWidth) / 2 + m_Down
                FG.CurrentX = BG.CurrentX - FG.Left
                Select Case m_PercentAlign
                    Case 1
                        BG.CurrentY = FG.Height + m_Down
                        PrB = True
                    Case 2
                        FG.CurrentY = FG.ScaleHeight - m_TextHeight + m_Down
                        PrF = True
                    Case 3
                        BG.CurrentY = FG.Height + ((BG.ScaleHeight - FG.Height) - m_TextHeight) / 2 + m_Down
                        PrB = True
                    Case 4
                        FG.CurrentY = (FG.ScaleHeight - m_TextHeight) / 2 + m_Down
                        PrF = True
                    Case 5
                        BG.CurrentY = m_Down
                        FG.CurrentY = BG.CurrentY - FG.Top
                        PrB = True
                        PrF = True
                    Case 6
                        BG.CurrentY = BG.ScaleHeight - m_TextHeight + m_Down
                        FG.CurrentY = BG.CurrentY - FG.Top
                        PrB = True
                        PrF = True
                    Case Else
                        BG.CurrentY = (BG.ScaleHeight - m_TextHeight) / 2 + m_Down
                        FG.CurrentY = BG.CurrentY - FG.Top
                        PrB = True
                        PrF = True
                End Select
        End Select
        If Fade = 0 And Not m_TextNoChange Then
            BG.BackColor = BG.BackColor
            FG.BackColor = FG.BackColor
            If PrB Then BG.Print Text
            If PrF Then FG.Print Text
        ElseIf Fade > 0 Then
            If PrB Then BG.Print Text
            If PrF Then FG.Print Text
        End If
    End If
    BG.Refresh
    FG.Refresh
End Sub
Public Property Get Fade() As munFade
Attribute Fade.VB_Description = "Show colorfade?"
Attribute Fade.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Fade = m_Fade
End Property
Public Property Let Fade(ByVal New_Fade As munFade)
    m_Fade = New_Fade
    PropertyChanged "Fade"
    Draw
End Property
Public Property Get FadeBG1() As OLE_COLOR
Attribute FadeBG1.VB_Description = "Sets object's fade color one for background."
Attribute FadeBG1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeBG1 = m_FadeBG1
End Property
Public Property Let FadeBG1(ByVal New_FadeBG1 As OLE_COLOR)
    m_FadeBG1 = New_FadeBG1
    PropertyChanged "FadeBG1"
    Draw
End Property
Public Property Get FadeBG2() As OLE_COLOR
Attribute FadeBG2.VB_Description = "Sets object's fade color two for background."
Attribute FadeBG2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeBG2 = m_FadeBG2
End Property
Public Property Let FadeBG2(ByVal New_FadeBG2 As OLE_COLOR)
    m_FadeBG2 = New_FadeBG2
    PropertyChanged "FadeBG2"
    Draw
End Property
Public Property Get FadeFG1() As OLE_COLOR
Attribute FadeFG1.VB_Description = "Sets object's fade color one for foreground."
Attribute FadeFG1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeFG1 = m_FadeFG1
End Property
Public Property Let FadeFG1(ByVal New_FadeFG1 As OLE_COLOR)
    m_FadeFG1 = New_FadeFG1
    PropertyChanged "FadeFG1"
    Draw
End Property
Public Property Get FadeFG2() As OLE_COLOR
Attribute FadeFG2.VB_Description = "Sets object's fade color two for foreground."
Attribute FadeFG2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeFG2 = m_FadeFG2
End Property
Public Property Let FadeFG2(ByVal New_FadeFG2 As OLE_COLOR)
    m_FadeFG2 = New_FadeFG2
    PropertyChanged "FadeFG2"
    Draw
End Property
Public Property Get FadeStyle() As munFadeStyle
Attribute FadeStyle.VB_Description = "Sets object's color fading style."
Attribute FadeStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeStyle = m_FadeStyle
End Property
Public Property Let FadeStyle(ByVal New_FadeStyle As munFadeStyle)
    If New_FadeStyle > 1 Then New_FadeStyle = 1
    m_FadeStyle = New_FadeStyle
    PropertyChanged "FadeStyle"
    Draw
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Sets object's font."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    Set BG.Font = New_Font
    Set FG.Font = New_Font
    PropertyChanged "Font"
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets object's foreground color."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    BG.ForeColor = New_ForeColor
    FG.BackColor = New_ForeColor
    PropertyChanged "ForeColor"
    Draw
End Property
Public Property Get ManualRefresh() As Boolean
Attribute ManualRefresh.VB_Description = "Refresh must be invoked manually through code?"
Attribute ManualRefresh.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ManualRefresh = m_ManualRefresh
End Property
Public Property Let ManualRefresh(ByVal New_ManualRefresh As Boolean)
    m_ManualRefresh = New_ManualRefresh
    PropertyChanged "ManualRefresh"
End Property
Public Property Get Max() As Currency
Attribute Max.VB_Description = "Sets object's max value."
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Misc"
    Max = m_Max
End Property
Public Property Let Max(ByVal New_Max As Currency)
    If New_Max < m_Min + 1 Then New_Max = m_Min + 1
    m_Max = New_Max
    PropertyChanged "Max"
    If m_Max < m_Value Then m_Value = m_Max
    If Not m_ManualRefresh Then Refresh: DrawFade
End Property
Public Property Get Min() As Currency
Attribute Min.VB_Description = "Sets object's minimal value."
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Misc"
    Min = m_Min
End Property
Public Property Let Min(ByVal New_Min As Currency)
    If New_Min > m_Max - 1 Then New_Min = m_Max - 1
    m_Min = New_Min
    PropertyChanged "Min"
    If m_Min > m_Value Then m_Value = m_Min
    If Not m_ManualRefresh Then Refresh: DrawFade
End Property
Public Property Get Percent() As Byte
Attribute Percent.VB_Description = "Returns percent value."
Attribute Percent.VB_MemberFlags = "400"
    Percent = m_Percent
End Property
Public Property Let Percent(ByVal New_Percent As Byte)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh object."
    Dim A As Integer
    On Error Resume Next
    Select Case m_Direction2
        Case 0
            If Temp <> 0 Then A = (BG.ScaleWidth + 1) * Temp / TempMax
            FG.Move -1, 0, A, BG.ScaleHeight
        Case 1
            If Temp <> 0 Then A = BG.ScaleHeight * Temp / TempMax
            FG.Move 0, BG.ScaleHeight - A, BG.ScaleWidth, A
        Case 2
            If Temp <> 0 Then A = BG.ScaleWidth * Temp / TempMax
            FG.Move BG.ScaleWidth - A, 0, A, BG.ScaleHeight
        Case 3
            If Temp <> 0 Then A = (BG.ScaleHeight + 1) * Temp / TempMax
            FG.Move 0, -1, BG.ScaleWidth, A
    End Select
    If m_ManualRefresh Then DrawFade
End Sub
Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "Returns/sets object's scalemode."
Attribute ScaleMode.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleMode = m_ScaleMode
End Property
Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
    m_ScaleMode = New_ScaleMode
    UserControl.ScaleMode = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property
Public Sub SetParent(ByVal hwnd As Long)
Attribute SetParent.VB_Description = "Sets object's parent."
    SetWindowParent UserControl.hwnd, hwnd
End Sub
Public Property Get ShowCustomCaption() As Boolean
    ShowCustomCaption = m_Custom
End Property
Public Property Let ShowCustomCaption(ByVal New_Custom As Boolean)
    m_Custom = New_Custom
    PropertyChanged "Custom"
    If Not m_ManualRefresh Then DrawFade
End Property
Private Sub TextRefresh()
    Dim OldText As String
    OldText = Text
    If Not m_NoPercent Then
        If m_Custom And m_Value = m_Max Then
            Text = m_CustomText
        Else
            Text = m_PercentBefore & Chr(32) & m_Percent & Chr(32) & m_PercentAfter
        End If
    ElseIf m_PercentBefore <> "" Then
        Text = m_PercentBefore
    End If
    If m_Fade = 0 Then
        If OldText = Text Then
            m_TextNoChange = True
            Exit Sub
        Else
            m_TextNoChange = False
        End If
    End If
    m_TextWidth = BG.TextWidth(Text)
    m_TextHeight = BG.TextHeight(Text)
End Sub
Public Property Get Value() As Currency
Attribute Value.VB_Description = "Sets object's position value."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Currency)
    Dim A As Byte
    On Error Resume Next
    If New_Value < m_Min Then New_Value = m_Min
    If New_Value > m_Max Then New_Value = m_Max
    If m_Value <> New_Value Then
        m_Value = New_Value
        Temp = m_Value - m_Min
        TempMax = m_Max - m_Min
        m_Percent = Fix(Temp / TempMax * 100)
        PropertyChanged "Value"
        If Not m_ManualRefresh Then Refresh
        If m_Percent <> m_OldPercent Then
            m_OldPercent = m_Percent
            If Not m_ManualRefresh And Not m_NoPercent Then DrawFade
            RaiseEvent Change
        End If
        If m_Value = m_Max Or m_Value = m_Min Then
            If Not m_ManualRefresh Then DrawFade
            RaiseEvent Change
        End If
    Else
        PropertyChanged "Value"
    End If
End Property
Private Sub FG_Resize()
    If Not m_ManualRefresh Then DrawFade
    RaiseEvent BarResize
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_Blend = m_def_Blend
    m_BorderStyle = m_def_BorderStyle
    m_Custom = m_def_Custom
    m_CustomText = m_def_CustomText
    m_Fade = m_def_Fade
    m_FadeBG1 = m_def_FadeBG1
    m_FadeBG2 = m_def_FadeBG2
    m_FadeFG1 = m_def_FadeFG1
    m_FadeFG2 = m_def_FadeFG2
    m_FadeStyle = m_def_FadeStyle
    Set m_Font = Ambient.Font
    Set BG.Font = Ambient.Font
    Set FG.Font = Ambient.Font
    m_ForeColor = m_def_ForeColor
    m_Interval = m_def_Interval
    m_ManualRefresh = m_def_ManualRefresh
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_NoPercent = m_def_NoPercent
    m_Percent = m_def_Percent
    m_PercentAfter = m_def_PercentAfter
    m_PercentBefore = m_def_PercentBefore
    m_Reverse = m_def_Reverse
    m_ScaleMode = m_def_ScaleMode
    m_PercentAlign = m_def_PercentAlign
    m_Vertical = m_def_Vertical
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Blend = PropBag.ReadProperty("Blend", m_def_Blend)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Custom = PropBag.ReadProperty("Custom", m_def_Custom)
    m_CustomText = PropBag.ReadProperty("CustomText", m_def_CustomText)
    m_Direction2 = PropBag.ReadProperty("Direction", m_def_Direction)
    m_Fade = PropBag.ReadProperty("Fade", m_def_Fade)
    m_FadeBG1 = PropBag.ReadProperty("FadeBG1", m_def_FadeBG1)
    m_FadeBG2 = PropBag.ReadProperty("FadeBG2", m_def_FadeBG2)
    m_FadeFG1 = PropBag.ReadProperty("FadeFG1", m_def_FadeFG1)
    m_FadeFG2 = PropBag.ReadProperty("FadeFG2", m_def_FadeFG2)
    m_FadeStyle = PropBag.ReadProperty("FadeStyle", m_def_FadeStyle)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Interval = PropBag.ReadProperty("Interval", m_def_Interval)
    m_ManualRefresh = PropBag.ReadProperty("ManualRefresh", m_def_ManualRefresh)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_NoPercent = PropBag.ReadProperty("NoPercent", m_def_NoPercent)
    m_Percent = PropBag.ReadProperty("Percent", m_def_Percent)
    m_PercentAfter = PropBag.ReadProperty("PercentAfter", m_def_PercentAfter)
    m_PercentAlign = PropBag.ReadProperty("PercentAlign", m_def_PercentAlign)
    m_PercentBefore = PropBag.ReadProperty("PercentBefore", m_def_PercentBefore)
    m_Reverse = PropBag.ReadProperty("Reverse", m_def_Reverse)
    m_ScaleMode = PropBag.ReadProperty("ScaleMode", m_def_ScaleMode)
    m_Value = PropBag.ReadProperty("Value", 0)
    m_Vertical = PropBag.ReadProperty("Vertical", m_def_Vertical)
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    DBG.Move 0, 0, ScaleWidth, ScaleHeight
    DFG.Move 0, 0, ScaleWidth, ScaleHeight
    BG.Move 0, 0, ScaleWidth, ScaleHeight
    m_ScaleWidth = ScaleWidth
    m_ScaleHeight = ScaleHeight
    If Not m_ManualRefresh Then Refresh
    Draw
    RaiseEvent Resize
End Sub
Private Sub UserControl_Show()
    BG.BackColor = m_BackColor
    FG.BackColor = m_ForeColor
    BG.ForeColor = m_ForeColor
    FG.ForeColor = m_BackColor
    Set BG.Font = m_Font
    Set FG.Font = m_Font
    Temp = m_Value - m_Min
    TempMax = m_Max - m_Min
    m_BGborder = EdgeSubClass(UserControl.hwnd, m_BorderStyle)
    m_Down = Abs(m_BorderStyle = munBStyleNone Or m_BorderStyle = munBStyleSunken Or m_BorderStyle = munBStyleSunkenOuter)
    TextRefresh
    UserControl_Resize
End Sub
Private Sub UserControl_Terminate()
    EdgeUnSubClass UserControl.hwnd
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Blend", m_Blend, m_def_Blend)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Custom", m_Custom, m_def_Custom)
    Call PropBag.WriteProperty("CustomText", m_CustomText, m_def_CustomText)
    Call PropBag.WriteProperty("Direction", m_Direction2, m_def_Direction)
    Call PropBag.WriteProperty("Fade", m_Fade, m_def_Fade)
    Call PropBag.WriteProperty("FadeBG1", m_FadeBG1, m_def_FadeBG1)
    Call PropBag.WriteProperty("FadeBG2", m_FadeBG2, m_def_FadeBG2)
    Call PropBag.WriteProperty("FadeFG1", m_FadeFG1, m_def_FadeFG1)
    Call PropBag.WriteProperty("FadeFG2", m_FadeFG2, m_def_FadeFG2)
    Call PropBag.WriteProperty("FadeStyle", m_FadeStyle, m_def_FadeStyle)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Interval", m_Interval, m_def_Interval)
    Call PropBag.WriteProperty("ManualRefresh", m_ManualRefresh, m_def_ManualRefresh)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("NoPercent", m_NoPercent, m_def_NoPercent)
    Call PropBag.WriteProperty("Percent", m_Percent, m_def_Percent)
    Call PropBag.WriteProperty("PercentAfter", m_PercentAfter, m_def_PercentAfter)
    Call PropBag.WriteProperty("PercentAlign", m_PercentAlign, m_def_PercentAlign)
    Call PropBag.WriteProperty("PercentBefore", m_PercentBefore, m_def_PercentBefore)
    Call PropBag.WriteProperty("Reverse", m_Reverse, m_def_Reverse)
    Call PropBag.WriteProperty("ScaleMode", m_ScaleMode, m_def_ScaleMode)
    Call PropBag.WriteProperty("Value", m_Value, 0)
    Call PropBag.WriteProperty("Vertical", m_Vertical, m_def_Vertical)
End Sub
