VERSION 5.00
Begin VB.Form frmMAIN 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Boxes (Verlet based approach for 2D game physics)"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   748
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar sFPS 
      Height          =   255
      Left            =   9720
      Max             =   100
      TabIndex        =   7
      Top             =   2880
      Value           =   80
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   9720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmMAIN.frx":0000
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox INFO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmMAIN.frx":00B0
      Top             =   6000
      Width           =   9135
   End
   Begin VB.Timer TimerFPS 
      Interval        =   2000
      Left            =   9360
      Top             =   720
   End
   Begin VB.CommandButton cmdGravYesNO 
      Caption         =   "GravityY - YesNO"
      Height          =   615
      Left            =   9960
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton CmdAddOBJ 
      Caption         =   "Add OBJ"
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   609
      TabIndex        =   1
      Top             =   120
      Width           =   9135
   End
   Begin VB.CommandButton cmdSTART 
      Caption         =   "RE-START"
      Height          =   615
      Left            =   9720
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Try to Keep Constant FPS"
      Height          =   615
      Left            =   9720
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lFPS 
      Height          =   495
      Left            =   9720
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'--------------------------------------------------------------------------------
'Original C++ code written by Benedikt Bitterli Copyright (c) 2009 [The code is released under the ZLib/LibPNG license]
'Original C++ code and tutorial available at links http://www.gamedev.net/reference/programming/features/verletPhys/default.asp
'http://www.gamedev.net/reference/articles/article2714.asp
'Forum:
'http://www.gamedev.net/community/forums/topic.asp?topic_id=553845

'Conversion from C++ to Java done by Craig Mitchell Copyright (c) 2010.
'Conversion from "C++ & Java" to VB6 done by Roberto Mior Copyright (c) 2010.

'Joints Handler By Roberto Mior


Option Explicit

Dim X
Dim Y
Dim S                  As String
Dim I                  As Long

Public Sub CmdAddOBJ_Click()


    If Rnd > 0.1 Then
        ADDBox (screenWidth \ 2), 10, _
               15 + Rnd * 30, 15 + Rnd * 30
    Else
        ADDTriangle screenWidth \ 2, 10 _
                                     , 15 + Rnd * 30, 15 + Rnd * 30
    End If

End Sub

Private Sub cmdGravYesNO_Click()
    If GravityY = 0 Then GravityY = 0.1 Else: GravityY = 0
End Sub

Private Sub cmdSTART_Click()
    screenWidth = PIC.Width - 1
    screenHeight = PIC.Height - 1


    NB = 2
    ReDim B(NB)



    X = 50
    Y = 50

    With B(1)
        .ADDPoint X, Y
        .ADDPoint X + 50, Y
        .ADDPoint X + 50, Y + 50

        .ADDEdge 1, 2
        .ADDEdge 2, 3
        .ADDEdge 3, 1
        .color = vbWhite
    End With

    X = 90
    Y = 180
    With B(2)



        .ADDPoint X, Y


        .ADDPoint X + 50, Y
        .ADDPoint X + 70, Y + 15

        .ADDPoint X + 50, Y + 30
        .ADDPoint X, Y + 30

        .ADDPoint X - 20, Y + 15

        .ADDEdge 1, 2
        .ADDEdge 2, 3
        .ADDEdge 3, 4
        .ADDEdge 4, 5
        .ADDEdge 5, 6
        .ADDEdge 6, 1

        .ADDEdge 2, 4, False
        .ADDEdge 1, 5, False

        .ADDEdge 2, 5, False
        .ADDEdge 1, 4, False

        .ADDEdge 3, 6, False


        .color = vbRed

    End With


    DuplicateOBJ 2, 90, 0

Skip:

    For I = 1 To 1
        ADDBox Rnd * (screenWidth - 100), Rnd * (screenHeight - 100), _
               15 + Rnd * 50, 15 + Rnd * 50
    Next

    ADDBox 200, 20, 80, 80


    'For I = 1 To B(2).NP
    'B(2).SetMASS(I) = InfiniteMASS
    'Next

    For I = 1 To NB
        B(I).DRAW PIC.hdc
    Next
    PIC.Refresh

    S = DetectCollision(1, 2) & vbCrLf
    S = S & "DEPTH " & vbTab & CollisionInfo.Depth & vbCrLf
    S = S & "OBJ1(Edge), E " & vbTab & CollisionInfo.OE & " " & CollisionInfo.WichEdgeOE & vbCrLf
    S = S & "OBJ2(Poin), P  " & vbTab & CollisionInfo.OP & " " & CollisionInfo.WichPointOP

    '    MsgBox S


    ADDjoint 2, 3, 3, 6      'RED ONES


    '******* This Adds & Joins 3 Boxes
    ADDBox 360, 40, 40, 40, True
    ADDBox 400, 40, 40, 40, True
    ADDBox 400, 80, 40, 80, True

    ADDjoint NB - 2, NB - 1, 2, 1
    ADDjoint NB - 2, NB - 1, 3, 4
    ADDjoint NB - 1, NB, 3, 2
    ADDjoint NB - 1, NB, 4, 1
    '*******


    '*** Thin Boxes
    ADDBox 350, 450, 4, 120, True
    DuplicateOBJ NB, 10, 0
    'DuplicateOBJ NB, 10, 0
    'DuplicateOBJ NB, 10, 0
    'DuplicateOBJ NB, 10, 0

    '***


    GravityY = 0.1
    MAINLOOP

End Sub

Private Sub Form_Activate()
    sFPS_Change

    PIC.Cls
    PIC.Height = PIC.Width * 0.618
    PIC.Refresh
    INFO.Width = PIC.Width


    cmdSTART_Click
    DoEvents
End Sub

Private Sub Form_Load()
    Randomize Timer
    Me.Caption = Me.Caption & " V" & App.Major & "." & App.Minor


End Sub

Private Sub Form_Unload(Cancel As Integer)
    End

End Sub


Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D              As Single
    Dim MinD           As Single
    MinD = 99999999999#
    For I = 1 To NB
        D = Distance(B(I).CenterX - X, B(I).CenterY - Y)
        If D < MinD Then
            MinD = D
            Omove = I
        End If
    Next
    Xmouse = X
    Ymouse = Y
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    If Omove <> 0 Then
        For I = 1 To B(Omove).NP
            B(Omove).SetPointX(I) = B(Omove).getPointX(I) + (X - B(Omove).getPointX(I)) * 0.01
            B(Omove).SetPointY(I) = B(Omove).getPointY(I) + (Y - B(Omove).getPointY(I)) * 0.01
        Next
    End If
    'Xmouse = X
    'Ymouse = Y

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Omove = 0
End Sub

Private Sub sFPS_Change()
    If sFPS <> 0 Then
        Label1 = "Try to Keep " & sFPS & " Constant FPS"
    Else
        Label1 = "Running at Max Speed"
    End If

End Sub

Private Sub sFPS_Scroll()
    If sFPS <> 0 Then
        Label1 = "Try to Keep " & sFPS & " Constant FPS"
    Else
        Label1 = "Running at Max Speed"
    End If

End Sub

Private Sub TimerFPS_Timer()

    lFPS = "FPS = " & (CNT - OldCNT) / (TimerFPS.Interval / 1000)
    lFPS = lFPS & "   Objs=" & NB
    OldCNT = CNT
    DoEvents

End Sub
