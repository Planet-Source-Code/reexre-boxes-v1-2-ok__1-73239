Attribute VB_Name = "modWorld"
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

Public Type tPoint
    X                  As Single
    Y                  As Single
    OldX               As Single
    OldY               As Single
    AccX               As Single
    AccY               As Single


    Mass               As Single


    IsNOTJoint         As Boolean


End Type

Public Type tEdge
    V1                 As Long
    V2                 As Long
    MainLength         As Single
    Boundary           As Boolean
End Type



Public Const TimeStep  As Single = 1
Public Const TimeStep2 As Single = TimeStep * TimeStep

Public NB              As Long
Public B()             As New clsBody


Public screenWidth     As Single
Public screenHeight    As Single

Public GravityX        As Single
Public GravityY        As Single

Public Const AirFriction = 0.9999

Public Const PI = 3.14159265358979

Public Const InfiniteMASS = 1E+99

Public CNT             As Long
Public OldCNT          As Long

Public Omove           As Long
Public Xmouse          As Single
Public Ymouse          As Single

Private cTime          As Double
Private pTime          As Double


Public Function Distance(Dx As Single, Dy As Single) As Single
    Distance = Sqr(Dx * Dx + Dy * Dy)
End Function
Public Function DistanceSQ(Dx As Single, Dy As Single) As Single
    DistanceSQ = (Dx * Dx + Dy * Dy)
End Function
Public Sub Normalize(ByRef X As Single, ByRef Y As Single)
    Dim L              As Single
    L = Sqr(X * X + Y * Y)
    If L <> 0 Then L = 1 / L    'Else: Stop

    X = X * L
    Y = Y * L
End Sub


Public Function MathMIN(ByRef A As Single, ByRef B As Single) As Single
    MathMIN = IIf(A < B, A, B)
End Function
Public Function MathMAX(ByRef A As Single, ByRef B As Single) As Single
    MathMAX = IIf(A > B, A, B)
End Function




Public Function IntervalDistance(ByRef MinA As Single, ByRef MaxA As Single, _
                                 ByRef MinB As Single, ByRef MaxB As Single) As Single
    If MinA < MinB Then
        IntervalDistance = MinB - MaxA
    Else
        IntervalDistance = MinA - MaxB
    End If


    '    IntervalDistance( float MinA, float MaxA, float MinB, float MaxB ) {
    '    if( MinA < MinB )
    '        return MinB - MaxA;
    '    Else
    '        return MinA - MaxB;'


End Function



Public Sub MAINLOOP()
    Const OneMillisec  As Long = 1

    Dim I              As Long
    Dim InvFPS         As Double


    Timing = 0
    pTime = Timing
    Do


        '***** Keep Constant FPS
        If frmMAIN.sFPS <> 0 Then
            InvFPS = 1 / frmMAIN.sFPS
            Do
                'Sleep OneMillisec
            Loop While (Timing < (pTime + InvFPS))
            pTime = Timing
        End If
        '****************


        BitBlt frmMAIN.PIC.hdc, 0, 0, frmMAIN.PIC.ScaleWidth, frmMAIN.PIC.ScaleHeight, frmMAIN.PIC.hdc, 0, 0, vbBlack    'ness
        For I = 1 To NB
            B(I).DRAW frmMAIN.PIC.hdc
        Next
        frmMAIN.PIC.Refresh
        DoEvents




        For I = 1 To NB

            B(I).UpDateForces
            B(I).UpDateVerlet

        Next



        IterateCollisions

        DoEvents

        CNT = CNT + 1
        If CNT Mod 200 = 0 Then frmMAIN.CmdAddOBJ_Click


    Loop While True

End Sub


Public Sub ADDBox(X, Y, W, H, Optional Perfect As Boolean = False)
    NB = NB + 1
    ReDim Preserve B(NB)

    With B(NB)

        If Perfect Then
            .ADDPoint X, Y
            .ADDPoint X + W, Y
            .ADDPoint X + W, Y + H
            .ADDPoint X, Y + H
        Else
            .ADDPoint X, Y
            .ADDPoint X + W + Rnd * 10, Y
            .ADDPoint X + W + Rnd * 10, Y + H + Rnd * 10
            .ADDPoint X, Y + H + Rnd * 10
        End If

        .ADDEdge 1, 2
        .ADDEdge 2, 3
        .ADDEdge 3, 4
        .ADDEdge 4, 1
        .ADDEdge 2, 4, False
        .ADDEdge 1, 3, False
        .color = RGB(80 + Rnd * 175, 80 + Rnd * 175, 80 + Rnd * 175)

    End With
End Sub

Public Sub ADDTriangle(X, Y, W, H)
    NB = NB + 1
    ReDim Preserve B(NB)

    With B(NB)
        .ADDPoint X, Y
        .ADDPoint X + W + Rnd * 10, Y
        .ADDPoint X + W + Rnd * 10, Y + H + Rnd * 10

        .ADDEdge 1, 2
        .ADDEdge 2, 3
        .ADDEdge 3, 1
        .color = RGB(80 + Rnd * 175, 80 + Rnd * 175, 80 + Rnd * 175)

    End With
End Sub

Public Sub DuplicateOBJ(ByVal wO, Dx, Dy)
    Dim I              As Long

    NB = NB + 1
    ReDim Preserve B(NB)

    With B(NB)
        For I = 1 To B(wO).NP
            .ADDPoint B(wO).getPointX(I) + Dx, B(wO).getPointY(I) + Dy
        Next
        For I = 1 To B(wO).NE
            .ADDEdge B(wO).getEdgeV1(I), B(wO).getEdgeV2(I), B(wO).getEdgeIsBoundary(I)
            .color = B(wO).color
        Next
    End With
End Sub
