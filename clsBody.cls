VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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


Private P()            As tPoint
Public NP              As Long


Private E()            As tEdge
Public NE              As Long

Public CenterX         As Single
Public CenterY         As Single
Public MinX            As Single
Public MinY            As Single
Public MaxX            As Single
Public MaxY            As Single

Public Area            As Single
Public Perim           As Single
Public color           As Long


Public Property Get IsNOTJoint(WP As Long) As Boolean
    IsNOTJoint = P(WP).IsNOTJoint
End Property
Public Property Let IsNOTJoint(WP As Long, V As Boolean)
    P(WP).IsNOTJoint = V
End Property

Public Property Get getEdgeV1(wE As Long) As Long
    getEdgeV1 = E(wE).V1
End Property
Public Property Get getEdgeV2(wE As Long) As Long
    getEdgeV2 = E(wE).V2
End Property
Public Property Get getEdgeIsBoundary(wE As Long) As Boolean
    getEdgeIsBoundary = E(wE).Boundary
End Property
Public Property Get getPointX(WP As Long) As Single
    getPointX = P(WP).X
End Property
Public Property Get getPointY(WP As Long) As Single
    getPointY = P(WP).Y
End Property
Public Property Get getPointOldX(WP As Long) As Single
    getPointOldX = P(WP).OldX
End Property
Public Property Get getPointOldY(WP As Long) As Single
    getPointOldY = P(WP).OldY
End Property

Public Property Let SetPointX(WP As Long, V As Single)
    P(WP).X = V
End Property
Public Property Let SetPointY(WP As Long, V As Single)
    P(WP).Y = V
End Property

Public Property Get getMASS(WP As Long) As Single
    getMASS = P(WP).Mass
End Property
Public Property Let SetMASS(WP As Long, V As Single)
    P(WP).Mass = V
End Property
Public Sub UpDateVerlet()
    Dim I              As Long
    Dim tX             As Single
    Dim tY             As Single

    For I = 1 To NP

        With P(I)
            ' If .Mass <> InfiniteMASS Then
            tX = .X
            tY = .Y
            .X = .X + (.X - .OldX) * AirFriction + .AccX * TimeStep2
            .Y = .Y + (.Y - .OldY) * AirFriction + .AccY * TimeStep2
            .OldX = tX
            .OldY = tY
            ' End If
        End With
    Next


End Sub



Public Sub UpDateEdges()

    Dim I              As Long
    Dim v1v2X          As Single
    Dim v1v2Y          As Single
    Dim V1V2Length     As Single
    Dim Diff           As Single

    For I = 1 To NE
        With E(I)

            v1v2X = P(.V2).X - P(.V1).X
            v1v2Y = P(.V2).Y - P(.V1).Y
            V1V2Length = Distance(v1v2X, v1v2Y)
            Diff = (V1V2Length - .MainLength) * 0.5    '*.5

            Normalize v1v2X, v1v2Y

            'If .Boundary Then
            '    Diff = Diff * 1
            'Else
            '    Diff = Diff * 1
            'End If
            
            v1v2X = v1v2X * Diff
            v1v2Y = v1v2Y * Diff

            P(.V1).X = P(.V1).X + v1v2X '* Diff
            P(.V1).Y = P(.V1).Y + v1v2Y '* Diff
            P(.V2).X = P(.V2).X - v1v2X '* Diff
            P(.V2).Y = P(.V2).Y - v1v2Y '* Diff


        End With
    Next




End Sub
Public Sub ProjectToAxis(AxisX As Single, AxisY As Single, ByRef Min As Single, ByRef Max As Single)
    Dim DotP           As Single
    Dim I              As Long

    DotP = AxisX * P(1).X + AxisY * P(1).Y
    Min = DotP
    Max = DotP
    For I = 2 To NP
        With P(I)
            DotP = AxisX * .X + AxisY * .Y
            Min = MathMIN(DotP, Min)
            Max = MathMAX(DotP, Max)
        End With

    Next

End Sub

Public Sub CalculateCenter()
    Dim I              As Long

    CenterX = 0
    CenterY = 0

    MinX = 100000
    MinY = 100000
    MaxX = -100000
    MaxY = -100000

    For I = 1 To NP
        With P(I)
            CenterX = CenterX + .X
            CenterY = CenterY + .Y
            MinX = MathMIN(MinX, .X)
            MinY = MathMIN(MinY, .Y)
            MaxX = MathMAX(MaxX, .X)
            MaxY = MathMAX(MaxY, .Y)
        End With

    Next

    CenterX = CenterX / NP
    CenterY = CenterY / NP

End Sub


Public Sub ADDPoint(ByVal X, ByVal Y)
    NP = NP + 1
    ReDim Preserve P(NP)
    With P(NP)
        .X = X
        .Y = Y
        .OldX = X
        .OldY = Y
        .AccX = 0
        .AccY = 0
        .IsNOTJoint = True
    End With
    CalculateCenter
End Sub
Public Sub ADDEdge(ByVal P1 As Long, ByVal P2 As Long, Optional IsBoundary = True)
'Stop

    NE = NE + 1
    ReDim Preserve E(NE)
    With E(NE)
        .V1 = P1
        .V2 = P2
        .Boundary = IsBoundary
        .MainLength = Distance(P(P2).X - P(P1).X, P(P2).Y - P(P1).Y)
    End With

    ComputeMasses

End Sub

Public Sub UpDateForces()
    Dim I              As Long

    For I = 1 To NP
        P(I).AccX = GravityX
        P(I).AccY = GravityY
    Next
End Sub

Public Sub DRAW(PicHdc As Long)
    Dim I              As Long
    Dim X1             As Long
    Dim Y1             As Long
    Dim X2             As Long
    Dim Y2             As Long

    Dim pts()          As POINTAPI
    Dim PolyRegion     As Long
    Dim Brush          As Long



    '    For I = 1 To NE
    '        With E(I)
    '               If .Boundary Then
    '                X1 = P(.V1).X \ 1
    '                Y1 = P(.V1).Y \ 1
    '                X2 = P(.V2).X \ 1
    '                Y2 = P(.V2).Y \ 1
    '                FastLine PicHdc, X1, Y1, X2, Y2, 1, color    'vbWhite
    '                '            Else
    '                'X1 = P(.V1).X \ 1
    '                'Y1 = P(.V1).Y \ 1
    '                'X2 = P(.V2).X \ 1
    '                'Y2 = P(.V2).Y \ 1
    '                '
    '                '                FastLine PicHdc, X1, Y1, X2, Y2, 1, vbBlue
    '            End If
    '        End With
    '    Next

    ReDim pts(0 To NP - 1)
    For I = 1 To NP
        pts(I - 1).X = P(I).X \ 1
        pts(I - 1).Y = P(I).Y \ 1
    Next


    'Polygon PicHdc, pts(0), NP
    PolyRegion = CreatePolygonRgn(pts(0), NP, 1&)
    Brush = CreateSolidBrush(color)
    FillRgn PicHdc, PolyRegion, Brush

    DeleteObject PolyRegion
    DeleteObject Brush


End Sub




Private Sub ComputeMasses()


    Dim I              As Long

    CalculateCenter

    Area = 0
    For I = 1 To NP - 1
        Area = Area + ((P(I).X - MinX) * (P(I + 1).Y - MinY) _
                       - (P(I + 1).X - MinX) * (P(I).Y - MinY))
    Next
    Area = Area * 0.5

    Perim = 0
    For I = 1 To NE
        If E(I).Boundary Then Perim = Perim + E(I).MainLength
    Next


    For I = 1 To NP
        'P(I).Mass = Area / NP
        'P(I).Mass = Perim / NP

        P(I).Mass = Area
    Next

End Sub

Public Sub KeepInWorld()
    Dim I              As Long

    For I = 1 To NP
        With P(I)
            .X = MathMAX(MathMIN(.X, screenWidth), 0)
            .Y = MathMAX(MathMIN(.Y, screenHeight), 0)
        End With
    Next

End Sub
