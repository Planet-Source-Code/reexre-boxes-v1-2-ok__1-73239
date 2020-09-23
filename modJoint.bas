Attribute VB_Name = "modJoint"
'Author :Roberto Mior
'     reexre@gmail.com
'--------------------------------------------------------------------------------

'Joints Handler By Roberto Mior

Option Explicit

Private Type tJoint
    O1                 As Long
    P1                 As Long
    O2                 As Long
    P2                 As Long
End Type


Public NJ              As Long
Private J()            As tJoint



Public Sub ADDjoint(O1, O2, P1, P2)


    NJ = NJ + 1
    ReDim Preserve J(NJ)
    With J(NJ)
        .O1 = O1
        .P1 = P1
        .O2 = O2
        .P2 = P2

        B(.O1).IsNOTJoint(.P1) = False
        B(.O2).IsNOTJoint(.P2) = False

    End With

End Sub


Public Sub UpDateJoints()
    Dim I              As Long
    Dim X1             As Single
    Dim Y1             As Single
    Dim X2             As Single
    Dim Y2             As Single
    Dim Dx             As Single
    Dim Dy             As Single

    For I = 1 To NJ
        With J(I)
            X1 = B(.O1).getPointX(.P1)
            Y1 = B(.O1).getPointY(.P1)
            X2 = B(.O2).getPointX(.P2)
            Y2 = B(.O2).getPointY(.P2)

            'Dx = X2 - X1
            'Dy = Y2 - Y1
            'Normalize Dx, Dy

            X1 = (X1 + X2) * 0.5
            Y1 = (Y1 + Y2) * 0.5

            B(.O1).SetPointX(.P1) = X1    '- Dx
            B(.O1).SetPointY(.P1) = Y1    '- Dy
            B(.O2).SetPointX(.P2) = X1    '+ Dx
            B(.O2).SetPointY(.P2) = Y1    '+ Dy


        End With
    Next

End Sub

