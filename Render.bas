Attribute VB_Name = "Render"
Option Explicit

Public Function UpdateDevice(ByRef pic As PictureBox) As Boolean
    pic.PSet (-1, -1), 0
End Function

Public Function AntCreature(ByRef pic As PictureBox, _
                            ByRef Creature As Ant, _
                            Optional ByVal lngSize As Long = 2) As Boolean
    Dim M(1) As Double
    Dim curFont As Long
    Dim oldFont As Long
    
    M(0) = Creature.X - Offset(0)
    M(1) = Creature.Y - Offset(1)
    
    'Draw motion vector
    'pic.DrawStyle = 0
    'pic.Line (M(0), M(1))-(M(0) + Sin(Creature.Direction) * 20, M(1) + Cos(Creature.Direction) * 20), vbBlack
    
    'Draw ant body
    pic.FillStyle = 0
    pic.FillColor = 0
    pic.DrawStyle = 5
    pic.Circle (M(0), M(1)), lngSize, 0
    pic.Circle (M(0) + lngSize * 1.5 * Sin(Creature.Direction), M(1) + lngSize * 1.5 * Cos(Creature.Direction)), lngSize
    pic.Circle (M(0) - lngSize * 1.5 * Sin(Creature.Direction), M(1) - lngSize * 1.5 * Cos(Creature.Direction)), lngSize
    
    'Draw cargo
    If Creature.Cargo > 0 Then
        pic.FillStyle = 0
        pic.FillColor = vbWhite
        pic.DrawStyle = 0
        pic.Circle (M(0) + lngSize * 3 * Sin(Creature.Direction), _
                    M(1) + lngSize * 3 * Cos(Creature.Direction)), 6
        curFont = CreateFont(12, 0, 0, 0, 400, False, False, False, 1, 7, 0, 2, 0, "MS Sans Serif")
        oldFont = SelectObject(pic.hdc, curFont)
        Call TextOut(pic.hdc, _
                     M(0) + lngSize * 3 * Sin(Creature.Direction) - 2, _
                     M(1) + lngSize * 3 * Cos(Creature.Direction) - 5, _
                     CStr(Creature.Cargo), Len(CStr(Creature.Cargo)))
        Call SelectObject(pic.hdc, oldFont)
        Call DeleteObject(curFont)
    End If
    
    AntCreature = True
End Function

Public Function GridElement(ByRef pic As PictureBox, _
                            ByVal X As Long, ByVal Y As Long, _
                            ByVal Width As Long, ByVal Height As Long, _
                            Optional ByVal EdgeColour As Long = -1, _
                            Optional ByVal FillColour As Long = -1) As Boolean
    GridElement = False
    Dim curPen As Long, oldPen As Long
    Dim curBrush As Long, oldBrush As Long
    
    If EdgeColour >= 0 Then curPen = CreatePen(0, 1, EdgeColour) Else curPen = CreatePen(5, 0, 0)
    If FillColour >= 0 Then curBrush = CreateSolidBrush(FillColour) Else curBrush = CreateSolidBrush(pic.BackColor)
    
    oldPen = SelectObject(pic.hdc, curPen)
    oldBrush = SelectObject(pic.hdc, curBrush)
    Call Rectangle(pic.hdc, X, Y, X + Width + 1, Y + Height + 1)
    Call SelectObject(pic.hdc, oldPen)
    Call SelectObject(pic.hdc, oldBrush)
    Call DeleteObject(curPen)
    Call DeleteObject(curBrush)

    GridElement = True
End Function

Public Function GridLabel(ByRef pic As PictureBox, _
                          ByVal X As Long, ByVal Y As Long, _
                          Optional ByVal LabelText As String = "", _
                          Optional ByVal LabelColour As Long = 0, _
                          Optional ByVal LabelSize As Long = 10) As Boolean
    GridLabel = False
    Dim curFont As Long, oldFont As Long

    If LabelText <> "" Then
        curFont = CreateFont(LabelSize, 0, 0, 0, 400, _
                             False, False, False, 1, 7, 0, 2, 0, "MS Sans Serif")
        oldFont = SelectObject(pic.hdc, curFont)
        Call TextOut(pic.hdc, X + 2, Y + 2, LabelText, Len(LabelText))
        Call SelectObject(pic.hdc, oldFont)
        Call DeleteObject(curFont)
    End If
        
    GridLabel = True
End Function

Public Function Label(ByRef pic As PictureBox, _
                      ByVal Text As String, _
                      ByVal X As Long, Y As Long, _
                      Optional ByVal TextSize As Long = 10, _
                      Optional ByVal blnBold As Boolean = False, _
                      Optional ByVal blnItalic As Boolean = False, _
                      Optional ByVal BackColour As Long = -1) As Boolean
    Label = False
    Dim TextExtend As PointAPI
    Dim curFont As Long, oldFont As Long
    Dim charWidth As Long
    
    If blnBold Then charWidth = 700 Else charWidth = 400
    curFont = CreateFont(TextSize, 0, 0, 0, charWidth, blnItalic, False, False, 1, 7, 0, 2, 0, "MS Sans Serif")
    oldFont = SelectObject(pic.hdc, curFont)
    Call GetTextExtentPoint32(pic.hdc, Text, Len(Text), TextExtend)
    
    If BackColour >= 0 Then
        Call GridElement(pic, X - 2, Y - 2, TextExtend.X + 4, TextExtend.Y + 2, 0, BackColour)
    End If
    Call TextOut(pic.hdc, X, Y, Text, Len(Text))
    Call SelectObject(pic.hdc, oldFont)
    Call DeleteObject(curFont)
End Function

Public Function Region(ByVal dblValue, ByVal dblMin As Double, ByVal dblMax As Double)
    Region = dblValue
    If dblValue < dblMin Then Region = dblMin
    If dblValue > dblMax Then Region = dblMax
End Function
