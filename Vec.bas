Attribute VB_Name = "Vec"
Option Explicit

Public Type Point
    X As Double
    Y As Double
End Type

Public Type Vector
    Head As Point
    Tail As Point
End Type

Public Function IsPointInRectangle(ByRef ptIn As Point, ByRef recUL As Point, ByRef recLR As Point) As Boolean
    If (ptIn.X >= recUL.X) And _
       (ptIn.X <= recLR.X) And _
       (ptIn.Y >= recUL.Y) And _
       (ptIn.Y <= recLR.Y) Then
        IsPointInRectangle = True
    Else
        IsPointInRectangle = False
    End If
End Function

Public Function CreateVector(ByRef ptTail As Point, ByRef ptHead As Point) As Vector
    CreateVector.Tail = ptTail
    CreateVector.Head = ptHead
End Function

Public Function CreatePoint(Optional ByVal X As Double = 0, Optional ByVal Y As Double = 0) As Point
    CreatePoint.X = X
    CreatePoint.Y = Y
End Function

Public Function VectorLength(ByRef vecIn As Vector) As Double
    VectorLength = (vecIn.Head.X - vecIn.Tail.X) ^ 2 + (vecIn.Head.Y - vecIn.Tail.Y) ^ 2
    VectorLength = Sqr(VectorLength)
End Function

Public Function SetVectorToOrigin(ByRef vecIn As Vector) As Vector
    SetVectorToOrigin.Head.X = vecIn.Head.X - vecIn.Tail.X
    SetVectorToOrigin.Head.Y = vecIn.Head.Y - vecIn.Tail.Y
    SetVectorToOrigin.Tail.X = 0
    SetVectorToOrigin.Tail.Y = 0
End Function

Public Function SetVectorToPoint(ByRef vecIn As Vector, ByRef ptOrigin As Point) As Vector
    SetVectorToPoint.Head.X = vecIn.Head.X + (ptOrigin.X - vecIn.Tail.X)
    SetVectorToPoint.Head.Y = vecIn.Head.Y + (ptOrigin.Y - vecIn.Tail.Y)
    SetVectorToPoint.Tail.X = ptOrigin.X
    SetVectorToPoint.Tail.Y = ptOrigin.Y
End Function

Public Function ResizeVector(ByRef vecIn As Vector, Optional ByVal dblNewLength As Double = 1) As Vector
    ResizeVector = vecIn
    Dim curLength As Double
    curLength = VectorLength(vecIn)
    If curLength = 0 Then Exit Function
    
    ResizeVector.Tail.X = vecIn.Tail.X
    ResizeVector.Tail.Y = vecIn.Tail.Y
    ResizeVector.Head.X = vecIn.Tail.X + ((vecIn.Head.X - vecIn.Tail.X) / curLength) * dblNewLength
    ResizeVector.Head.Y = vecIn.Tail.Y + ((vecIn.Head.Y - vecIn.Tail.Y) / curLength) * dblNewLength
End Function

Public Function ScaleVector(ByRef vecIn As Vector, Optional ByVal dblScaleFactor As Double = 2) As Vector
    ScaleVector = vecIn
    If dblScaleFactor = 1 Then Exit Function
    
    ScaleVector.Tail.X = vecIn.Tail.X
    ScaleVector.Tail.Y = vecIn.Tail.Y
    ScaleVector.Head.X = vecIn.Tail.X + (vecIn.Head.X - vecIn.Tail.X) * dblScaleFactor
    ScaleVector.Head.Y = vecIn.Tail.Y + (vecIn.Head.Y - vecIn.Tail.Y) * dblScaleFactor
End Function

Public Function SumVector(ByRef arrVectors() As Vector, _
                          ByRef ptOrigin As Point) As Vector
    Dim i As Long
    Dim vecSum As Vector
    Dim arrVec() As Vector
    
    ReDim arrVec(UBound(arrVectors))
    For i = 0 To UBound(arrVectors)
        arrVec(i) = SetVectorToPoint(arrVectors(i), CreatePoint)
    Next
    
    vecSum = arrVec(0)
    If UBound(arrVectors) = 0 Then SumVector = vecSum: Exit Function
    For i = 1 To UBound(arrVec)
        vecSum.Head.X = vecSum.Head.X + arrVec(i).Head.X
        vecSum.Head.Y = vecSum.Head.Y + arrVec(i).Head.Y
    Next
    
    vecSum = SetVectorToPoint(vecSum, ptOrigin)
    SumVector = vecSum
End Function

