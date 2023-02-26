Attribute VB_Name = "OBB"
'   ____  ____  ____     ___    __                 _ __  __
'  / __ \/ __ )/ __ )   /   |  / /___ _____  _____(_) /_/ /_  ____ ___
' / / / / __  / __  |  / /| | / / __ `/ __ \/ ___/ / __/ __ \/ __ `__ \
'/ /_/ / /_/ / /_/ /  / ___ |/ / /_/ / /_/ / /  / / /_/ / / / / / / / /
'\____/_____/_____/  /_/  |_/_/\__, /\____/_/  /_/\__/_/ /_/_/ /_/ /_/
'                             /____/
' --> By arfu
' --> Version 1.2.1

Option Explicit

' --> Math Functions

Private Function Min4(a As Double, b As Double, c As Double, d As Double) As Double
    Min4 = a
    If b < Min4 Then Min4 = b
    If c < Min4 Then Min4 = c
    If d < Min4 Then Min4 = d
End Function

Private Function Max4(a As Double, b As Double, c As Double, d As Double) As Double
    Max4 = a
    If b > Max4 Then Max4 = b
    If c > Max4 Then Max4 = c
    If d > Max4 Then Max4 = d
End Function

' --> Private Functions
Private Function GetVectors(ShapeA As Shape, ShapeB As Shape)
    Const Pi As Double = 3.14159265358979
    Dim Arot As Double, Brot As Double, VectorA As Variant, VectorB As Variant
    
    Arot = ShapeA.Rotation * Pi / 180
    Brot = ShapeB.Rotation * Pi / 180
    
    VectorA = Array(Array(Cos(Arot), -Sin(Arot)), Array(Sin(Arot), Cos(Arot)))
    VectorB = Array(Array(Cos(Brot), -Sin(Brot)), Array(Sin(Brot), Cos(Brot)))
    
    GetVectors = Array(VectorA, VectorB)
End Function

Private Function GetVertices(Shape As Shape, vector As Variant)
    Dim hx As Double, hy As Double, v As Variant, vp(0 To 3), i As Byte
    hx = Shape.Width / 2
    hy = Shape.Height / 2
    
    v = Array(Array(-hx, -hy), Array(hx, -hy), Array(hx, hy), Array(-hx, hy))
    
    For i = 0 To 3
        vp(i) = Array((Shape.Left + hx) + vector(0)(0) * v(i)(0) + vector(0)(1) * v(i)(1), (Shape.Top + hy) + vector(1)(0) * v(i)(0) + vector(1)(1) * v(i)(1))
    Next
    
    GetVertices = vp
End Function

' --> Collision
Public Function OBBIntersect(ShapeA As Shape, ShapeB As Shape) As Boolean
    Dim Vectors As Variant, i As Byte, j As Byte, Axis(0 To 1) As Double, VerticesA As Variant, VerticesB As Variant
    Vectors = GetVectors(ShapeA, ShapeB)
    VerticesA = GetVertices(ShapeA, Vectors(0))
    VerticesB = GetVertices(ShapeB, Vectors(1))
    
    For i = 0 To 1
        For j = 0 To 1
            Axis(0) = Vectors(i)(0)(j)
            Axis(1) = Vectors(i)(1)(j)
            
            Dim minA As Double, maxA As Double, minB As Double, maxB As Double
            minA = Min4(VerticesA(0)(0) * Axis(0) + VerticesA(0)(1) * Axis(1), _
                        VerticesA(1)(0) * Axis(0) + VerticesA(1)(1) * Axis(1), _
                        VerticesA(2)(0) * Axis(0) + VerticesA(2)(1) * Axis(1), _
                        VerticesA(3)(0) * Axis(0) + VerticesA(3)(1) * Axis(1))
            maxA = Max4(VerticesA(0)(0) * Axis(0) + VerticesA(0)(1) * Axis(1), _
                        VerticesA(1)(0) * Axis(0) + VerticesA(1)(1) * Axis(1), _
                        VerticesA(2)(0) * Axis(0) + VerticesA(2)(1) * Axis(1), _
                        VerticesA(3)(0) * Axis(0) + VerticesA(3)(1) * Axis(1))
            minB = Min4(VerticesB(0)(0) * Axis(0) + VerticesB(0)(1) * Axis(1), _
                        VerticesB(1)(0) * Axis(0) + VerticesB(1)(1) * Axis(1), _
                        VerticesB(2)(0) * Axis(0) + VerticesB(2)(1) * Axis(1), _
                        VerticesB(3)(0) * Axis(0) + VerticesB(3)(1) * Axis(1))
            maxB = Max4(VerticesB(0)(0) * Axis(0) + VerticesB(0)(1) * Axis(1), _
                        VerticesB(1)(0) * Axis(0) + VerticesB(1)(1) * Axis(1), _
                        VerticesB(2)(0) * Axis(0) + VerticesB(2)(1) * Axis(1), _
                        VerticesB(3)(0) * Axis(0) + VerticesB(3)(1) * Axis(1))
            
            If maxA < minB Or minA > maxB Or ShapeB.Visible = msoFalse Then
                OBBIntersect = False
                Exit Function
            End If
        Next
    Next
    
    OBBIntersect = True
End Function

