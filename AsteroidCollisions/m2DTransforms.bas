Attribute VB_Name = "m2DTransforms"
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const mc_strModuleName As String = "m2DTransforms"

Public Function MatrixIdentity() As mdrMATRIX3x3
    
    With MatrixIdentity
    
        .rc11 = 1: .rc12 = 0: .rc13 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1
    
    End With
    
End Function

Public Function MatrixMultiplyVector(MatrixIn As mdrMATRIX3x3, VectorIn As mdrVector3) As mdrVector3
        
    With MatrixMultiplyVector
    
        .x = (MatrixIn.rc11 * VectorIn.x) + (MatrixIn.rc12 * VectorIn.y) + (MatrixIn.rc13 * VectorIn.w)
        .y = (MatrixIn.rc21 * VectorIn.x) + (MatrixIn.rc22 * VectorIn.y) + (MatrixIn.rc23 * VectorIn.w)
        .w = 1
        
    End With
    
End Function

Public Function Vect3Subtract(V1 As mdrVector3, V2 As mdrVector3) As mdrVector3

    ' Subtracts V2 from V1
    
    Vect3Subtract.x = V1.x - V2.x
    Vect3Subtract.y = V1.y - V2.y
    
    ' We can safely ignore the W component.
    Vect3Subtract.w = 1
    
End Function
Public Function Vect3Addition(V1 As mdrVector3, V2 As mdrVector3) As mdrVector3

    ' Adds V1 and V2 together.
    
    Vect3Addition.x = V1.x + V2.x
    Vect3Addition.y = V1.y + V2.y
    
    ' We can safely ignore the W component.
    Vect3Addition.w = 1
    
End Function

Public Function MatrixMultiply(m1 As mdrMATRIX3x3, m2 As mdrMATRIX3x3) As mdrMATRIX3x3
        
    MatrixMultiply = MatrixIdentity
        
    With MatrixMultiply
        
        .rc11 = (m1.rc11 * m2.rc11) + (m1.rc21 * m2.rc12) + (m1.rc31 * m2.rc13)
        .rc12 = (m1.rc12 * m2.rc11) + (m1.rc22 * m2.rc12) + (m1.rc32 * m2.rc13)
        .rc13 = (m1.rc13 * m2.rc11) + (m1.rc23 * m2.rc12) + (m1.rc33 * m2.rc13)
        
        .rc21 = (m1.rc11 * m2.rc21) + (m1.rc21 * m2.rc22) + (m1.rc31 * m2.rc23)
        .rc22 = (m1.rc12 * m2.rc21) + (m1.rc22 * m2.rc22) + (m1.rc32 * m2.rc23)
        .rc23 = (m1.rc13 * m2.rc21) + (m1.rc23 * m2.rc22) + (m1.rc33 * m2.rc23)
        
        .rc31 = 0 '(m1.rc11 * m2.rc31) + (m1.rc21 * m2.rc32) + (m1.rc31 * m2.rc33)
        .rc32 = 0 '(m1.rc12 * m2.rc31) + (m1.rc22 * m2.rc32) + (m1.rc32 * m2.rc33)
        .rc33 = 1 '(m1.rc13 * m2.rc31) + (m1.rc23 * m2.rc32) + (m1.rc33 * m2.rc33)
    
    End With
    
    ' For this version of the application, row 3 of the matrix will not change.
    
End Function

Public Function Vec3Length(V1 As mdrVector3) As Single

    ' Returns the length of a 3-D vector.
    ' The length of a vector is from the origin (0,0) to x,y
    ' We work this out using Pythagoras theorem:  c^2 = a^2 + b^2
    
    Vec3Length = Sqr((V1.x ^ 2) + (V1.y ^ 2))
    
    ' We can safely ignore the W component.
    
End Function
Public Function Vec3MultiplyByScalar(V1 As mdrVector3, Scalar As Single) As mdrVector3
    
    Vec3MultiplyByScalar.x = V1.x * Scalar
    Vec3MultiplyByScalar.y = V1.y * Scalar
    
    ' We can safely ignore the W component.
    Vec3MultiplyByScalar.w = 1
    
End Function
Public Function Vec3Normalize(V1 As mdrVector3) As mdrVector3

    ' Returns the normalized version of a 3D vector.
    '
    ' When you divide a vector by it's own length (from origin 0,0 to x,y)
    ' you'll get a vector who's length is exactly 1.0
    
    Dim sngLength As Single
    
    sngLength = Vec3Length(V1)
    
    If sngLength = 0 Then sngLength = 1
    
    Vec3Normalize.x = V1.x / sngLength
    Vec3Normalize.y = V1.y / sngLength
    
    ' We can safely ignore the W component.
    Vec3Normalize.w = 1
    
End Function
Public Function DotProduct(VectorU As mdrVector3, VectorV As mdrVector3) As Single

    ' Determines the dot-product of two vectors.
    DotProduct = (VectorU.x * VectorV.x) + (VectorU.y * VectorV.y)
    
End Function

Public Function MatrixTranslation(OffsetX As Single, OffsetY As Single) As mdrMATRIX3x3
    
    MatrixTranslation = MatrixIdentity
    
    MatrixTranslation.rc13 = OffsetX
    MatrixTranslation.rc23 = OffsetY
    
End Function


Public Function MatrixScaling(ScaleX As Single, ScaleY As Single) As mdrMATRIX3x3
    
    MatrixScaling = MatrixIdentity
    
    MatrixScaling.rc11 = ScaleX
    MatrixScaling.rc22 = ScaleY
    
End Function

Public Function MatrixRotationZ(Radians As Single) As mdrMATRIX3x3

    ' ===========================================================================================
    '               *** This application uses a Right-Handed Coordinate System ***
    ' ===========================================================================================
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points out of the monitor towards you.
    '
    ' Note: DirectX uses a Left-Handed Coordinate system, which many people find more intuitive.
    ' This coordinate system is much closer to OpenGL.
    ' ===========================================================================================
    
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 6)
    sngSine = Round(Sin(Radians), 6)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationZ = MatrixIdentity()

    ' =======================================================================================================
    ' Positive rotations in a right-handed coordinate system are such that, when looking from a
    ' positive axis back towards the origin (0,0,0), a 90° "counter-clockwise" rotation will
    ' transform one positive axis into the other:
    '
    ' Z-Axis rotation.
    ' A positive rotation of 90° transforms the +X axis into the +Y axis.
    ' An additional positive rotation of 90° transforms the +Y axis into the -X axis.
    ' An additional positive rotation of 90° transforms the -X axis into the -Y axis.
    ' An additional positive rotation of 90° transforms the -Y axis into the +X axis (back where we started).
    ' =======================================================================================================
    With MatrixRotationZ
        .rc11 = sngCosine
        .rc21 = sngSine
        .rc12 = -sngSine
        .rc22 = sngCosine
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Sines and Cosines in different positions (like the columns
    ' and rows have been swapped over (ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.

End Function




Public Function MatrixViewMapping(WindowOfWorld As mdrWindow, ViewPort As mdrWindow) As mdrMATRIX3x3
    
    Dim matTranslateA As mdrMATRIX3x3
    Dim matScale As mdrMATRIX3x3
    Dim matTranslateB As mdrMATRIX3x3
    Dim sngScaleUX As Single
    Dim sngScaleVY As Single
    
    matTranslateA = MatrixTranslation(-WindowOfWorld.xMin, -WindowOfWorld.yMin)
    
    sngScaleUX = (ViewPort.xMax - ViewPort.xMin) / (WindowOfWorld.xMax - WindowOfWorld.xMin)
    sngScaleVY = (ViewPort.yMax - ViewPort.yMin) / (WindowOfWorld.yMax - WindowOfWorld.yMin)
    
    matScale = MatrixScaling(sngScaleUX, sngScaleVY)
    
    matTranslateB = MatrixTranslation(ViewPort.xMin, ViewPort.yMin)
    
    MatrixViewMapping = MatrixIdentity
    MatrixViewMapping = MatrixMultiply(MatrixViewMapping, matTranslateA)
    MatrixViewMapping = MatrixMultiply(MatrixViewMapping, matScale)
    MatrixViewMapping = MatrixMultiply(MatrixViewMapping, matTranslateB)
    
End Function

