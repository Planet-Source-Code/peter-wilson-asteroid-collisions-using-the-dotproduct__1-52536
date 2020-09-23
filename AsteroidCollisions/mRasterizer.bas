Attribute VB_Name = "mRasterizer"
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const mc_strModuleName As String = "mRasterizer"

' The Polygon function draws a polygon consisting of two or more vertices connected by straight lines. The polygon is outlined by using the current pen and filled by using the current brush and polygon fill mode.
Type POINT_TYPE
  x As Long
  y As Long
End Type
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINT_TYPE, ByVal nCount As Long) As Long

Public Sub Draw_Vertices2(CurrentObject() As mdr2DObject, CurrentPictureBox As PictureBox)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    On Error GoTo errTrap
    
    CurrentPictureBox.DrawStyle = g_lngDrawStyle
    CurrentPictureBox.DrawMode = vbCopyPen
    CurrentPictureBox.DrawWidth = 1
    
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
            
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                
                ' Set colour of Object
                CurrentPictureBox.ForeColor = RGB(.Red, .Green, .Blue)
                
                ' Loop through the Vertices
                For intVertexIndex = LBound(.Vertex) To UBound(.Vertex)
                    xPos = .WorldPos.x
                    yPos = -.WorldPos.y
                    CurrentPictureBox.PSet (xPos, yPos)
                Next intVertexIndex
                
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub

Public Sub Draw_Faces(CurrentObject() As mdr2DObject, CurrentForm As Form)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    On Error GoTo errTrap
    
    CurrentForm.DrawStyle = g_lngDrawStyle
    CurrentForm.DrawMode = vbCopyPen
    CurrentForm.DrawWidth = 1
    
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
            
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                
                ' Set colour of Object
                CurrentForm.ForeColor = RGB(.Red, .Green, .Blue)
                
                If .Caption = "DefaultPlayerAmmo" Then
                    xPos = .TVertex(0).x
                    yPos = .TVertex(0).y
                    CurrentForm.PSet (xPos, yPos)
                Else
                    For intFaceIndex = LBound(.Face) To UBound(.Face)
                        
                        For intK = LBound(.Face(intFaceIndex)) To UBound(.Face(intFaceIndex))
                        
                            intVertexIndex = .Face(intFaceIndex)(intK)
                            xPos = .TVertex(intVertexIndex).x
                            yPos = .TVertex(intVertexIndex).y
                            
                            If LBound(.Face(intFaceIndex)) = UBound(.Face(intFaceIndex)) Then
    
                            Else
                            
                                ' Normal Face; move to first point, then draw to the others.
                                ' ==========================================================
                                If intK = LBound(.Face(intFaceIndex)) Then
                                    ' Move to first point
                                    CurrentForm.Line (xPos, yPos)-(xPos, yPos)
'                                    Call mdrLine(xPos, yPos, xPos, yPos)
                                Else
                                    ' Draw to point
                                    CurrentForm.Line -(xPos, yPos)
'                                    Call mdrLine(frmCanvas.CurrentX, frmCanvas.CurrentY, xPos, yPos)
                                End If
                                
                            End If
                            
                        Next intK
                    Next intFaceIndex
                    
'                    CurrentForm.Print .CurrentThreatLevel

                End If ' Is DefaultPlayerAmmo?
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub
Public Sub Draw_Faces2(CurrentObject() As mdr2DObject, CurrentForm As Form)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    Dim sx As Single
    Dim sy As Single
    Dim lngReturnResult  As Long
    Dim FaceVertices() As POINT_TYPE
    
    On Error GoTo errTrap
    
    CurrentForm.DrawStyle = g_lngDrawStyle
    CurrentForm.DrawMode = vbCopyPen
    CurrentForm.DrawWidth = 1
    
    ' Set colour of Object
    CurrentForm.FillStyle = vbFSSolid
    CurrentForm.FillColor = RGB(0, 0, 0)
    
    ' Convert: Custom-Coordinates >> Twips >> Pixels
    sx = CurrentForm.Width / CurrentForm.ScaleWidth / Screen.TwipsPerPixelX
    sy = CurrentForm.Height / CurrentForm.ScaleHeight / Screen.TwipsPerPixelY
        
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
                        
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                                
                For intFaceIndex = LBound(.Face) To UBound(.Face)
                    ReDim FaceVertices(UBound(.Face(intFaceIndex)) + 1) As POINT_TYPE
                    ReDim FaceShadow(UBound(.Face(intFaceIndex)) + 1) As POINT_TYPE
                    
                    For intK = LBound(.Face(intFaceIndex)) To UBound(.Face(intFaceIndex))
                        intVertexIndex = .Face(intFaceIndex)(intK)
                        FaceVertices(intK).x = .TVertex(intVertexIndex).x * sx - (CurrentForm.ScaleLeft * sx)
                        FaceVertices(intK).y = .TVertex(intVertexIndex).y * sy - (CurrentForm.ScaleLeft * sy)
                    Next intK
                    
                    CurrentForm.ForeColor = RGB(.Red, .Green, .Blue)
                    lngReturnResult = Polygon(CurrentForm.hdc, FaceVertices(0), UBound(FaceVertices))
                
'                    If .Caption = "Asteroid" And (intFaceIndex = 1) Then
'                        CurrentForm.CurrentX = .TVertex(intVertexIndex).x
'                        CurrentForm.CurrentY = .TVertex(intVertexIndex).y
'                        CurrentForm.Print "Min: " & Format(.MinSize, "0") & ", Max: " & Format(.MaxSize, "0")
'                    End If
                
                Next intFaceIndex
            End If ' Is Enabled?
            
        End With
    Next intN

    Exit Sub
errTrap:

End Sub

Public Sub Draw_Faces3(CurrentObject() As mdr2DObject, CurrentForm As Form)

    Dim intN As Integer
    Dim intFaceIndex As Integer
    Dim intK As Integer
    Dim intVertexIndex As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim matTranslate As mdrMATRIX3x3
    Dim tempXPos As Single
    Dim tempYPos As Single
    
    Dim sx As Single
    Dim sy As Single
    Dim lngReturnResult  As Long
    Dim FaceVertices() As POINT_TYPE
    
    On Error GoTo errTrap
    
    CurrentForm.DrawStyle = g_lngDrawStyle
    CurrentForm.DrawMode = vbCopyPen
    CurrentForm.DrawWidth = 1
    
    ' Set colour of Object
    CurrentForm.FillStyle = vbFSSolid
    CurrentForm.FillColor = RGB(0, 0, 0)
        
    For intN = LBound(CurrentObject) To UBound(CurrentObject)
        With CurrentObject(intN)
            
            If .Enabled = True Then
                        
                ' Clamp values to safe levels
                If .Red < 0 Then .Red = 0
                If .Green < 0 Then .Green = 0
                If .Blue < 0 Then .Blue = 0
                If .Red > 255 Then .Red = 255
                If .Green > 255 Then .Green = 255
                If .Blue > 255 Then .Blue = 255
                                
                For intFaceIndex = LBound(.Face) To UBound(.Face)
                    ReDim FaceVertices(UBound(.Face(intFaceIndex)) + 1) As POINT_TYPE
                    ReDim FaceShadow(UBound(.Face(intFaceIndex)) + 1) As POINT_TYPE
                    
                    For intK = LBound(.Face(intFaceIndex)) To UBound(.Face(intFaceIndex))
                        intVertexIndex = .Face(intFaceIndex)(intK)
                        FaceVertices(intK).x = .TVertex(intVertexIndex).x
                        FaceVertices(intK).y = .TVertex(intVertexIndex).y
                    Next intK
                    
                    CurrentForm.ForeColor = RGB(.Red, .Green, .Blue)
                    CurrentForm.FillColor = RGB(.Red / 3, .Green / 3, .Blue / 3)
                    lngReturnResult = Polygon(CurrentForm.hdc, FaceVertices(0), UBound(FaceVertices))
'
'                    If .Caption = "Asteroid" And (intFaceIndex = 1) Then
'                        CurrentForm.CurrentX = .TVertex(intVertexIndex).X
'                        CurrentForm.CurrentY = .TVertex(intVertexIndex).y
'                        If .TempAngle <> 0 Then CurrentForm.Print .TempAngle
''                        CurrentForm.Print "Min: " & Format(.MinSize, "0") & ", Max: " & Format(.MaxSize, "0")
'                    End If
'
                Next intFaceIndex
            End If ' Is Enabled?
            
        End With
    Next intN

    Exit Sub
errTrap:

End Sub


