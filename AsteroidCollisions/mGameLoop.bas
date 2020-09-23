Attribute VB_Name = "mGameLoop"
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const mc_strModuleName As String = "mGameLoop"


Public g_strGameState As String
Private m_lngApplicationState As Long

Public g_intCurrentLevel As Integer

Private m_MaxAsteroids As Integer
Public m_Asteroids() As mdr2DObject


Public m_MaxParticles As Integer
Public m_Particles() As mdr2DObject

' Game World Window Limits ie. This is the game's world coordinates (which could be very large)
Private m_GameWorld As mdrWindow

' We will view the Game's world through the following window/s.
Public m_Window As mdrWindow

' Whatever we can see through the window, will be displayed on the viewport/s.
Private m_ViewPort As mdrWindow

Private m_Xmin As Single
Private m_Xmax As Single
Private m_Ymin As Single
Private m_Ymax As Single



' ViewPort Limits ie. Usually the limits of a VB form, or picturebox (which could be very small)
' Note: Just because your game's world is large, does not mean you need to display the whole world at
'       once. You can easily zoom in on some action just by changing the ViewPort values below.
'       Remember, the viewport is what you are looking at.
Private m_Umin As Single
Private m_Umax As Single
Private m_Vmin As Single
Private m_Vmax As Single


' Module Level Matrices (that don't change much)
Public m_matScale As mdrMATRIX3x3
Public g_lngDrawStyle As DrawStyleConstants
Public g_matViewMapping As mdrMATRIX3x3
Public g_matViewMapping2 As mdrMATRIX3x3

Public g_blnDontClearScreen As Boolean


Public Function Create_Particles(Caption As String, NumberOfParticles As Integer, MinSize As Single, MaxSize As Single, WorldX As Single, WorldY As Single, VectorX As Single, VectorY As Single, Red As Integer, Green As Integer, Blue As Integer, LifeTime As Single, ZRotation As Single) As Integer

    ' "Attempts to create" the specified number of particles
    Create_Particles = 0
    
    Dim intN As Integer
    Dim sngRadius As Single
    
    intN = 0
    Do
        If m_Particles(intN).Enabled = False Then
            ' This particle is no longer used, so we can use this one --> m_Particles(intN)
            
            If NumberOfParticles > 0 Then
                
                Select Case Caption
                    Case "Asteroid"
                        ' Create a random sized asteroid within the min/max parameters specified.
                        sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                        m_Particles(intN) = CreateRandomShapeAsteroid(sngRadius)
                        
                    Case Else
                        ' Create a random sized asteroid within the min/max parameters specified.
                        sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                        m_Particles(intN) = CreateRandomShapeAsteroid(sngRadius)
                        
                End Select
                
                ' Fill-in some properties.
                With m_Particles(intN)
                    .Enabled = True
                    .Caption = Caption
                    .ParticleLifeRemaining = LifeTime
                    .WorldPos.x = WorldX
                    .WorldPos.y = WorldY
                    .WorldPos.w = 1
                    
                    ' Initial Vector
                    If VectorX = 0 Then
                        .Vector.x = GetRNDNumberBetween(-2, 2)
                    Else
                        .Vector.x = VectorX
                    End If
                    If VectorY = 0 Then
                        .Vector.y = GetRNDNumberBetween(-2, 2)
                    Else
                        .Vector.y = VectorY
                    End If
                    .Vector.w = 1
                        
                    If ZRotation = 0 Then .SpinVector = GetRNDNumberBetween(-4, 4)
                
                    .RotationAboutZ = ZRotation
                    .Red = Red: .Green = Green: .Blue = Blue
                End With
                
                NumberOfParticles = NumberOfParticles - 1
            End If
        End If
        
        intN = intN + 1
    Loop Until (intN = m_MaxParticles) Or (NumberOfParticles = 0)
        
End Function
Public Function Create_Asteroids(ByVal Qty As Integer, MinSize As Integer, MaxSize As Integer, WorldX As Single, WorldY As Single, Red As Integer, Green As Integer, Blue As Integer, LifeTime As Single) As Integer

    ' "Attempts to create" the specified number of Asteroids,
    ' and returns the number of Asteroids "actually created".
    Create_Asteroids = 0
    
    Dim intN As Integer
    Dim sngRadius As Single
    
    intN = 0
    Do
        If m_Asteroids(intN).Enabled = False Then
            ' This Asteroid is no longer used, so we can use this one --> m_Asteroids(intN)
            
            If Qty > 0 Then
                
                ' Create a random sized asteroid within the min/max parameters specified.
                sngRadius = GetRNDNumberBetween(MinSize, MaxSize)
                m_Asteroids(intN) = CreateRandomShapeAsteroid(sngRadius)
                
                ' Reset the properties of this Asteroid.
                With m_Asteroids(intN)
                    .Enabled = True
                    .Caption = "Asteroid"
                    .Health = 100
                    .ParticleLifeRemaining = LifeTime
                    
                    ' Set a random starting position
                    Dim sngRadians As Single
                    sngRadians = ConvertDeg2Rad(GetRNDNumberBetween(0, 359))
                    .WorldPos.x = Cos(sngRadians) * (m_GameWorld.xMax * 2)
                    .WorldPos.y = Sin(sngRadians) * (m_GameWorld.yMax * 2)
                    .WorldPos.w = 1
                    
                    ' Initial Vector (Direction and Magnitude) depending on the size of the Asteroid.
                    Dim sngTemp As Single
                    sngTemp = (30 / sngRadius)
                    
                    .Vector.x = GetRNDNumberBetween(-sngTemp, sngTemp)
                    .Vector.y = GetRNDNumberBetween(-sngTemp, sngTemp)
                    .Vector.w = 1
                    
                    .SpinVector = GetRNDNumberBetween(-2, 2)
                    .RotationAboutZ = 0
                    
                    .Red = Red: .Green = Green: .Blue = Blue
                    .Green = 0
                    .Blue = 0
                    .Red = 255 - (.MaxSize * 3)
                End With
                
                Qty = Qty - 1
            End If
        End If
        
        intN = intN + 1
    Loop Until (intN = m_MaxParticles) Or (Qty = 0)
        
End Function

Public Sub zCreate_Asteroids2()


    ' ===================================================
    ' Create "test" Asteroids for testing collision code.
    ' ===================================================
    ReDim m_Asteroids(0)
    
    m_Asteroids(0) = CreateRandomShapeAsteroid(10)
    With m_Asteroids(0)
        .WorldPos.x = 0
        .WorldPos.y = 0
        .Vector.x = GetRNDNumberBetween(-10, 10)
        .Vector.y = GetRNDNumberBetween(-10, 10)
        .Enabled = True
        .Caption = "Asteroid"
        .Green = 255
        .Health = 1000
    End With

    Dim sngX As Single
    Dim sngY As Single
    Dim intAsteroidCount As Integer
    
    For sngX = -m_GameWorld.xMax To m_GameWorld.xMax Step (m_GameWorld.xMax / 4)
        For sngY = -m_GameWorld.yMax To m_GameWorld.yMax Step (m_GameWorld.yMax / 4)
        
            If ((sngX = 0) And (sngY = 0)) Or ((Abs(sngX) = Abs(m_GameWorld.xMax)) Or (Abs(sngY) = Abs(m_GameWorld.yMax))) Then
               ' do nothing
            Else
                intAsteroidCount = intAsteroidCount + 1
                ReDim Preserve m_Asteroids(intAsteroidCount)
                
                m_Asteroids(intAsteroidCount) = CreateRandomShapeAsteroid(GetRNDNumberBetween(5, 35))
                With m_Asteroids(intAsteroidCount)
                    .WorldPos.x = sngX
                    .WorldPos.y = sngY
                    .Vector.x = 0
                    .Vector.y = 0
                    .Enabled = True
                    .Caption = "Asteroid"
                    .Red = 255
                    .Health = 100
                    .SpinVector = GetRNDNumberBetween(-2, 2)
                End With
            End If
            
        Next sngY
    Next sngX
    
End Sub




Private Sub Init_Game()
        
    ' ================================================================
    ' Hide Mouse (by moving it to the far bottom right)
    ' This method causes less problems than actually hiding the mouse,
    ' although moving the mouse can confuse the user, so careful.
    ' ================================================================
    Call SetCursorPos(Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)

    ' ====================================================
    ' Show the Form's needed for drawing the game (Canvas)
    ' ====================================================
    frmCanvas.Show
    g_intCurrentLevel = 0 ' If you want the User to start on Level 5, put 4 here.
    
    ' Init Particles
    ' ==============
    m_MaxParticles = 64
    ReDim m_Particles(m_MaxParticles - 1)
    
    g_strGameState = "LevelComplete"
    
End Sub

Public Sub Main()

    ' ==========================================================================
    ' This routine get's called by a Timer Event regardless of what's happening.
    ' (Although you can have multiple Timer Controls, it tends to make programs
    '  disorganised and less predictable. By using only a single Timer control,
    '  I have very strict control over what occurs and when. This routine is
    '  actually a mini-"state machine"... well actually most computer programs
    '  are, but I digress... look them up, learn them, they are cool.)
    ' ==========================================================================
        
    Select Case g_strGameState
        Case ""
            Call Init_Game
            
        Case "PlayingLevel"
            Call PlayGame
            
        Case "LevelComplete"
            ' User has finished a level. Increment and reset game data for the next level.
            ' ============================================================================
            g_intCurrentLevel = g_intCurrentLevel + 1
            Call LoadLevel(g_intCurrentLevel)
            g_strGameState = "PlayingLevel"
            
        Case "Quit"
            frmCanvas.Timer_DoAnimation.Enabled = False
            Unload frmCanvas
            
    End Select
    
    Call ProcessKeyboardInput
    
End Sub

Private Sub LoadLevel(ByVal Level As Integer)
    
    ' Initializes the random-number generator.
    Call Rnd(-1)
    Call Randomize(Level)

    ' Reset Global Scale
    
    m_matScale = MatrixScaling(1, 1)
    
    ' Create Asteroids
    Call zCreate_Asteroids2
    
End Sub

Public Sub PlayGame()
    
    ' Prepare objects for display.
    ' ============================
    Call Calculate_Asteroids_Collisions
    Call Calculate_Asteroids(g_matViewMapping2)
    Call Calculate_Particles(g_matViewMapping)
    
    ' Display results.
    ' ================
    Call Refresh_GameScreen(frmCanvas)
    
    
'''    ' ===============================================================
'''    ' Save Animation Frames to HDD (only if Left-Control key is down)
'''    ' This is good for creating images ready for a GIF animator.
'''    ' ===============================================================
'''    Dim intKeyState As Integer
'''    Static lngImageCount As Long
'''    Dim strFileName As String
'''
'''    intKeyState = GetKeyState(VK_LCONTROL)
'''    If intKeyState And &H8000 Then
'''        lngImageCount = lngImageCount + 1
'''        strFileName = "d:\Snapshot" & Format(lngImageCount, "000") & ".bmp"
'''        Call SavePicture(frmCanvas.Image, strFileName)
'''    End If
    
    
    
End Sub

Public Sub Calculate_Asteroids(ViewMapping As mdrMATRIX3x3)

    Dim intN As Integer
    Dim intJ As Integer
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
    Dim sngLength As Single
    
    On Error GoTo errTrap
    
    
    For intN = LBound(m_Asteroids) To UBound(m_Asteroids)
        With m_Asteroids(intN)
            If .Enabled = True Then
            
                If (.Vector.x = 0) And (.Vector.y = 0) And (.TempDrawAtLeastOnce = True) Then
                    ' Do Nothing
                    .Red = 128
                    .Green = 0
                    .Blue = 0
                    
                Else
                    
                    .TempDrawAtLeastOnce = True
                    
                    ' Apply the direction/magnitude vector to the world coordinates.
                    ' ==============================================================
                    .WorldPos.x = .WorldPos.x + .Vector.x
                    .WorldPos.y = .WorldPos.y + .Vector.y
                    
                    
'''                    ' =============================================================
'''                    ' Set Colour depending on Vector length (optional... of course)
'''                    ' =============================================================
                    sngLength = Vec3Length(.Vector)
                    If sngLength > 200 Then g_strGameState = "LevelComplete"
'''                    .Red = 255 - (.Health * 2.55)
'''                    .Green = 64 * sngLength
'''                    .Blue = .Health * 2.55
                    
                    .Red = 0
                    .Green = 255
                    .Blue = 0
                    
                    ' Clamp world position values to the game's world coordinate system
                    ' (ie. Wrap aseteroids around the game's world)
                    If .WorldPos.x > m_GameWorld.xMax Then .WorldPos.x = .WorldPos.x - (m_GameWorld.xMax - m_GameWorld.xMin)
                    If .WorldPos.x < m_GameWorld.xMin Then .WorldPos.x = .WorldPos.x + (m_GameWorld.xMax - m_GameWorld.xMin)
                    If .WorldPos.y > m_GameWorld.yMax Then .WorldPos.y = .WorldPos.y - (m_GameWorld.yMax - m_GameWorld.yMin)
                    If .WorldPos.y < m_GameWorld.yMin Then .WorldPos.y = .WorldPos.y + (m_GameWorld.yMax - m_GameWorld.yMin)
                    
                    
                    ' ===========================
                    ' Setup a Translation matrix.
                    ' ===========================
                    matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
                    
                    ' =======================================================
                    ' Apply the spin vector to the Asteroid's rotation value.
                    ' =======================================================
                    .RotationAboutZ = .RotationAboutZ + .SpinVector
                    
                    ' ========================
                    ' Setup a Rotation matrix.
                    ' ========================
                    matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
                    
                    ' =======================================
                    ' Multiply matrices in the correct order.
                    ' =======================================
                    matResult = MatrixIdentity
                    matResult = MatrixMultiply(matResult, m_matScale)
                    matResult = MatrixMultiply(matResult, matRotationAboutZ)
                    matResult = MatrixMultiply(matResult, matTranslate)
                    matResult = MatrixMultiply(matResult, ViewMapping)
                    
                    ' =========================================
                    ' Apply the transformation to the vertices.
                    ' =========================================
                    For intJ = LBound(.Vertex) To UBound(.Vertex)
                        .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
                    Next intJ
                    
                End If
            End If ' Is Enabled?
        End With
    Next intN

    Exit Sub
errTrap:

End Sub


Public Sub Calculate_Particles(ViewMapping As mdrMATRIX3x3)

    ' Processes all Particles (Asteroids, Exhaust, Bullets, Smoke, Flames, Explosions, etc.)
    
    Dim intN As Integer
    Dim intJ As Integer
    Dim matCustomScale As mdrMATRIX3x3
    Dim matTranslate As mdrMATRIX3x3
    Dim matRotationAboutZ As mdrMATRIX3x3
    Dim sngAngleZ As Single
    Dim matResult As mdrMATRIX3x3
        
    For intN = LBound(m_Particles) To UBound(m_Particles)
        
        matCustomScale = MatrixScaling(1, 1)

        With m_Particles(intN)
            If .Enabled = True Then
                    
                Select Case .Caption
                    Case "Asteroid"
                        ' Fade to Dull Red, then to black.
                        .Red = .Red - 2
                        .Green = .Green - 4
                        .Blue = .Blue - 4
                        
                        ' Reduce Particle life
                        .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                        If .ParticleLifeRemaining < 1 Then .Enabled = False
                        
                    Case "Exhaust_Smoke"
                        .ParticleMisc1 = .ParticleMisc1 + 0.4
                        ' Smoke has 16 steps (not that it really matters)
                        If .ParticleMisc1 > 10 Then
                            .Red = .Red - 16
                            .Green = .Green - 16
                            .Blue = .Blue - 16
                        ElseIf .ParticleMisc1 > 3 Then
                            .Red = .Red + 24
                            .Green = .Green + 24
                            .Blue = .Blue + 24
                        End If
                        matCustomScale = MatrixScaling(.ParticleMisc1, .ParticleMisc1)
                        
                        .ParticleLifeRemaining = .ParticleLifeRemaining - 1
                        If .ParticleLifeRemaining < 1 Then .Enabled = False
                    
                End Select
                
                ' Translate
                ' =========
                .WorldPos.x = .WorldPos.x + .Vector.x
                .WorldPos.y = .WorldPos.y + .Vector.y
                
                If .WorldPos.x > m_GameWorld.xMax Then .WorldPos.x = .WorldPos.x - (m_GameWorld.xMax - m_GameWorld.xMin)
                If .WorldPos.x < m_GameWorld.xMin Then .WorldPos.x = .WorldPos.x + (m_GameWorld.xMax - m_GameWorld.xMin)
                If .WorldPos.y > m_GameWorld.yMax Then .WorldPos.y = .WorldPos.y - (m_GameWorld.yMax - m_GameWorld.yMin)
                If .WorldPos.y < m_GameWorld.yMin Then .WorldPos.y = .WorldPos.y + (m_GameWorld.yMax - m_GameWorld.yMin)
                
                matTranslate = MatrixTranslation(.WorldPos.x, .WorldPos.y)
                
                .RotationAboutZ = .RotationAboutZ + .SpinVector
                matRotationAboutZ = MatrixRotationZ(ConvertDeg2Rad(.RotationAboutZ))
    
                ' Multiply matrices in the correct order.
                matResult = MatrixIdentity
                matResult = MatrixMultiply(matResult, matCustomScale)
                matResult = MatrixMultiply(matResult, m_matScale)
                matResult = MatrixMultiply(matResult, matRotationAboutZ)
                matResult = MatrixMultiply(matResult, matTranslate)
                matResult = MatrixMultiply(matResult, ViewMapping)
                
                For intJ = LBound(.Vertex) To UBound(.Vertex)
                    .TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
                Next intJ
                
            End If ' Is Enabled?
            
        End With
    Next intN

End Sub

Public Sub Calculate_Asteroids_Collisions()

    Dim intAsteroid As Integer
    Dim intOtherAsteroid As Integer
    Dim tempV As mdrVector3
    Dim VDisplay As mdrVector3
    Dim tempV2 As mdrVector3
    Dim sngDistance As Single
    Dim sngMultiplier As Single
    Dim vectN As mdrVector3
    Dim vectJ As mdrVector3
    
    ' Loop through all Asteroids...
    For intAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
        With m_Asteroids(intAsteroid)
            If (.Enabled = True) Then
            
                ' ...and compare with every other asteroid.
                ' (this can be a slow process... n*(n-1) asteroids.)
                For intOtherAsteroid = LBound(m_Asteroids) To UBound(m_Asteroids)
                    If (m_Asteroids(intOtherAsteroid).Enabled = True) And (intOtherAsteroid <> intAsteroid) And (.TempVar1 = 0) Then
                    
                        ' =================
                        ' Collision Normal.
                        ' =================
                        vectN = Vect3Subtract(.WorldPos, m_Asteroids(intOtherAsteroid).WorldPos)
                        sngDistance = Vec3Length(vectN)
                        
                        If sngDistance < (.AvgSize + m_Asteroids(intOtherAsteroid).AvgSize) Then
                            ' Asteroids have collided.
                            
                            
                            ' Prevent Asteroids getting counted twice.
                            .TempVar1 = 1
                            m_Asteroids(intOtherAsteroid).TempVar1 = 1
                            
                            
                            ' ====================
                            ' Health of Asteroids.
                            ' ====================
                            .Health = .Health - 1
                            m_Asteroids(intOtherAsteroid).Health = m_Asteroids(intOtherAsteroid).Health - 10
                            If .Health < 0 Then
                                .Enabled = False
                                Call Create_Particles("Asteroid", 4, .AvgSize, .AvgSize, .WorldPos.x - (vectN.x / 2), .WorldPos.y - (vectN.y / 2), 0, 0, 255, 255, 255, 16, 0)
                                Call Create_Particles("Asteroid", 32, .AvgSize, .AvgSize, .WorldPos.x - (vectN.x / 2), .WorldPos.y - (vectN.y / 2), 0, 0, 255, 0, 0, 96, 0)
                            End If
                            
                            
                            ' ==============
                            ' Create Sparks.
                            ' ==============
                            Call Create_Particles("Asteroid", 4, .AvgSize / 3, .AvgSize / 3, .WorldPos.x - (vectN.x / 2), .WorldPos.y - (vectN.y / 2), 0, 0, 255, 255, 0, 32, 0)
                            
                            ' =================
                            ' Collision Normal.
                            ' =================
                            vectN = Vect3Subtract(.WorldPos, m_Asteroids(intOtherAsteroid).WorldPos)
                            vectN = Vec3Normalize(vectN)
                            vectJ.x = -vectN.x
                            vectJ.y = -vectN.y
                            
                            ' Find new velocity vectors
                            Dim Vn1 As mdrVector3
                            Dim Vn2 As mdrVector3
                            Dim sngDP As Single
                            
                            sngDP = DotProduct(.Vector, vectJ)
                            Vn1.x = sngDP * vectJ.x
                            Vn1.y = sngDP * vectJ.y
                            
                            sngDP = DotProduct(m_Asteroids(intOtherAsteroid).Vector, vectN)
                            Vn2.x = sngDP * vectN.x
                            Vn2.y = sngDP * vectN.y
                            
                            
                            ' ====================================
                            ' Find tangential velocity components.
                            ' ====================================
                            Dim Vt1 As mdrVector3
                            Dim Vt2 As mdrVector3
                            
                            Vt1 = Vect3Subtract(.Vector, Vn1)
                            Vt2 = Vect3Subtract(m_Asteroids(intOtherAsteroid).Vector, Vn2)
                            
                            ' Apply Velocities.
                            .Vector = Vect3Addition(Vt1, Vn2)
                            m_Asteroids(intOtherAsteroid).Vector = Vect3Addition(Vt2, Vn1)
                            
                            ' Reverse the Spin.
                            Dim sngSpin As Single
                            sngSpin = m_Asteroids(intOtherAsteroid).SpinVector
                            m_Asteroids(intOtherAsteroid).SpinVector = -.SpinVector
                            .SpinVector = sngSpin
                            
                            ' *********************************
                            ' CRAZY-ASS MULTIPLICATION!!!
                            ' (REMark out for normal operation)
                            ' *********************************
                            .Vector = Vec3MultiplyByScalar(.Vector, 1.01)
                            m_Asteroids(intOtherAsteroid).Vector = Vec3MultiplyByScalar(m_Asteroids(intOtherAsteroid).Vector, 1.01)
                            
                            
                        End If
                        
                    End If ' Is Other Asteroid Enabled?
                Next intOtherAsteroid
            End If ' Is Asteroid Enabled?
            
            .TempVar1 = 0

        End With
    Next intAsteroid

End Sub


Public Sub Init_ViewMapping()

    ' The aspect ratio of most screen resolutions (ie. 1024x768 or 800x600) have an aspect ratio of 1.332 : 1
    ' Therefore I have made the Game's World coordinates slightly wider, so that everything looks square when
    ' you maximize the form.
    
    
    ' ===============================
    ' Set the size of the Game World.
    ' ===============================
    '   * The positive X axis points towards the right.
    '   * The positive Y axis points upwards to the top of the screen.
    '   * The positive Z axis points *out of* the monitor. This is used for rotation.
    m_GameWorld.xMin = (-500 * 1.333)
    m_GameWorld.xMax = (500 * 1.333)
    m_GameWorld.yMin = -500
    m_GameWorld.yMax = 500
    
    
    ' Set the size of the window, through which we will view the Game world.
    ' (Change this window to scroll and zoom around the Game World)
    If (m_Window.xMin = m_Window.xMax) Then
        m_Window.xMin = m_GameWorld.xMin
        m_Window.xMax = m_GameWorld.xMax
        m_Window.yMin = m_GameWorld.yMin
        m_Window.yMax = m_GameWorld.yMax
    End If
    
    
    ' ==================================================================
    ' Set the size of the ViewPort (ie. normally a form, or picture box.
    ' ==================================================================
    '   This is normally set to the size of the form's internal drawing area (ie. ScaleWidth & ScaleHeight)
    m_ViewPort.xMin = 0
    m_ViewPort.xMax = frmCanvas.ScaleWidth
    m_ViewPort.yMin = frmCanvas.ScaleHeight
    m_ViewPort.yMax = 0
    
    
    ' ==========================
    ' Set the ViewMapping matrix
    ' ==========================
    g_matViewMapping = MatrixViewMapping(m_Window, m_ViewPort)
    
    Dim pixelViewPort As mdrWindow
    pixelViewPort.xMin = 0
    pixelViewPort.xMax = frmCanvas.ScaleWidth / Screen.TwipsPerPixelX
    pixelViewPort.yMin = frmCanvas.ScaleHeight / Screen.TwipsPerPixelY
    pixelViewPort.yMax = 0
    
    g_matViewMapping2 = MatrixViewMapping(m_Window, pixelViewPort)
    
End Sub

Private Sub Refresh_GameScreen(CurrentForm As Form)

    ' If I see one more game that uses BitBlt - I am going to Scream!  Arrrggghhhh!!!!!
    ' =================================================================================
    '   * You don't need BitBlt. I have absolutely no clue why people use it 99% of the time, when it is simply not needed.
    '   * You don't need DoEvents (Actually, this can cause more problems than it solves, so do yourself
    '     a big-big-BIG favour and just pretend it doesn't exist.)
    '   * You don't need to use Refresh (unless you want to slow down your program... which might be good for debugging)
    '     Set the form (or pictureboxes) AutoDraw property to True.
    '   * You don't need to use more than a single Timer control for your game.... really... you don't!
    '   * Try to learn a few API's, Particulary for drawing graphics and handling the mouse and keyboard input.
    '     They are not too hard once you get the hang of them.  Some are REALLY easy to use!
    '   * Try to get your head around 'coordinate systems'. In this Asteroids game I have several of them, and often
    '     switch between them. I will admit this can be very confusing, but this is what professional game programmers
    '     do all the time! This is how they get super intelligent AI routines, or clever collision detection etc.
    
    
    ' Important: Clear the screen before drawing anything.
    ' ====================================================
    ' This may sound obvious to some, but if you don't then you'll
    ' need to use BitBlt (or something similar) which would be a bit silly.
    ' You should always try to minimise the flicker in your game... only when this fails, should you use a Blittling Process.
    CurrentForm.BackColor = vbBlack ' RGB(64, 64, 64)
    If g_blnDontClearScreen = False Then CurrentForm.Cls
    
'    Call DrawCrossHairs(CurrentForm)
    Call Draw_Faces3(m_Asteroids, CurrentForm)
    Call Draw_Faces(m_Particles, CurrentForm)
    
End Sub

