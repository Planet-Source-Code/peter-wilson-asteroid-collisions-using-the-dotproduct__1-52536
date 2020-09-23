Attribute VB_Name = "mKeyboard"
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const mc_strModuleName As String = "mKeyboard"

' The GetKeyState function retrieves the status of the specified virtual key. The status specifies whether the key is up, down, or toggled (on, off—alternating each time the key is pressed).
Public Const VK_LSHIFT As Long = &HA0
Public Const VK_RSHIFT As Long = &HA1
Public Const VK_LCONTROL As Long = &HA2
Public Const VK_RCONTROL As Long = &HA3
Public Const VK_LMENU As Long = &HA4
Public Const VK_RMENU As Long = &HA5
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub ProcessKeyboardInput()

    On Error GoTo errTrap
    
    Static s_strPreviousValue As String
    Static s_blnKeyDeBounce As Boolean
    Static s_blnKeyDeBounce_TAB As Boolean
    Static s_blnKeyDeBounce_FullScreen As Boolean
    
    ' Keyboard DeBounce
    Static s_lngKeyCombinations As Long
    
    Dim lngKeyCombinations As Long
    Dim lngKeyState As Long
    Dim sngSpeedIncrement As Single
    Dim sngSpeedMagnitude As Single
    Dim sngRadians As Single
    
    ' =======================
    ' DO NOT clear the screen
    ' =======================
    lngKeyState = GetKeyState(vbKeyC)
    If (lngKeyState And &H8000) Then g_blnDontClearScreen = True Else g_blnDontClearScreen = False
    
    
    ' Check the Space Bar for level complete. (Also shows how to "de-bounce" the space bar)
    lngKeyState = GetKeyState(vbKeySpace)
    If (lngKeyState And &H8000) Then
        If s_blnKeyDeBounce = False Then
            s_blnKeyDeBounce = True
            g_strGameState = "LevelComplete"
        End If
    Else
        s_blnKeyDeBounce = False
    End If
    
    
    ' Check for ESCAPE key.
    lngKeyState = GetKeyState(vbKeyEscape)
    If (lngKeyState And &H8000) Then g_strGameState = "Quit"
    
    s_lngKeyCombinations = lngKeyCombinations
    
    Exit Sub
errTrap:
    
End Sub

