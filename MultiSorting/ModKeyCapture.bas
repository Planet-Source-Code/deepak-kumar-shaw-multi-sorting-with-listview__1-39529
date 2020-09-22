Attribute VB_Name = "ModKeyCapture"
'*************************************************'
'*******      Capturing Additional Key      ******'
'*************************************************'
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As _
    Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
    ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As _
    Long
Private Const KEYEVENTF_KEYUP = &H2

' press and/or release any key, given its virtual code
'
' virtKeyCode can be any vbKey* constants, except mouse constants
' Action can be 0 or omitted (press and then release the key)
'                     >0  (only press the key)
'                     <0  (only release the key)

Sub PressVirtualKey(ByVal virtKeyCode As KeyCodeConstants, _
    Optional ByVal Action As Integer)
    ' press the key if the argument is 0 or greater
    If Action >= 0 Then
        keybd_event virtKeyCode, 0, 0, 0
    End If
    ' then release the key if the argument is 0 or lesser
    If Action <= 0 Then
        keybd_event virtKeyCode, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub


' Return True if all the specified keys are pressed
'
' you can specify individual keys using VB constants,
' e.g. If KeysPressed(vbKeyControl, vbKeyDown) Then ...

Function KeysPressed(ByVal KeyCode1 As KeyCodeConstants, _
    Optional ByVal KeyCode2 As KeyCodeConstants, Optional ByVal KeyCode3 As _
    KeyCodeConstants) As Boolean
    If GetAsyncKeyState(KeyCode1) >= 0 Then Exit Function
    If KeyCode2 = 0 Then KeysPressed = True: Exit Function
    If GetAsyncKeyState(KeyCode2) >= 0 Then Exit Function
    If KeyCode3 = 0 Then KeysPressed = True: Exit Function
    If GetAsyncKeyState(KeyCode3) >= 0 Then Exit Function
    KeysPressed = True
End Function



