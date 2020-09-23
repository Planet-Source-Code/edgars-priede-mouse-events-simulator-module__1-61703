Attribute VB_Name = "modMouseEvents"
'///////////////////////////////////////////////////////////
'//       Module: Mouse Events Simulator                  //
'//       Author: Edgars Priede                           //
'//       E-Mail: edgars.software@inbox.lv                //
'//       Date:   15.07.2005                              //
'//                                                       //
'//       Description: With this module you can           //
'//                    simulate mouse events.             //
'//                    Enjoy! :)                          //
'//                                                       //
'//       Events: MouseMove; MouseLeftClick;              //
'//               MouseLeftDbClick; MouseLeftDown;        //
'//               MouseLeftUp; MouseRightClick;           //
'//               MouseRightDbClick; MouseRightDown;      //
'//               MouseRightUp; MouseMiddleClick;         //
'//               MouseMiddleDbClick; MouseMiddleDown;    //
'//               MouseMiddleUp; GetMousePos; SetPause.   //
'///////////////////////////////////////////////////////////


'API Declarations
Private Declare Sub mouse_event Lib "user32" _
(ByVal dwFlags As Long, ByVal dx As Long, _
ByVal dy As Long, ByVal cButtons As Long, _
ByVal dwExtraInfo As Long)

Private Declare Function SetCursorPos Lib "user32" _
(ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

Private Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)

'Types
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'Mouse Events Constants
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10


'////////////////////////////////////////////////////////////
'///////////////////// Simple Set Pause /////////////////////
'////////////////////////////////////////////////////////////

Public Sub SetPause(ByVal Seconds As Integer)
    
    Call Sleep(Seconds * 1000) 'Call Sleep API for pause
    
End Sub

'////////////////////////////////////////////////////////////
'///////////////////// Move Mouse Event /////////////////////
'////////////////////////////////////////////////////////////

Public Sub MouseMove(ByVal PosX As Long, ByVal PosY As Long)
    
    Call SetCursorPos(PosX, PosY) 'Move cursor
    
End Sub

'////////////////////////////////////////////////////////////
'/////////////////// Get Mouse Position /////////////////////
'////////////////////////////////////////////////////////////

Public Sub GetMousePos(PosX As Long, PosY As Long)

    Dim Pos As POINTAPI
    
    Call GetCursorPos(Pos) 'Get cursor position
    
    PosX = Pos.X '\
'                  Set position coordinates
    PosY = Pos.Y '/
    
End Sub


'////////////////////////////////////////////////////////////
'///////////////// Left Mouse Button Events /////////////////
'////////////////////////////////////////////////////////////

Public Sub MouseLeftClick(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, PosX, PosY, 0, 0) 'Set Left Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_LEFTUP, PosX, PosY, 0, 0) 'Set Left Mouse Button up
    
End Sub

Public Sub MouseLeftDbClick(ByVal PosX As Long, ByVal PosY As Long, ByVal Delay As Integer)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, PosX, PosY, 0, 0) 'Set Left Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_LEFTUP, PosX, PosY, 0, 0) 'Set Left Mouse Button up
    
    Call SetPause(Delay) 'Set pause , because Double Click
    
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, PosX, PosY, 0, 0) 'Set Left Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_LEFTUP, PosX, PosY, 0, 0) 'Set Left Mouse Button up

End Sub

Public Sub MouseLeftDown(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, PosX, PosY, 0, 0) 'Set Left Mouse Button down

End Sub

Public Sub MouseLeftUp(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_LEFTUP, PosX, PosY, 0, 0) 'Set Left Mouse Button up
    
End Sub

'////////////////////////////////////////////////////////////
'//////////////// Right Mouse Button Events /////////////////
'////////////////////////////////////////////////////////////

Public Sub MouseRightClick(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_RIGHTDOWN, PosX, PosY, 0, 0) 'Set Right Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_RIGHTUP, PosX, PosY, 0, 0) 'Set Right Mouse Button up
    
End Sub

Public Sub MouseRightDbClick(ByVal PosX As Long, ByVal PosY As Long, ByVal Delay As Integer)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_RIGHTDOWN, PosX, PosY, 0, 0) 'Set Right Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_RIGHTUP, PosX, PosY, 0, 0) 'Set Right Mouse Button up
    
    Call SetPause(Delay) 'Set pause, because Double Click
    
    Call mouse_event(MOUSEEVENTF_RIGHTDOWN, PosX, PosY, 0, 0) 'Set Right Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_RIGHTUP, PosX, PosY, 0, 0) 'Set Right Mouse Button up

    
End Sub


Public Sub MouseRightDown(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_RIGHTDOWN, PosX, PosY, 0, 0) 'Set Right Mouse Button down
  
End Sub


Public Sub MouseRightUp(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_RIGHTUP, PosX, PosY, 0, 0) 'Set Right Mouse Button up
  
End Sub

'////////////////////////////////////////////////////////////
'/////////////// Middle Mouse Button Events /////////////////
'////////////////////////////////////////////////////////////

Public Sub MouseMiddleClick(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, PosX, PosY, 0, 0) 'Set Middle Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_MIDDLEUP, PosX, PosY, 0, 0) 'Set Middle Mouse Button up
    
End Sub

Public Sub MouseMiddleDbClick(ByVal PosX As Long, ByVal PosY As Long, ByVal Delay As Integer)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, PosX, PosY, 0, 0) 'Set Middle Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_MIDDLEUP, PosX, PosY, 0, 0) 'Set Middle Mouse Button up
    
    Call SetPause(Delay) 'Set pause, because Double Click
    
    Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, PosX, PosY, 0, 0) 'Set Middle Mouse Button down
    
    Call mouse_event(MOUSEEVENTF_MIDDLEUP, PosX, PosY, 0, 0) 'Set Middle Mouse Button up

    
End Sub


Public Sub MouseMiddleDown(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, PosX, PosY, 0, 0) 'Set Middle Mouse Button down
  
End Sub


Public Sub MouseMiddleUp(ByVal PosX As Long, ByVal PosY As Long)

    Call SetCursorPos(PosX, PosY) 'Move cursor

    Call SetPause(0.2) 'Set pause for 200 milliseconds
    
    Call mouse_event(MOUSEEVENTF_MIDDLEUP, PosX, PosY, 0, 0) 'Set Middle Mouse Button up
  
End Sub

