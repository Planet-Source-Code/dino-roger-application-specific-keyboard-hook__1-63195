<div align="center">

## Application\(Specific\) Keyboard Hook


</div>

### Description

This code is a combination of a few great submissions at PSC put together. Needed code that would load/unload a system wide keyboard hook but only show me the keys when a specific application had focus. The example uses Notepad. When a key is pressed only in Notepad the key code will be show as the form caption.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dino Roger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dino-roger.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dino-roger-application-specific-keyboard-hook__1-63195/archive/master.zip)





### Source Code

```
'INSTRUCTIONS:
'1.) Create a form called frmMain and place a 100 interval timer called Timer1
'2.) Set the startup object as Sub Main. )Project - Project Properties)
'-----------------------------
'In a module -----------------
'-----------------------------
Option Explicit
Public hKbdHook As Long
Private Const WH_KEYBOARD_LL As Integer = 13
Private Const HC_ACTION As Integer = 0
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Type KBDLLHOOKSTRUCT
 vkCode As Integer
 scanCode As Integer
 flags As Integer
 time As Integer
 dwExtraInfo As Integer
End Type
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public AppIsActive As Boolean
Private Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Integer
 Dim kbdllhs As KBDLLHOOKSTRUCT
 CopyMemory kbdllhs, ByVal lParam, Len(kbdllhs)
 If nCode = HC_ACTION Then
  LowLevelKeyboardProc = CallNextHookEx(hKbdHook, nCode, wParam, lParam)
  Select Case wParam
   Case WM_KEYDOWN
    If AppIsActive = True Then frmMain.Caption = kbdllhs.vkCode
   Case WM_KEYUP
  End Select
  Else: LowLevelKeyboardProc = CallNextHookEx(hKbdHook, nCode, wParam, lParam)
 End If
End Function
Sub Main()
 AppIsActive = False
 hKbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&)
 If hKbdHook = 0 Then
  MsgBox "Initialisation of keyboard hook failed.", vbCritical, "Keyboard Hook"
  Exit Sub
 End If
 frmMain.Show
End Sub
'---------------------------------------------
'In a form called frmMain---------------------
'---------------------------------------------
Option Explicit
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Const GW_HWNDNEXT = 2
Const GW_OWNER = 4
Dim AppName As String
Public Function GetWinCaption(hwnd) As String
 Dim sTitle As String
 sTitle = String(GetWindowTextLength(hwnd), 0)
 GetWindowText hwnd, sTitle, GetWindowTextLength(hwnd) + 1
 GetWinCaption = sTitle
End Function
Private Sub Form_Load()
 AppName = "notepad"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Call UnhookWindowsHookEx(hKbdHook)
End Sub
Private Sub Timer1_Timer()
 If InStr(1, UCase(GetWinCaption(GetForegroundWindow())), UCase(AppName)) > 0 Then
  AppIsActive = True
  Else
   AppIsActive = False
 End If
End Sub
```

