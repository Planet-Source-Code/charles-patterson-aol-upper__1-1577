<div align="center">

## AOL Upper


</div>

### Description

Application creates a systray icon which has a pop-up menu allowing you to minimize and restore the AOL3.0/4.0 Upload window.

Good demo code for systray icon with popup menus no matter what you want to use it for. Also includes my Time Delay code.
 
### More Info
 
Should compile as is. No warranties expressed or implied.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Charles Patterson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/charles-patterson.md)
**Level**          |Unknown
**User Rating**    |3.3 (10 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/charles-patterson-aol-upper__1-1577/archive/master.zip)

### API Declarations

```
Option Explicit
Public Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uID As Long
 uFlags As Long
 uCallbackMessage As Long
 hIcon As Long
 sTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
'Make your own constant, e.g.:
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const SW_RESTORE = 9
Public Const SW_MINIMIZE = 6
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
```


### Source Code

```
'First create a form with a menu item listing 3 sub menus. mnuExit, mnuMinUpload and mnuResUpload.
Option Explicit
Dim Tic As NOTIFYICONDATA
Private Sub Form_Activate()
 Dim TimeDelay&
 Label2.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision & " " & Label2.Caption
 TimeDelay = Timer + 3
 While Timer <= TimeDelay
  DoEvents
 Wend
 Me.Hide
 mnuSystemTray.Visible = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'Event occurs when the mouse pointer is within the rectangular
 'boundaries of the icon in the taskbar status area.
 Dim msg As Long
 Dim sFilter As String
 msg = X / Screen.TwipsPerPixelX
 Select Case msg
  Case WM_LBUTTONDBLCLK
   mnuMinUpload_Click
  Case WM_RBUTTONUP
   PopupMenu mnuSystemTray, , , , mnuMinUpload
 End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Shell_NotifyIcon NIM_DELETE, Tic
End Sub
Private Sub Form_Load()
 If App.PrevInstance Then End
 Dim rc As Long
 Tic.cbSize = Len(Tic)
 Tic.hwnd = Me.hwnd
 Tic.uID = vbNull
 Tic.uFlags = NIF_DOALL
 Tic.uCallbackMessage = WM_MOUSEMOVE
 Tic.hIcon = Me.Icon
 Tic.sTip = "AOL Upload Minimizer" & vbNullChar
 rc = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
Private Sub mnuExit_Click()
 End
End Sub
Private Sub mnuResUpload_Click()
 Dim AOL As Long
 Dim AOModal As Long
 Dim AOGauge As Long
 AOL = FindWindow("AOL Frame25", vbNullString)
 AOModal = FindWindow("_AOL_Modal", vbNullString)
 AOGauge = FindChildByClass(AOModal, "_AOL_Gauge")
 If AOGauge <> 0 Then
  EnableWindow AOL, 1
  ShowWindow AOModal, SW_RESTORE
 End If
End Sub
Private Sub mnuMinUpload_Click()
 Dim AOL As Long
 Dim AOModal As Long
 Dim AOGauge As Long
 AOL = FindWindow("AOL Frame25", vbNullString)
 AOModal = FindWindow("_AOL_Modal", vbNullString)
 AOGauge = FindChildByClass(AOModal, "_AOL_Gauge")
 If AOGauge <> 0 Then
  EnableWindow AOL, 1
  ShowWindow AOModal, SW_MINIMIZE
 End If
End Sub
Private Function FindChildByClass(Parent&, Child$) As Integer
 Dim ChildFocus%, Buffer$, ClassBuffer%
 ChildFocus% = GetWindow(Parent, 5)
 While ChildFocus%
  Buffer$ = String$(250, 0)
  ClassBuffer% = GetClassName(ChildFocus%, Buffer$, 250)
  If InStr(UCase(Buffer$), UCase(Child)) Then
   FindChildByClass = ChildFocus%
   Exit Function
  End If
  ChildFocus% = GetWindow(ChildFocus%, 2)
 Wend
End Function
```

