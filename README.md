<div align="center">

## 32\-bit Floating Toolbar \(NT & 95\)


</div>

### Description

This code gives you the ability to create a 'floating toolbar' within your application. The old SetWindowWord function is only good for 16-bit applications, so it won't run under a 32-bit OS (like NT4). The API call you should use if you are programming a 32-bit application is SetWindowLong. It works the same way as SetWindowWord, only uses DWORD(Long) values instead of WORD values for the 32-bit OS.
 
### More Info
 
You will need to create 2 forms (Form1 & Form2).

On Form1, place a Command button (Command1)

On Form2, set the Window Style to 4-FixedToolWindow (not nessesary)

This function will make a form a 'child window' of any form you specify.

Sets Form2 to be a child of Form1 (giving it a 'floating toolbar' effect)

Won't work with 16-bit OS's. Use SetWindowWord for 16-bit.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-jones.md)
**Level**          |Unknown
**User Rating**    |4.0 (166 globes from 41 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-jones-32-bit-floating-toolbar-nt-95__1-1146/archive/master.zip)

### API Declarations

```
' Place this code into a module
Public Const GWL_HWNDPARENT = (-8)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
```


### Source Code

```
' Place this code in the General Declarations section of Form1.
Private Sub Command1_Click()
  'Open the toolbar window
  Form2.Show
  'Move the toolbar to the right
  'of Form1.
  '(gives it a docking effect)
  Form2.Height = Form1.Height - 330
  'Subtract the titlebar height -^
  Form2.Left = Form1.Left + Form1.Width - Form2.Width
  Form2.Top = Form1.Top + Form1.Height - Form2.Height
End Sub
Private Sub Form_Load()
  'Set the button properties
  Command1.Caption = "Show Toolbar"
  Command1.Width = 2055
  Command1.Height = 375
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'If Form2 is opened when you close
  'Form1, it will not end your app, so
  'you have to manually unload Form2.
  Unload Form2
End Sub
' Place this code in the Form_Load event of Form2
Private Sub Form_Load()
SetWindowLong Me.hwnd, GWL_HWNDPARENT, Form1.hwnd
End Sub
```

