<div align="center">

## How to get the Mouse Position and how to Set the mouse position with Windows API Functions


</div>

### Description

My Code shows you how to use the windows API functions so that you can view the x,y coordinates of a mouse.

It also allows you to be able to Set the position of the mouse
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/how-to-get-the-mouse-position-and-how-to-set-the-mouse-position-with-windows-api-functions__1-8785/archive/master.zip)





### Source Code

```
'***************************************************************
'*Feel Free to use this souce whenever.            *
'*                               *
'*Author: ToddSoft                       *
'*Subject: Set Cursor Position and Get Cursor Position     *               *
'*Date: 6-9-200                        *
'*                               *
'*Hey, Check out this site: www.ToddSoft.com         *
'*                               *
'***************************************************************
in a module:
'This API call is for the SetCursorPos
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'************************************************************************************
General Declerations
'Api call for the GetCursorPos function
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'This is the data type
Private Type POINTAPI
    x As Long 'x coordinate of the mouse
    y As Long 'y coordinate of the mouse
End Type
Dim pp As POINTAPI
Private Sub Command1_Click()
SetCursorPos Form1.Left / 15, Form1.Top / 15
'You have to divide the form by 15 because the position you set it to is applied
'to the screen and not the form. By clicking on the button it moves the mouse
'directly up to the top left part of the form
Command1.Caption = "Click Me"
End Sub
Private Sub Form_Load()
MsgBox "Please visit www.ToddSoft.com", vbOKOnly, "ToddSoft"
End Sub
Private Sub Timer1_Timer()
'Note that if x = 0 and y = 0 that is the top left part of the monitor screen_
'Not the form
'This calls the GetCursorPos Function to get the x and y positions of the mouse
GetCursorPos pp
'This is displaying the x and y coordinates of the mouse
Label1.Caption = "X: " & pp.x & " Y: " & pp.y
End Sub
```

