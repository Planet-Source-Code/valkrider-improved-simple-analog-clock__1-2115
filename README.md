<div align="center">

## Improved Simple Analog Clock


</div>

### Description

Displays a basic analog clock on a form.

CREDIT FOR THE ORIGINAL CODE GOES TO: Boriza

I enjoyed his code so much I had to add a couple

of quick improvements. Incrementally uncomment

the code lines in the Form_Load event for

different clock positions on the form.
 
### More Info
 
Just add a timer("timer1") to a form and

paste this code into the form.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ValkRider](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/valkrider.md)
**Level**          |Unknown
**User Rating**    |4.2 (164 globes from 39 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/valkrider-improved-simple-analog-clock__1-2115/archive/master.zip)





### Source Code

```
Option Explicit
Private Const Offset = 50		' Border offset
Private cX As Single, cY As Single	' Center Point
Private r As Integer			' Radius
Private Sub Form_DblClick()
  ' Allow form double-click to unload clock
  Unload Me
End Sub
Private Sub Form_Load()
  ' Remove redraw flicker
  Me.AutoRedraw = True
  Timer1.Interval = 500
  ' Clock size (radius)
  r = 500
  ' You can center clock on the form...
  cX = Me.Width / 2 - Offset
  cY = Me.Height / 2 - Offset * 2
  ' OR you can put clock top-left on form...
  ' UNCOMMENT TO SEE
'  cX = r + Offset * 2
'  cY = r + Offset * 2
  ' OR even a kind of combination - REMOVE THE FORM's CAPTION AND
  '                 CONTROL BOX FOR FULL EFFECT.
  ' UNCOMMENT TO SEE
'  Me.Width = r * 2.5
'  Me.Height = r * 2.5
'  cX = Me.Width / 2 - Offset / 2
'  cY = Me.Height / 2 - Offset / 2
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Timer1.Enabled = False
End Sub
Private Sub Timer1_Timer()
  Static i As Integer
  Me.Cls
  Me.PSet (cX, cY), vbWhite
  '----------
  'print face
  '----------
  '12 O'Clock
  SetPoint 58, 0.99
  Print "12"
  '3 O'Clock
  SetPoint 13, 0.85
  Print "3"
  '6 O'Clock
  SetPoint 31, 0.7
  Print "6"
  '9 O'Clock
  SetPoint 47, 1
  Print "9"
  '-------
  'seconds
  '-------
  DrawLine Second(Now), 6, 0.98, 1
  '-------
  'minutes
  '-------
  DrawLine Minute(Now), 6, 0.9, 3
  '-------
  'hour
  '-------
  DrawLine Hour(Now), 30, 0.6, 4
  '-------
  'border
  '-------
  Me.DrawWidth = 2
  Me.Circle (cX, cY), r + Offset
End Sub
Private Sub SetPoint(Position As Integer, StartPercent As Single)
  CurrentX = Sin((180 - Position * 6) * 3.1415926 / 180) * _
       (StartPercent * r) + cX
  CurrentY = Cos((180 - Position * 6) * 3.1415926 / 180) * _
       (StartPercent * r) + cY
End Sub
Private Sub DrawLine(Position As Integer, Units As Integer, _
           LengthPercent As Single, Size As Integer)
  Me.DrawWidth = Size
  Me.Line (cX, cY)-(Sin((180 - Position * Units) * _
       3.1415926 / 180) * (LengthPercent * r) + cX, _
       Cos((180 - Position * Units) * 3.1415926 / 180) * _
       (LengthPercent * r) + cY)
End Sub
```

