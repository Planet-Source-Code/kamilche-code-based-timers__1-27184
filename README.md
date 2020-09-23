<div align="center">

## Code\-Based Timers


</div>

### Description

Start and kill a timer using API calls only! Useful when you need timers that can't be placed on a form.
 
### More Info
 
The routine that will be called every Timer milliseconds, MUST be placed in a standard module! It can't be placed on a form, or in a class.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kamilche](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kamilche.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kamilche-code-based-timers__1-27184/archive/master.zip)

### API Declarations

```
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
```


### Source Code

```
'Make this a global variable, or site it
'in the same module as MainLoop.
Public Timer as Long
'To set the timer, issue the
'following, where MainLoop
'is the name of the procedure
'to call every 500 milliseconds.
'Note that MainLoop MUST exist
'in a BAS module!
Timer = SetTimer(0, 0, 500, AddressOf MainLoop)
'To kill the timer,
'issue the following:
KillTimer 0, Timer
```

