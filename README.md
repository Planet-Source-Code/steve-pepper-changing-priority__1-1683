<div align="center">

## Changing priority


</div>

### Description

This code fragment demonstrates how you can get a program to change it's own priority.

Sometimes it is necessary to change the priority of a process from the default.

The example that prompted this code was a program I had to launch from a commercial scheduler always defaulted to idle priority. This caused the program to miss it's processing deadline. Increasing the priority to normal or high solved the problem.
 
### More Info
 
To run this code, add a form to a new project with a timer and a label.

The priority will toggle between idle and high every two seconds.

Under NT4, you can use the task manager to see the base priority of this process changing.

Changing the priority of your process to REALTIME_PRIORITY_CLASS is a bad idea in Visual Basic.

I have only tested this code under NT4 sp4 and VB5 but I think it should work under Windows9x


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Pepper](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-pepper.md)
**Level**          |Unknown
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-pepper-changing-priority__1-1683/archive/master.zip)

### API Declarations

```
Option Explicit
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100
Private Const PROCESS_DUP_HANDLE = &H40
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function SetPriorityClass& Lib "kernel32" (ByVal hProcess As Long, _
    ByVal dwPriorityClass As Long)
```


### Source Code

```
Sub ChangePriority(dwPriorityClass As Long)
  Dim hProcess&
  Dim ret&, pid&
  pid = GetCurrentProcessId() ' get my proccess id
  ' get a handle to the process
  hProcess = OpenProcess(PROCESS_DUP_HANDLE, True, pid)
  If hProcess = 0 Then
    Err.Raise 2, "ChangePriority", "Unable to open the source process"
    Exit Sub
  End If
  ' change the priority
  ret = SetPriorityClass(hProcess, dwPriorityClass)
  ' Close the source process handle
  Call CloseHandle(hProcess)
  If ret = 0 Then
    Err.Raise 4, "ChangePriority", "Unable to close source handle"
    Exit Sub
  End If
End Sub
Private Sub Form_Load()
  Timer1.Interval = 2000
  Call Timer1_Timer
End Sub
Private Sub Timer1_Timer()
  Static Priority&
  If Priority = IDLE_PRIORITY_CLASS Then
   Priority = HIGH_PRIORITY_CLASS
   Label1.Caption = "HIGH priority"
  Else
   Label1.Caption = "IDLE priority"
   Priority = IDLE_PRIORITY_CLASS
  End If
  Call ChangePriority(Priority)
End Sub
```

