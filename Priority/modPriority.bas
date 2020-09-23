Attribute VB_Name = "modPriority"
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const HIGH_PRIORITY_CLASS = &H80

Public Declare Function SetThreadPriority Lib "KERNEL32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SetPriorityClass Lib "KERNEL32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetCurrentThread Lib "KERNEL32" () As Long
Public Declare Function GetCurrentProcess Lib "KERNEL32" () As Long
