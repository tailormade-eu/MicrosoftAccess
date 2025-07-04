Attribute VB_Name = "omKernalFunctions"
Option Compare Database
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
