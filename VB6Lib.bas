Attribute VB_Name = "VB6Lib"
Option Explicit

Public Const DLL_PROCESS_DETACH = 0
Public Const DLL_PROCESS_ATTACH = 1
Public Const DLL_THREAD_ATTACH = 2
Public Const DLL_THREAD_DETACH = 3

Public Function DllMain(hInst As Long, fdwReason As Long, lpvReserved As Long) As Boolean
   Select Case fdwReason
      Case DLL_PROCESS_DETACH: ' No per-process cleanup needed
      Case DLL_PROCESS_ATTACH: DllMain = True
      Case DLL_THREAD_ATTACH: ' No per-thread initialization needed
      Case DLL_THREAD_DETACH: ' No per-thread cleanup needed
   End Select
End Function

' Return a Fibonacci number.
Public Function Fibo(ByVal N As Integer) As Long
    If N <= 1 Then
        Fibo = 1
    Else
        Fibo = Fibo(N - 1) + Fibo(N - 2)
    End If
End Function

Public Function vbCDate(ByVal V As Variant):  vbCDate = CDate(V): End Function
Public Function vbDate() As Date:  vbDate = Date: End Function
Public Function vbDateAdd() As Date: vbDateAdd = DateAdd: End Function

'DateDiff  Returns the number of intervals between two dates
'DatePart  Returns the specified part of a given date
'DateSerial  Returns the date for a specified year, month, and day
'DateValue Returns a date
'Day Returns a number that represents the day of the month (between 1 and 31, inclusive)
'FormatDateTime  Returns an expression formatted as a date or time
'Hour  Returns a number that represents the hour of the day (between 0 and 23, inclusive)
'IsDate  Returns a Boolean value that indicates if the evaluated expression can be converted to a date
'Minute  Returns a number that represents the minute of the hour (between 0 and 59, inclusive)
'Month Returns a number that represents the month of the year (between 1 and 12, inclusive)
'MonthName Returns the name of a specified month
'Now Returns the current system date and time
'Second  Returns a number that represents the second of the minute (between 0 and 59, inclusive)
'Time  Returns the current system time
'Timer Returns the number of seconds since 12:00 AM
'TimeSerial  Returns the time for a specific hour, minute, and second
'TimeValue Returns a time
'Weekday Returns a number that represents the day of the week (between 1 and 7, inclusive)
'WeekdayName Returns the weekday name of a specified day of the week
'Year

