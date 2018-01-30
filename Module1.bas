Attribute VB_Name = "Module1"
'-----------public variables--------------
Public ptmatrix(1 To 50, 1 To 9, 1 To 9) As Boolean
Public pcount As String 'number of processes
Public ccount As String 'number of processors
Public t(1 To 9) As Single 'tasks list
Public cpu(1 To 50, 1 To 9) As Integer
Public tft(1 To 50) As Integer 'total finish time for each chromosome
Public cpu_job(1 To 50, 1 To 9) As String
Public ft(1 To 9) As Single
Public wt(1 To 9) As Single
Public art As Single 'average response time
Public awt As Single 'average wait time
Public solution As Integer
Public tem(1 To 1000, 1 To 9, 1 To 9) As Boolean
