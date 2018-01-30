VERSION 5.00
Begin VB.Form frmdetail 
   Caption         =   "Time Tables"
   ClientHeight    =   3105
   ClientLeft      =   5505
   ClientTop       =   3330
   ClientWidth     =   4650
   Icon            =   "frmdetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4650
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   45
      End
      Begin VB.Label Label6 
         Caption         =   "CPU #"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmdetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer
Dim j As Integer
Dim k As Integer
'show the datails--> show each cpu`s tasks time table
For i = 1 To ccount
    Load Label6(Label6.UBound + 1)
    Label6(Label6.UBound).Visible = True
    Label6(Label6.UBound).Top = Label6(i - 1).Top + 500
    Load Label7(Label7.UBound + 1)
    Label7(Label7.UBound).Visible = True
    Label7(Label7.UBound).Top = Label7(i - 1).Top + 500
    Label7(i).Caption = i & ":"
    Load Label1(Label1.UBound + 1)
    Label1(Label1.UBound).Visible = True
    Label1(Label1.UBound).Top = Label1(i - 1).Top + 500
    For j = 1 To Len(cpu_job(solution, i))
    If t(Mid(cpu_job(solution, i), j, 1)) <> 0 Then
    Label1(i).Caption = Label1(i).Caption & "t" & Mid(cpu_job(solution, i), j, 1) & " =" & Str(t(Mid(cpu_job(solution, i), j, 1)))
    Label1(i).Caption = Label1(i).Caption & "  ||  "
    End If
    Next j
Next i
Frame1.Height = Label6(ccount).Top + 500
frmdetail.Height = Frame1.Height + 900
End Sub


