VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4815
   ClientLeft      =   3540
   ClientTop       =   2070
   ClientWidth     =   8160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4770
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "using a Genetic Algorithm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Top             =   2640
         Width           =   2940
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3735
         Left            =   5400
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label Label5 
         Caption         =   "Press any key to continue..."
         Height          =   255
         Left            =   5880
         TabIndex        =   5
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parisa Kianmajd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   4
         Top             =   4200
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Task Scheduling"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   660
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   4470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Task Scheduling"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   660
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   4470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   " for multiprocessor systems"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   2280
         Width           =   3180
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    frmpro.Show
End Sub

Private Sub mnuExit_Click()
Dim intmsg As Integer
 intmsg = MsgBox("Do you really want to end the program? ", vbYesNoCancel + vbQuestion, "Confirm") 'we can use 3 instead of vbyesnocancel
 If (intmsg = vbNo) Or (intmsg = vbCancel) Then
  Cancel = True
 Else
  Unload Me
 End If
End Sub


