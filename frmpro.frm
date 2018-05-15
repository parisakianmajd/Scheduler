VERSION 5.00
Begin VB.Form frmpro 
   Caption         =   "task  Scheduling for Multiprocessor Systems"
   ClientHeight    =   4770
   ClientLeft      =   5115
   ClientTop       =   3330
   ClientWidth     =   5595
   ClipControls    =   0   'False
   Icon            =   "frmpro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5595
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox ptxt 
         Height          =   375
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox ctxt 
         Height          =   375
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdenter 
         Caption         =   "&Enter"
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   405
         HideSelection   =   0   'False
         Index           =   0
         Left            =   720
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txttft 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox fttxt 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   7
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox wttxt 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox awttxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox arttxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Enter"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdext 
         Caption         =   "&EXIT"
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmddetail 
         Caption         =   "&Details>>"
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblawt 
         AutoSize        =   -1  'True
         Height          =   315
         Left            =   8280
         TabIndex        =   25
         Top             =   1920
         Width           =   165
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         Height          =   315
         Left            =   8280
         TabIndex        =   24
         Top             =   960
         Width           =   165
      End
      Begin VB.Label labelwt 
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label labelrt 
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Labelft 
         Height          =   315
         Index           =   0
         Left            =   4680
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Number of processes"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Number of processors"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Taks:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbltft 
         Caption         =   "TFT"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   0
         X2              =   5160
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   0
         X2              =   5160
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblft 
         Caption         =   "FT:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblwt 
         Caption         =   "WT:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "AWT"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   4200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "ART"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmpro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim population As Integer 'the firsy generation
Dim strvalid As String
Private Sub cmddetail_Click()
frmdetail.Show
End Sub
Private Sub cmdenter_Click()
Dim interror As Integer
Dim i As Integer
' the default value is 0
If ctxt.Text = "" Then ctxt.Text = 0
If ptxt.Text = "" Then ptxt.Text = 0
ccount = ctxt.Text
pcount = ptxt.Text
'  number of processors  >= 2.. a multiprocessor system
If (ccount < 2) Or (pcount < 2) Then
interror = MsgBox("Invalid value", vbCritical)
ctxt.Text = ""
ptxt.Text = ""
Exit Sub
End If
cmdenter.Visible = False
ctxt.Enabled = False
ptxt.Enabled = False
Label3.Visible = True
Command1.Visible = True
Line1.Visible = True
'-----------loading-----
For i = 1 To pcount
    Load Text1(Text1.UBound + 1)
    Text1(Text1.UBound).Visible = True
    Text1(Text1.UBound).Left = (i * 500) + 450
Next i
'adjusting the size of the form
  Frame1.Width = Text1(pcount).Left + 800
  If Frame1.Width < 3900 Then Frame1.Width = 3900
  frmpro.Width = Frame1.Width + 400
  Line1.X2 = Frame1.Width
  Line2.X2 = Frame1.Width
  Frame1.Height = Frame1.Height + 2500
  frmpro.Height = Frame1.Height + 900
End Sub
Private Sub cmdext_Click()
Unload Me
End Sub
Private Sub Command1_Click()
Dim i As Integer
For i = 1 To pcount
  If Text1(i).Text = "" Then Text1(i).Text = "0"
Next i
For i = 1 To pcount
 t(i) = Text1(i).Text
Next i
Call initialize
'Call selection
Call crossover
Call mutation
Call find_fitness
End Sub
Private Sub ptxt_KeyPress(KeyAscii As Integer)
' only numeric values are accepted
  If (KeyAscii > 26) Then 'if is not a control key
  If InStr(strvalid, Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
  End If
 End If
End Sub
Private Sub ctxt_KeyPress(KeyAscii As Integer)
  If (KeyAscii > 26) Then 'if is not a control key
  If InStr(strvalid, Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
  End If
 End If
End Sub
Private Sub Form_Load()
frmpro.Height = 2340
frmpro.Width = 4695
Frame1.Height = 1455
Frame1.Width = 4335
strvalid = "0123456789"
population = 50 'make 50 chromosomes as the first generation
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  If (KeyAscii > 26) Then 'if is not a control key
  If InStr(strvalid, Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
  End If
 End If
End Sub
Private Sub initialize()
Dim i As Integer
Dim j As Integer
Dim k As Integer
For k = 1 To population
For i = 1 To pcount
'each chromosome is shown as a binary matrix. with the rows showing the number of processes(pcount)
'and the column showing the number of processors(pcount)
'each process is run on a processor---i.e. each column matrix just one row (selected randomly) is 1 and the rest are 0
'----------------------------------
'this proc initializes the first generation (random matrixes)
'finds a random number between 1 & ccount (that is a random row in each column)
Randomize
j = Round(Rnd * (ccount - 1)) + 1
ptmatrix(k, j, i) = True
Next i
Next k
End Sub
Private Sub selection()
'we use "roulette-wheel selection" for selecting potentially useful solutions for recombination.
'we want the chromosomes with lower tft to have more chance to be selected
'I have defined an array of 1000 chromosomes(roulette-wheel) ... each chromosome would be repeated (sum of all tfts/its tft) times in this roulette-wheel
'then we select 50 random chromosomes form roulette-wheel for the next generation
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l, c As Integer
Dim sumtft As Integer
Dim repeat(1 To 50) As Integer 'probability of being selected for each chromosome
For k = 1 To population
For i = 1 To ccount
For j = 1 To pcount
If ptmatrix(k, i, j) = True Then
cpu(k, i) = cpu(k, i) + t(j)
cpu_job(k, i) = cpu_job(k, i) & j
End If
Next j
Next i
Next k
'find tft for each chromosome
For k = 1 To population
tft(k) = cpu(k, 1)
For i = 1 To ccount
If cpu(k, i) > tft(k) Then tft(k) = cpu(k, i)
Next i
Next k
For i = 1 To 50
sumtft = sumtft + tft(i)
Next i
For i = 1 To 50
repeat(i) = sumtft / tft(i)
Next i
c = 1
For i = 1 To 50
For j = 1 To repeat(i)
For k = 1 To 9
For l = 1 To 9
tem(c, k, l) = ptmatrix(i, k, l)
Next l
Next k
c = c + 1
If c > 1000 Then GoTo ex
Next j
Next i
ex:
End Sub
Private Sub crossover()
Dim c1, c2 As Integer
Dim k As Integer
Dim i As Integer
Dim j As Integer
Dim cross_point As Integer
Dim temp(1 To 2, 1 To 9, 1 To 9) As Boolean
For k = 1 To population
'find 2 random chromosome
Randomize
c1 = Round(Rnd * (population - 1)) + 1
Randomize
c2 = Round(Rnd * (population - 1)) + 1
'find a random crossover point between 1 & pcount
Randomize
cross_point = Round(Rnd * (pcount - 1)) + 1
'make two new chromosome
For i = 1 To ccount
For j = 1 To cross_point
temp(1, i, j) = ptmatrix(c1, i, j)
temp(2, i, j) = ptmatrix(c2, i, j)
Next j
Next i
For i = 1 To ccount
For j = cross_point To pcount
temp(1, i, j) = ptmatrix(c2, i, j)
temp(2, i, j) = ptmatrix(c1, i, j)
Next j
Next i
'replace new chromosomes with existing ones
For i = 1 To ccount
For j = 1 To pcount
ptmatrix(c1, i, j) = temp(1, i, j)
ptmatrix(c2, i, j) = temp(2, i, j)
Next j
Next i
Next k
End Sub
Private Sub mutation()
Dim i As Integer
Dim k As Integer
Dim ch, col, row, r As Integer
For k = 1 To population
'find a random chromosome
Randomize
ch = Round(Rnd * (population - 1)) + 1
Randomize
'find a random column in this chromosome
col = Round(Rnd * (pcount - 1)) + 1
'find which one of the rows in this column is 1, change it to 0 and make another random row 1
For i = 1 To ccount
If ptmatrix(ch, i, col) = True Then
ptmatrix(ch, i, col) = False
r = i
End If
Next i
1:
Randomize
'find a random row in the currnet column
row = Round(Rnd * (ccount - 1)) + 1
If r = row Then GoTo 1
ptmatrix(ch, row, col) = True
Next k
End Sub
Private Sub avg()
Dim i As Integer
'find art and awt
For i = 1 To pcount
art = art + ft(i)
wt(i) = ft(i) - t(i)
wttxt(i).Text = wt(i)
awt = awt + wt(i)
Next i
art = art / pcount
awt = awt / pcount
awttxt.Text = Round(awt, 3)
arttxt.Text = Round(art, 3)
awttxt.Visible = True
arttxt.Visible = True
Label5.Visible = True
Label4.Visible = True
Frame1.Height = 5000
frmpro.Height = Frame1.Height + 900
For i = 1 To pcount
Text1(i).Enabled = False
Next i
cmdext.Visible = True
cmddetail.Visible = True
End Sub
Private Sub showresult()
Dim i As Integer
Command1.Visible = False
For i = 1 To pcount
    Load fttxt(fttxt.UBound + 1)
    fttxt(fttxt.UBound).Visible = True
    fttxt(fttxt.UBound).Left = (i * 500) + 450
    Load wttxt(wttxt.UBound + 1)
    wttxt(wttxt.UBound).Visible = True
    wttxt(wttxt.UBound).Left = (i * 500) + 450
Next i
For i = 1 To pcount
 fttxt(i).Text = ft(i)
 wttxt(i).Text = wt(i)
Next i
lblft.Visible = True
lblwt.Visible = True
Line2.Visible = True
End Sub

Private Sub find_fitness()
Dim temp As Integer
'I have chosen tft as the fitness criteria
'each chromosome that have the lowest tft would be the answer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim lowest_tft As Integer
For k = 1 To population
For i = 1 To ccount
For j = 1 To pcount
If ptmatrix(k, i, j) = True Then
cpu(k, i) = cpu(k, i) + t(j)
cpu_job(k, i) = cpu_job(k, i) & j
End If
Next j
Next i
Next k
'find tft for each chromosome
For k = 1 To population
tft(k) = cpu(k, 1)
For i = 1 To ccount
If cpu(k, i) > tft(k) Then tft(k) = cpu(k, i)
Next i
Next k
1:
'find the lowest tft
lowest_tft = tft(1)
For k = 1 To population
If tft(k) < lowest_tft Then
lowest_tft = tft(k)
solution = k
End If
Next k
'-------------------
For i = 1 To ccount
temp = 0
If (solution = 0) Then GoTo 1
For j = 1 To Len(cpu_job(solution, i))
temp = temp + t(Mid(cpu_job(solution, i), j, 1))
ft(Mid(cpu_job(solution, i), j, 1)) = temp
Next j
Next i
txttft.Visible = True
txttft.Text = lowest_tft
lbltft.Visible = True
Call showresult
Call avg
End Sub





