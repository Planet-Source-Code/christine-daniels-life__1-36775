VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "LIFE"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1365
      TabIndex        =   7
      Top             =   4050
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speed"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3375
      Width           =   3615
      Begin VB.OptionButton OptSpeed 
         Caption         =   "Slow"
         Height          =   195
         Index           =   2
         Left            =   2655
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptSpeed 
         Caption         =   "Medium"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptSpeed 
         Caption         =   "Quickest"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1620
      Top             =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   2610
      TabIndex        =   2
      Top             =   4050
      Width           =   1095
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Random"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4050
      Width           =   1095
   End
   Begin VB.Label Vakje 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   315
      TabIndex        =   0
      ToolTipText     =   "Click on a square to change color"
      Top             =   90
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const AANTALVAKJES = 900
Dim Kleur As ColorConstants


Private Type Lijn       'backup voor de vakjes
    vak(30) As Integer
End Type

Dim Spelgestart As Boolean

Private VakArray(29) As Lijn

Private Sub Bereken_RijKolom(ByVal i As Integer, _
                            ByRef r As Integer, _
                            ByRef k As Integer)
Dim rr As Double
    rr = Fix(i / 30)
    r = rr
    k = i - r * 30
End Sub

Private Function IsBuur(p As Integer, q As Integer) As Boolean
Dim r1 As Integer, k1 As Integer
Dim r2 As Integer, k2 As Integer
If (p < 0) Or (p > 899) Or (q < 0) Or (q > 899) Then
    Isburen = False
    Exit Function
End If
Bereken_RijKolom p, r1, k1
Bereken_RijKolom q, r2, k2
If (r1 = r2) And (Abs(k1 - k2) = 1) Or _
   (k1 = k2) And (Abs(r1 - r2) = 1) Or _
    (Abs(r1 - r2) = 1) And (Abs(k1 - k2) = 1) Then
    IsBuur = True
Else
    IsBuur = False
End If

End Function


Private Function TelBuren(ByVal p As Integer) As Integer
Dim n As Integer
Dim i As Integer
i = p - 31
If IsBuur(p, i) Then
    If Vakje(i).BackColor = Kleur Then
        n = n + 1
    End If
End If
i = i + 1
If IsBuur(p, i) Then
    If Vakje(i).BackColor = Kleur Then
        n = n + 1
    End If
End If
i = i + 1
If IsBuur(p, i) Then
    If Vakje(i).BackColor = Kleur Then
        n = n + 1
    End If
End If
i = p - 1
If IsBuur(p, i) Then
    If Vakje(i).BackColor = Kleur Then
        n = n + 1
    End If
End If
i = p + 1
If IsBuur(p, i) Then
    If Vakje(i).BackColor = Kleur Then
        n = n + 1
    End If
End If
i = p + 29
If IsBuur(p, i) Then
    If Vakje(i).BackColor = Kleur Then
        n = n + 1
    End If
End If
i = i + 1
If IsBuur(p, i) Then
    If Vakje(i).BackColor = Kleur Then
        n = n + 1
    End If
End If
i = i + 1
If IsBuur(p, i) Then
    If Vakje(i).BackColor = Kleur Then
        n = n + 1
    End If
End If

TelBuren = n
End Function


Private Sub NieuweGeneratie()
Dim AantalBuren As Integer
Dim i As Integer, r As Integer, k As Integer

For i = 0 To AANTALVAKJES - 1
    AantalBuren = TelBuren(i)
    Bereken_RijKolom i, r, k
    If AantalBuren < 2 Or AantalBuren > 3 Then
        VakArray(r).vak(k) = 0
    ElseIf AantalBuren = 3 Then
        VakArray(r).vak(k) = 1
    End If
Next
For i = 0 To 899
    Bereken_RijKolom i, r, k
    If VakArray(r).vak(k) = 0 Then
        Vakje(i).BackColor = vbWhite
    Else
        Vakje(i).BackColor = Kleur
    End If
Next
End Sub

Private Sub InitializeRandom()
Dim i As Integer, Rij As Integer, Kolom As Integer
For i = 0 To AANTALVAKJES - 1
    Bereken_RijKolom i, Rij, Kolom
    VakArray(Rij).vak(Kolom) = Fix(Rnd * 2)
    If VakArray(Rij).vak(Kolom) = 0 Then
        Vakje(i).BackColor = vbWhite
    Else
        Vakje(i).BackColor = Kleur
    End If
   
Next
End Sub

Private Sub cmdClear_Click()
Dim i As Integer
Dim r As Integer, k As Integer
For i = 0 To 899
    Vakje(i).BackColor = vbWhite
    Bereken_RijKolom i, r, k
    VakArray(r).vak(k) = 0
Next
End Sub

Private Sub cmdRandom_Click()
    InitializeRandom
End Sub

Private Sub cmdStart_Click()
If Timer1.Enabled Then
    Timer1.Enabled = False
    cmdStart.Caption = "Start"
    cmdRandom.Enabled = True
Else
    Timer1.Enabled = True
    cmdStart.Caption = "Stop"
    cmdRandom.Enabled = False
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Rij As Integer
Dim Kolom As Integer

Randomize
Spelgestart = False
Kleur = vbRed
Timer1.Enabled = False
' speed is maximum
OptSpeed(0).Value = True
Timer1.Interval = 1
VakArray(0).vak(0) = Fix(Rnd * 2)

Vakje(0).Width = 120
Vakje(0).Height = 120
Top = Vakje(0).Top
Left = Vakje(0).Left
Rij = 0
For i = 1 To AANTALVAKJES - 1
    Load Vakje(i)
    If i Mod 30 = 0 Then
        Left = Vakje(0).Left
        Top = Top + Vakje(0).Height - 15
        Vakje(i).Left = Left
        Vakje(i).Top = Top
    Else
        Left = Left + Vakje(0).Width - 15
        Vakje(i).Left = Left
        Vakje(i).Top = Top
    End If
    Vakje(i).Visible = True
    

Next
InitializeRandom

End Sub

Private Sub Random_Click()

End Sub

Private Sub OptSpeed_Click(Index As Integer)
Select Case Index
    Case 0: Timer1.Interval = 1
    Case 1: Timer1.Interval = 200
    Case 2: Timer1.Interval = 1000
End Select
End Sub

Private Sub Timer1_Timer()
    NieuweGeneratie
    DoEvents
End Sub

Private Sub Vakje_Click(Index As Integer)
Dim r As Integer, k As Integer
        If Vakje(Index).BackColor = Kleur Then
            Vakje(Index).BackColor = vbWhite
            Bereken_RijKolom Index, r, k
            VakArray(r).vak(k) = 0
        Else
            Vakje(Index).BackColor = Kleur
            Bereken_RijKolom Index, r, k
            VakArray(r).vak(k) = Kleur

        End If
End Sub
