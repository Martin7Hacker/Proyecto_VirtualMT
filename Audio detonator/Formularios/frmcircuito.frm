VERSION 5.00
Begin VB.Form frmcircuito 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circuito Electrónico"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   ControlBox      =   0   'False
   Icon            =   "frmcircuito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdImprimir 
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   5640
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      btype           =   3
      tx              =   "&Imprimir"
      enab            =   -1  'True
      font            =   "frmcircuito.frx":0CCA
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   0
      bcolo           =   0
      fcol            =   14737632
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmcircuito.frx":0CF6
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdaceptar 
      Height          =   375
      Left            =   6960
      TabIndex        =   22
      Top             =   5640
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      btype           =   3
      tx              =   "&Aceptar"
      enab            =   -1  'True
      font            =   "frmcircuito.frx":0D14
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   0
      bcolo           =   0
      fcol            =   14737632
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmcircuito.frx":0D40
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   -120
      Picture         =   "frmcircuito.frx":0D5E
      ScaleHeight     =   420
      ScaleWidth      =   8130
      TabIndex        =   27
      Top             =   0
      Width           =   8160
      Begin VB.Label labtitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Esqemas del Circuito."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   5175
      End
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdSigiente 
      Height          =   375
      Left            =   1200
      TabIndex        =   26
      Top             =   5400
      Width           =   255
      _extentx        =   450
      _extenty        =   661
      btype           =   3
      tx              =   "4"
      enab            =   -1  'True
      font            =   "frmcircuito.frx":FE20
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   0
      bcolo           =   0
      fcol            =   14737632
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmcircuito.frx":FE48
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdDondeQuedo 
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   5400
      Width           =   495
      _extentx        =   873
      _extenty        =   661
      btype           =   3
      tx              =   "a"
      enab            =   -1  'True
      font            =   "frmcircuito.frx":FE66
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   0
      bcolo           =   0
      fcol            =   14737632
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmcircuito.frx":FE8E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAtras 
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   5400
      Width           =   255
      _extentx        =   450
      _extenty        =   661
      btype           =   3
      tx              =   "3"
      enab            =   -1  'True
      font            =   "frmcircuito.frx":FEAC
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   0
      bcolo           =   0
      fcol            =   14737632
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmcircuito.frx":FED4
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.PictureBox Picg 
      Height          =   255
      Index           =   4
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picg 
      Height          =   255
      Index           =   3
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picg 
      Height          =   255
      Index           =   2
      Left            =   360
      Picture         =   "frmcircuito.frx":FEF2
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picg 
      Height          =   255
      Index           =   1
      Left            =   360
      Picture         =   "frmcircuito.frx":763F0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picg 
      Height          =   255
      Index           =   0
      Left            =   360
      Picture         =   "frmcircuito.frx":DC8EE
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picb 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   4
      Left            =   1800
      ScaleHeight     =   5295
      ScaleWidth      =   7335
      TabIndex        =   10
      Top             =   600
      Width           =   7335
      Begin VB.PictureBox picd 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   4695
         Index           =   4
         Left            =   0
         ScaleHeight     =   4695
         ScaleWidth      =   6135
         TabIndex        =   11
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.PictureBox picb 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   1
      Left            =   240
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
      Begin VB.PictureBox picd 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   1
         Left            =   100
         ScaleHeight     =   855
         ScaleWidth      =   975
         TabIndex        =   5
         Top             =   120
         Width           =   975
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   1
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox picd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   0
      Left            =   360
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   3
      Top             =   720
      Width           =   975
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.PictureBox picb 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   0
      Left            =   240
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox piccirc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5505
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   405
      Width           =   1665
      Begin VB.PictureBox picb 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   240
         ScaleHeight     =   1095
         ScaleWidth      =   1215
         TabIndex        =   8
         Top             =   3840
         Width           =   1215
         Begin VB.PictureBox picd 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   3
            Left            =   100
            ScaleHeight     =   855
            ScaleWidth      =   975
            TabIndex        =   9
            Top             =   120
            Width           =   975
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "4)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   3
               Left            =   360
               TabIndex        =   15
               Top             =   240
               Width           =   255
            End
         End
      End
      Begin VB.PictureBox picd 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   2
         Left            =   370
         ScaleHeight     =   855
         ScaleWidth      =   975
         TabIndex        =   6
         Top             =   2760
         Width           =   975
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   2
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox picb 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   240
         ScaleHeight     =   1095
         ScaleWidth      =   1215
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   -120
      ScaleHeight     =   5895
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmcircuito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Diagrama del Circuito Electrónico Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Dim fotoacual, recfoto As Byte: Private Const azulado = &HC000C0
Private Const winnor_ = &H404040


Private Sub cmdAceptar_Click()
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub cmdAceptar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdAtras_Click()
On Error GoTo no_se
 If recfoto <> -1 Then
 selecionar recfoto
 recfoto = recfoto - 1
 End If
no_se:
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdSigiente_Click()
 If recfoto < 4 Then
  selecionar recfoto
  recfoto = recfoto + 1
 End If
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdDondeQuedo_Click()
 selecionar fotoacual
End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdImprimir_Click()
 On Error GoTo no_se
    Printer.Print
    Printer.PaintPicture picd(4).Picture, 0, 0, picd(4).Width, picd(4).Height
    Printer.EndDoc
no_se:
End Sub

Private Sub Command4_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 recfoto = 0
 cargar_lenguage ' cargar lenguage
End Sub

Private Sub Label1_Click(Index As Integer)
 picd_Click (Index)
End Sub

Private Sub picb_KeyPress(Index As Integer, KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub piccirc_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub picd_Click(Index As Integer)
 fotoacual = Index
  Select Case Index
        Case (0)
            picb(0).BackColor = vbRed
        Case (1)
            picb(1).BackColor = vbRed
        Case (2)
            picb(2).BackColor = vbRed
        Case (3)
            picb(3).BackColor = vbRed
 End Select
 selecionar fotoacual
End Sub

Private Sub selecionar(ByVal control As Byte)
 Dim c As Byte
              For c = 0 To 3
                  If picb(control).BackColor = azulado Then
                  Exit For
                  Else
                  picb(c).BackColor = winnor_
                  picd(4).Picture = Nothing
                  End If
              Next
              
  Select Case control

       Case (0)
            picb(0).BackColor = azulado
            picd(4).Picture = Picg(0).Picture
       Case (1)
            picb(1).BackColor = azulado
            picd(4).Picture = Picg(1).Picture
       Case (2)
            picb(2).BackColor = azulado
            picd(4).Picture = Picg(2).Picture
       Case (3)
            picb(3).BackColor = azulado
            picd(4).Picture = Picg(3).Picture
 End Select
End Sub

Private Sub picd_KeyPress(Index As Integer, KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub picd_MouseMove(Index As Integer, Button As Integer, _
Shift As Integer, X As Single, Y As Single)
 selecionar Index
End Sub

Private Sub salir_op(ByVal dig As Byte)
'sale oprimendo Esc
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmcircuito
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Picture2_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Picture3_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguage_circuito(0)
 labtitulo.Caption = Lenguage.lenguage_circuito(1)
 cmdImprimir.Caption = Lenguage.lenguage_circuito(2)
 cmdaceptar.Caption = Lenguage.lenguage_circuito(3)
End Sub
