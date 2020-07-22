VERSION 5.00
Begin VB.Form frmimpresor 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresor por cantidad"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   Icon            =   "frmimpresor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdmas 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmimpresor.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdmenos 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmimpresor.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtcop 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.PictureBox Picture2 
         Height          =   855
         Left            =   720
         ScaleHeight     =   795
         ScaleWidth      =   0
         TabIndex        =   2
         Top             =   180
         Width           =   60
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copias:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   255
         Width           =   645
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "frmimpresor.frx":0D02
         Top             =   310
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmimpresor.frx":19CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdaceptar 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Mandar a Imprimir las Copias"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmimpresor.frx":19E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmimpresor.frx":1A04
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "frmimpresor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Poder Imprimir para Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Dim copias As Long

Private Sub cmdAceptar_Click()
 Dim ip As Long
 For ip = 1 To copias
 ModImprimir.Imprimir_ListView
 Next
 Unload Me
End Sub

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmdmas_Click()
 copias = copias + 1
 txtcop.Text = copias
 cmdmenos.Enabled = True
End Sub

Private Sub cmdmenos_Click()
 If copias = 1 Then
 cmdmenos.Enabled = False
 Else
 copias = copias - 1
 End If
 txtcop.Text = copias
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 copias = 1
 txtcop.Text = copias
 cargar_lenguage 'cargar lenguage
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguage_estcopias(0)
 Label1.Caption = Lenguage.lenguage_estcopias(1)
 cmdmenos.Caption = Lenguage.lenguage_estcopias(2)
 cmdmas.Caption = Lenguage.lenguage_estcopias(3)
 cmdcancelar.Caption = Lenguage.lenguage_estcopias(4)
 cmdAceptar.Caption = Lenguage.lenguage_estcopias(5)
End Sub
