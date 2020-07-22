VERSION 5.00
Begin VB.Form frmarrancarconwindows 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Martin temporize: Iniciar con Windows"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   Icon            =   "frmarrancarconwindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   -600
      Picture         =   "frmarrancarconwindows.frx":0CCA
      ScaleHeight     =   420
      ScaleWidth      =   8130
      TabIndex        =   5
      Top             =   0
      Width           =   8160
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Iniciar con Windows Automaticamente"
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
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   5175
      End
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdArranciar 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Salir"
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
      MICON           =   "frmarrancarconwindows.frx":FD8C
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      Begin Virtual_Martin_temporize.ChameleonBtn cmdaplicar 
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Si"
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
         MICON           =   "frmarrancarconwindows.frx":FDA8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Virtual_Martin_temporize.ChameleonBtn cmdnoarrancar 
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&No"
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
         MICON           =   "frmarrancarconwindows.frx":FDC4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿ Iniciar programa con el Sistema Operativo Windows ?"
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
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4725
      End
   End
End
Attribute VB_Name = "frmarrancarconwindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Iniciar con Windows para el programa Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Option Explicit
'Constantes de la Rama del registro para los path de _
 las aplicaciones que inician con Windows
Const RAMA_RUN_WINDOWS As String = "SOFTWARE\Microsoft\" & _
                                   "Windows\CurrentVersion\Run"
Private Sub cmdAplicar_Click()
frmprograma.Enabled = True
Unload Me
End Sub

Private Sub cmdaplicar_KeyPress(KeyAscii As Integer)
salir_op KeyAscii
End Sub

Private Sub cmdArranciar_Click()
 Dim Path_Programa, _
 Titulo_Programa As String
 Dim Ret As Boolean
  On Error GoTo nose
    Path_Programa = App.Path & "\" & App.EXEName & ".exe"
    Titulo_Programa = App.Title
    Ret = EstablecerValor(HKEY_LOCAL_MACHINE1, _
                    RAMA_RUN_WINDOWS, _
                    Titulo_Programa, _
                    Path_Programa, REG_SZ1)
'si retorna True es por que creó el dato correctamente
    If Ret Then
       MsgBox Lenguage.lenguage_iniciarwindows(5), vbInformation
    Else
       MsgBox Lenguage.lenguage_iniciarwindows(6), vbCritical
    End If
nose:
End Sub

Private Sub cmdArranciar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdnoarrancar_Click()
 Dim Titulo_Programa As String
 Titulo_Programa = App.Title
 Call EliminarValor(HKEY_LOCAL_MACHINE, _
                   RAMA_RUN_WINDOWS, _
                   Titulo_Programa)
                   MsgBox Lenguage.lenguage_iniciarwindows(7), vbInformation
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmarrancarconwindows
End Sub

Private Sub cmdnoarrancar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsalir_Click()
 Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cargar_lenguage ' carga el lenguage del programa
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguage_iniciarwindows(0)
 Label1.Caption = Lenguage.lenguage_iniciarwindows(1)
 cmdArranciar.Caption = Lenguage.lenguage_iniciarwindows(2)
 cmdnoarrancar.Caption = Lenguage.lenguage_iniciarwindows(3)
 cmdaplicar.Caption = Lenguage.lenguage_iniciarwindows(4)
End Sub
