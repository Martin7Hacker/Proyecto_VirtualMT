VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmfunciones 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "funciones al sistema"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "frmfunciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8955
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cob1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8655
      Begin VB.Frame frame2 
         BackColor       =   &H00000000&
         Height          =   2055
         Left            =   4200
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   4335
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   420
            Left            =   2040
            TabIndex        =   5
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   741
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   0
            CustomFormat    =   "m"
            Format          =   49479683
            UpDown          =   -1  'True
            CurrentDate     =   0.805555555555556
         End
         Begin VB.TextBox txtd 
            BackColor       =   &H00000000&
            ForeColor       =   &H000000FF&
            Height          =   1815
            Left            =   120
            MaxLength       =   127
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label labinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Sin dialogo..."
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1680
            TabIndex        =   7
            Top             =   840
            Width           =   1785
         End
         Begin VB.Label lbld 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo ="
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
            Left            =   1200
            TabIndex        =   6
            Top             =   960
            Width           =   795
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   2535
         Left            =   3960
         ScaleHeight     =   2475
         ScaleWidth      =   0
         TabIndex        =   2
         Top             =   120
         Width           =   60
      End
      Begin VB.Image Image1 
         Height          =   2295
         Left            =   240
         Picture         =   "frmfunciones.frx":0CCA
         Top             =   240
         Width           =   3525
      End
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmfunciones.frx":1B430
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAplicar 
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aplicar"
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
      MICON           =   "frmfunciones.frx":1B44C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcomentarios 
      Height          =   255
      Left            =   7440
      TabIndex        =   11
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "comentarios:"
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
      MICON           =   "frmfunciones.frx":1B468
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
      BackStyle       =   0  'Transparent
      Caption         =   "Funciones de Sistemas Operables."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4005
   End
   Begin VB.Menu comentarios 
      Caption         =   "comentarios"
      Visible         =   0   'False
      Begin VB.Menu f0 
         Caption         =   "Finalizo el turno Apagar PC"
         Index           =   0
      End
      Begin VB.Menu f0 
         Caption         =   "Apagar pc error en el sistema Operativo"
         Index           =   1
      End
      Begin VB.Menu f0 
         Caption         =   "Reniciar el PC"
         Index           =   2
      End
      Begin VB.Menu f0 
         Caption         =   "Forzar el Cierre de Aplicaciones y Apagar"
         Index           =   3
      End
      Begin VB.Menu f0 
         Caption         =   "Apagar finalizo la tarea"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmfunciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Funciones para Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Public devolver_comando As String

Private Sub cargar_controles()
 With cob1
 .AddItem Lenguage.lenguage_estFunciones(3)
 .AddItem Lenguage.lenguage_estFunciones(4)
 .AddItem Lenguage.lenguage_estFunciones(5)
 .AddItem Lenguage.lenguage_estFunciones(6)
 .AddItem Lenguage.lenguage_estFunciones(7)
 .AddItem Lenguage.lenguage_estFunciones(8)
 .AddItem Lenguage.lenguage_estFunciones(9)
 .AddItem Lenguage.lenguage_estFunciones(10)
 End With
End Sub

Private Sub cmdAplicar_Click()
devolverString: sistema.tomarDatos
 frmnuevoevento.Text1.Text = txtd.Text & " tiempo :" & DTPicker1.Minute
 frmnuevoevento.Combo1(1).Text = DTPicker1.Minute
 If cmdAplicar.Caption = "&Modificar" Then
 sistema.modificarDatos 'modifica los datos ingresado
 End If
 If cob1.Text = "" Then
 MsgBox "Ingrese una Opción del Sistema"
 End If
Unload Me
End Sub

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmdcomentarios_Click()
 PopupMenu comentarios
End Sub

Private Sub cmdcomentarios_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsel_KeyPress(Index As Integer, KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cob1_Click()
 If cob1.ListIndex = 5 Then
 labinfo.Visible = False
 frame2.Visible = True
 txtd.Visible = False
 lbld.Visible = True
 DTPicker1.Visible = True
 ElseIf cob1.ListIndex = 6 Then
 labinfo.Visible = False
 frame2.Visible = True
 txtd.Visible = True
 DTPicker1.Visible = False
 lbld.Visible = False
 Else
 labinfo.Visible = True
 frame2.Visible = True
 txtd.Visible = False
 DTPicker1.Visible = False
 lbld.Visible = False
 End If
End Sub

Private Sub devolverString()
 On Error GoTo no_se
 Select Case cob1.ListIndex
  Case (0)
  devolver_comando = ""   '// sin opcion
  Case (1)
  devolver_comando = ""   '// no hacer nada
  Case (2)
  devolver_comando = "so.dll -s -f" '// Apagar el equipo
  Case (3)
  devolver_comando = "so.dll -r"    '// reiniciar el equipo
  Case (4)
  devolver_comando = "so.dll -a"    '// anular el apagado del equipo
  Case (5)
  devolver_comando = "so.dll -m"    '// equipo que se / apagara / reiniciara / anulara
  Case (6)
  devolver_comando = "so.dll -t"    '// establecer el tiempo de cierre de apagado en xx segundos
  Case (7)
  devolver_comando = "so.dll -c"    '// le puedes aplicar comentarios
  Case (9)
  devolver_comando = "so.dll -f"    '// fuerza el cierre de aplicaciones sin advertir
 End Select
no_se:
End Sub

Private Sub cob1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub f0_Click(Index As Integer)
 txtd.Text = f0.Item(Index).Caption
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmfunciones
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cargar_controles
 On Error GoTo no_se
 If txtd.Text = "" Then
 txtd.Text = sistema.comentario
 End If
no_se:
 cargar_lenguage ' cargar lenguage
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub txtd_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguage_estFunciones(0)
 Label1.Caption = Lenguage.lenguage_estFunciones(1)
 cmdcomentarios.Caption = Lenguage.lenguage_estFunciones(2)
 cmdcancelar.Caption = Lenguage.lenguage_estFunciones(15)
 cmdAplicar.Caption = Lenguage.lenguage_estFunciones(16)
 labinfo.Caption = Lenguage.lenguage_estFunciones(17)
 lbld.Caption = Lenguage.lenguage_opciones_generador(18)
End Sub
