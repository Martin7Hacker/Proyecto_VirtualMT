VERSION 5.00
Begin VB.Form frmpuerto 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puerto de Salida"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   ClipControls    =   0   'False
   Icon            =   "frmpuerto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox pin8 
      BackColor       =   &H000000FF&
      Caption         =   "Pin 8"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.CheckBox pin7 
      BackColor       =   &H000000FF&
      Caption         =   "Pin 7"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox pin6 
      BackColor       =   &H000000FF&
      Caption         =   "Pin 6"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox pin5 
      BackColor       =   &H000000FF&
      Caption         =   "Pin 5"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox pin4 
      BackColor       =   &H000000FF&
      Caption         =   "Pin 4"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox pin3 
      BackColor       =   &H000000FF&
      Caption         =   "Pin 3"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox pin2 
      BackColor       =   &H000000FF&
      Caption         =   "Pin 2"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox pin1 
      BackColor       =   &H000000FF&
      Caption         =   "Pin 1"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "&Salida 5v"
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   1200
      TabIndex        =   9
      ToolTipText     =   $"frmpuerto.frx":0CCA
      Top             =   480
      Width           =   3375
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmpuerto.frx":0D51
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdnormal 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&normal"
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
      MICON           =   "frmpuerto.frx":0D6D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdsalir 
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aceptar"
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
      MICON           =   "frmpuerto.frx":0D89
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Labelbuerto 
      BackStyle       =   0  'Transparent
      Caption         =   "Usted Tiene que tener conocimiento antes de realizar algun cambio Aquí."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5205
   End
End
Attribute VB_Name = "frmpuerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Conexión por Puerto Paralelo de Virtual Martin temporize v1.0
'*
'*
'***************************************************************************

Private Sub cmdcancelar_Click()
 cerrar
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdnormal_Click()
 pin1.Value = 1
 pin2.Value = 0
 pin3.Value = 1
 pin4.Value = 0
 pin5.Value = 0
 pin6.Value = 0
 pin7.Value = 0
 pin8.Value = 0
 almacenar_datos 'llamada al procedimiento
End Sub

Private Sub cmdnormal_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsalir_Click()
 cerrar
End Sub

Private Sub cerrar()
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub almacenar_datos()
 Mopuerto.pu1 = pin1.Value
 Mopuerto.pu2 = pin2.Value
 Mopuerto.pu3 = pin3.Value
 Mopuerto.pu4 = pin4.Value
 Mopuerto.pu5 = pin5.Value
 Mopuerto.pu6 = pin6.Value
 Mopuerto.pu7 = pin7.Value
 Mopuerto.pu8 = pin8.Value
End Sub

Private Sub cargar_datos()
 pin1.Value = Mopuerto.pu1
 pin2.Value = Mopuerto.pu2
 pin3.Value = Mopuerto.pu3
 pin4.Value = Mopuerto.pu4
 pin5.Value = Mopuerto.pu5
 pin6.Value = Mopuerto.pu6
 pin7.Value = Mopuerto.pu7
 pin8.Value = Mopuerto.pu8
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 cargar_datos
 Me.Icon = frmprograma.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
 almacenar_datos
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmpuerto
End Sub
