VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmmensage 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   Icon            =   "frmmensage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   255
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdsalir 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      MICON           =   "frmmensage.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdguardarysalir 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Guardar y Salir"
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
      MICON           =   "frmmensage.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcancelar 
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   3120
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
      MICON           =   "frmmensage.frx":0D02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿Existen Archivos en memoria que desea Hacer?"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   4170
   End
End
Attribute VB_Name = "frmmensage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Existen Archivos para Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Dim d As Long

Private Sub cmdcancelar_Click()
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdguardarysalir_Click()
 frmprograma.guardard_Click
 End
End Sub

Private Sub cmdguardarysalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsalir_Click()
 End
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cargar_datos
 cargar_lenguage
End Sub

Private Sub cargar_datos()
 Const espacio As String = "                               "
  On Error GoTo no_se
  With frmprograma
  Dim ah As Integer
  Dim v As String
  Dim et As ListItem
  With ListView1.ColumnHeaders
  .Add , , Lenguage.lenguage_memoria(2)
  .Add , , Lenguage.lenguage_memoria(3)
  .Add , , Lenguage.lenguage_memoria(4)
  .Add , , Lenguage.lenguage_memoria(5)
  .Add , , Lenguage.lenguage_memoria(6)
  End With
  With ListView1
  ' Las pruebas serán en modo "detalle"
  .View = lvwReport
  .LabelEdit = lvwManual
  ' Permitir múltiple selección
  .MultiSelect = False
  ' Para que al perder el foco,
  ' se siga viendo el que está seleccionado
  .HideSelection = False
  End With
  ListView1.View = lvwReport
  For ah = 0 To .listado(0).ListCount - 1
  If .listado(1).List(ah) = "Salida" Then
  v = "   "
  Else
  v = ""
  End If
  d = Int(ah) + 1
  With ListView1.ListItems.Add(, , "Evento_____ " & d)
  .SubItems(1) = frmprograma.listado(0).List(ah)
  .SubItems(2) = frmprograma.listado(1).List(ah)
  .SubItems(3) = frmprograma.listado(2).List(ah)
  .SubItems(4) = frmprograma.listado(3).List(ah)
  End With
  Next ah
End With
no_se:
End Sub
Private Sub salir_op(ByVal dig As Byte)
'sale oprimendo Esc
fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmmensage
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cargar_lenguage()
 Me.Caption = Lenguage.lenguage_memoria(0)
 labdatos.Caption = Lenguage.lenguage_memoria(1)
 cmdsalir.Caption = Lenguage.lenguage_memoria(7)
 cmdguardarysalir.Caption = Lenguage.lenguage_memoria(8)
 cmdcancelar.Caption = Lenguage.lenguage_memoria(9)
End Sub
