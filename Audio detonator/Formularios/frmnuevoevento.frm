VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmnuevoevento 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   "
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   Icon            =   "frmnuevoevento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcomentarios 
      Height          =   255
      Left            =   5640
      TabIndex        =   23
      Top             =   645
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "&comentarios:"
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
      MICON           =   "frmnuevoevento.frx":0CCA
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
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   3240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         ToolTipText     =   "Agregar Texto Aqui ..."
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FF00&
         Caption         =   "Domingos."
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000080FF&
         Caption         =   "Sabado ."
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Viernes ."
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Jueves ."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Miercoles."
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Martes ."
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Lunes ."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   2
         ItemData        =   "frmnuevoevento.frx":0CE6
         Left            =   840
         List            =   "frmnuevoevento.frx":0CF0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   1
         ItemData        =   "frmnuevoevento.frx":0D0B
         Left            =   840
         List            =   "frmnuevoevento.frx":0D0D
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1050
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         ItemData        =   "frmnuevoevento.frx":0D0F
         Left            =   600
         List            =   "frmnuevoevento.frx":0D19
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   295
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         Format          =   49676290
         UpDown          =   -1  'True
         CurrentDate     =   0.805555555555556
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   18
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   650
         Width           =   600
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
   End
   Begin Virtual_Martin_temporize.ChameleonBtn boton 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   1095
      _ExtentX        =   2566
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
      MICON           =   "frmnuevoevento.frx":0D2E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   20
      Top             =   0
      Width           =   0
   End
   Begin Virtual_Martin_temporize.ChameleonBtn boton 
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   21
      Top             =   3240
      Width           =   1095
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Crear"
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
      MICON           =   "frmnuevoevento.frx":0D4A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdfunct 
      Height          =   255
      Left            =   4920
      TabIndex        =   22
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "&Funciones al sistema"
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
      MICON           =   "frmnuevoevento.frx":0D66
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   2520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label labinfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar Nuevo Evento"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1950
   End
   Begin VB.Menu comentar 
      Caption         =   "menú"
      Visible         =   0   'False
      Begin VB.Menu mc 
         Caption         =   "Salida de Clase"
         Index           =   0
      End
      Begin VB.Menu mc 
         Caption         =   "Entrada a clase"
         Index           =   1
      End
      Begin VB.Menu mc 
         Caption         =   "Salida al Patio"
         Index           =   2
      End
      Begin VB.Menu mc 
         Caption         =   "Salida al Recreo"
         Index           =   3
      End
      Begin VB.Menu mc 
         Caption         =   "Entrada de la Directora"
         Index           =   4
      End
      Begin VB.Menu mc 
         Caption         =   "Entrada personal docente"
         Index           =   5
      End
      Begin VB.Menu mc 
         Caption         =   "Entrada  Personal de limpieza y aseado "
         Index           =   6
      End
      Begin VB.Menu mc 
         Caption         =   "Evento especial:"
         Index           =   7
      End
      Begin VB.Menu mc 
         Caption         =   "Evento Importante:"
         Index           =   8
      End
      Begin VB.Menu mc 
         Caption         =   "Evento sin Importancia:"
         Index           =   9
      End
      Begin VB.Menu mc 
         Caption         =   "Recordatorio de estudio"
         Index           =   10
      End
      Begin VB.Menu mc 
         Caption         =   "Recordatorio descanso"
         Index           =   11
      End
      Begin VB.Menu mc 
         Caption         =   "Recordatorio por enfermedad"
         Index           =   12
      End
      Begin VB.Menu mc 
         Caption         =   "Personal dejo de asistir"
         Index           =   13
      End
      Begin VB.Menu mc 
         Caption         =   "Personal  renuncio"
         Index           =   14
      End
      Begin VB.Menu mc 
         Caption         =   "Error S.O"
         Index           =   15
      End
      Begin VB.Menu mc 
         Caption         =   "Mensaje x en:"
         Index           =   16
      End
      Begin VB.Menu mc 
         Caption         =   "Retorno Windows 95:"
         Index           =   17
      End
      Begin VB.Menu mc 
         Caption         =   "Retorno Windows 98:"
         Index           =   18
      End
      Begin VB.Menu mc 
         Caption         =   "Retorno Windows 98 me:"
         Index           =   19
      End
      Begin VB.Menu mc 
         Caption         =   "Retorno Windows 2000:"
         Index           =   20
      End
      Begin VB.Menu mc 
         Caption         =   "Retorno Windows 2000 me:"
         Index           =   21
      End
      Begin VB.Menu mc 
         Caption         =   "Retorno Windows  XP:"
         Index           =   22
      End
      Begin VB.Menu mc 
         Caption         =   "Retorno Windows  Vista:"
         Index           =   23
      End
      Begin VB.Menu mc 
         Caption         =   "Retorno Windows  7:"
         Index           =   24
      End
      Begin VB.Menu mc 
         Caption         =   "Evento reunión:"
         Index           =   25
      End
      Begin VB.Menu mc 
         Caption         =   "Evento solo Aviso:"
         Index           =   26
      End
      Begin VB.Menu mc 
         Caption         =   "Evento acudir al programador:"
         Index           =   27
      End
      Begin VB.Menu mc 
         Caption         =   "Sin Evento."
         Index           =   28
      End
      Begin VB.Menu mc 
         Caption         =   "Tiempo de entrada:"
         Index           =   29
      End
      Begin VB.Menu mc 
         Caption         =   "Tiempo de Salida:"
         Index           =   30
      End
      Begin VB.Menu mc 
         Caption         =   "Salida del Turno Matutino."
         Index           =   31
      End
      Begin VB.Menu mc 
         Caption         =   "Salida del Turno Nocturno."
         Index           =   32
      End
      Begin VB.Menu mc 
         Caption         =   "Salida del Turno Despretino."
         Index           =   33
      End
      Begin VB.Menu mc 
         Caption         =   "Entrada al Turno Matutino."
         Index           =   34
      End
      Begin VB.Menu mc 
         Caption         =   "Entrada al Turno Nocturno."
         Index           =   35
      End
      Begin VB.Menu mc 
         Caption         =   "Entrada al Turno Despretino."
         Index           =   36
      End
      Begin VB.Menu mc 
         Caption         =   "Tiempo permanece activado el timbre:"
         Index           =   37
      End
      Begin VB.Menu mc 
         Caption         =   "________________________________"
         Index           =   38
      End
   End
End
Attribute VB_Name = "frmnuevoevento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Nuevo Evento y Modificación para Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Dim nuevoEvento As evento

Private Sub boton_Click(Index As Integer)
 frmprograma.Enabled = True
 With frmprograma
 Select Case Index
  Case (0)
  Unload Me
  Case (1)
 If boton(1).Caption = "&Crear" Then
 nuevo_evento_de_dias
 Crear ' crea un unevo evento de timbre
 sistema.ingresarDatos
 ElseIf boton(1).Caption = "&Modificar" Then
 'selección
 Select Case MsgBox("Quieres Aplicar las Modificaciones del Evento." _
 , vbYesNo + vbInformation, "Opciones")
  Case (vbYes)
  Me.Caption = "Modificar Evento ."
  labinfo.Caption = "Modificar Evento ."
  .listado(0).List(.listado(0).ListIndex) = DTPicker1.Value
  .listado(1).List(.listado(1).ListIndex) = Combo1(0).Text
  .listado(2).List(.listado(2).ListIndex) = Combo1(1).Text
  .listado(3).List(.listado(3).ListIndex) = Text1.Text
  'dias Set'
 set_dias ' cambia los dias de la semana
 .Filtro.List(.Filtro.ListIndex) = Combo1(2).ListIndex
 Unload Me
 End Select
 End If
 End Select
 End With
End Sub

Private Sub Crear()
 Set nuevoEvento = New evento
 With nuevoEvento
 .vHora.Add DTPicker1.Value
 .vTipo.Add Combo1(0).Text
 .vIntervalo.Add Combo1(1).Text
 .vtipod.Add Combo1(2).Text
 .vDescripcion.Add Text1.Text
 End With
 With frmprograma
 Dim recor As Integer
 For recor = 1 To nuevoEvento.vHora.Count
 .listado(0).AddItem nuevoEvento.vHora(recor)
 .listado(1).AddItem nuevoEvento.vTipo(recor)
 .listado(2).AddItem nuevoEvento.vIntervalo(recor)
 .listado(3).AddItem nuevoEvento.vDescripcion(recor)
 Next
 End With
 Unload Me
End Sub

Private Sub boton_KeyPress(Index As Integer, KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdobsiones_Click()
 PopupMenu obsiones
End Sub

Private Sub cmdobsiones_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdcomentarios_Click()
 PopupMenu comentar
End Sub

Private Sub cmdfunct_Click()
 If boton(1).Caption = "&Modificar" Then
 frmfunciones.cmdAplicar.Caption = "&Modificar"
 End If
 frmfunciones.Show 1
End Sub

Private Sub cmdfunct_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Combo1_Click(Index As Integer)
 Select Case Index
  Case (2)
 Select Case Combo1(2).ListIndex
  Case (0)
  visiblex False
  activado 0
  Dim td As Byte
  For td = 0 To 6
  Check1(CInt(td)).Value = 1
  Next
 Case (1)
 visiblex True
 activado 0
 End Select
 End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 Combo1(2).ListIndex = CInt(MemoriaF.numero)
 visiblex CInt(MemoriaF.numero)
 DTPicker1.Value = Time
 agregar_elementos
 If MemoriaF.dias = True Then
 devolver_dias
 End If
End Sub

Private Sub agregar_elementos()
 Dim numero As Integer
 Combo1(0).ListIndex = 0
 For numero = 1 To 777
 Combo1(1).AddItem (numero)
 Next
 Combo1(1).ListIndex = 4
End Sub

Private Sub visiblex(ByVal visilblex As Boolean)
 Dim rx As Integer
 For rx = 0 To 6
 Check1(rx).Enabled = visilblex
 Next
End Sub

Private Sub activado(ByVal activado As Byte)
 Dim rx As Integer
 For rx = 0 To 6
 Check1(rx).Value = activado
 Next
End Sub

Private Sub almanaque_Click()
 frmalmanaque.Show 1
End Sub

Private Sub nuevo_evento_de_dias()
 Const nulo As String = "0"      'nulo
 Const lunes As String = "2"     'lunes
 Const martes As String = "3"    'martes
 Const miercoles As String = "4" 'miercoles
 Const jueves As String = "5"    'jueves
 Const viernes As String = "6"   'viernes
 Const sabado As String = "7"    'sabado
 Const domingo As String = "1"   'domingo
 With frmprograma
 Select Case Check1(0).Value     ' Lunes
  Case (1)
  .lunes(0).AddItem lunes
  Case (0)
  .lunes(0).AddItem nulo
 End Select
Select Case Check1(1).Value      ' Martes
 Case (1)
 .martes.AddItem martes          ' Martes
 Case (0)
 .martes.AddItem nulo
End Select
Select Case Check1(2).Value ' Miercoles
 Case (1)
 .miercoles.AddItem miercoles
 Case (0)
 .miercoles.AddItem nulo
End Select
Select Case Check1(3).Value ' Jueves
 Case (1)
 .jueves.AddItem jueves
 Case (0)
 .jueves.AddItem nulo
End Select
Select Case Check1(4).Value ' Viernes
 Case (1)
 .viernes.AddItem viernes
 Case (0)
 .viernes.AddItem nulo
End Select
Select Case Check1(5).Value ' Sabado
 Case (1)
 .sabado.AddItem sabado
 Case (0)
 .sabado.AddItem nulo
End Select
Select Case Check1(6).Value ' Domingo
 Case (1)
 .domingo.AddItem domingo
 Case (0)
 .domingo.AddItem nulo
End Select
'***************'> Asignacion de Filtro <******************'
.Filtro.AddItem Combo1(2).ListIndex
 End With
End Sub

Public Sub devolver_dias()
 Dim dev As Integer
 For dev = 0 To frmprograma.listado(0).ListCount
 With frmprograma
 'lunes
 Select Case .lunes(0).List(.lunes(0).ListIndex)
  Case (2)
  Check1(0).Value = 1
  Case (0)
  Check1(0).Value = 0
 End Select
'martes
Select Case .martes.List(.martes.ListIndex)
 Case (3)
 Check1(1).Value = 1
 Case (0)
 Check1(1).Value = 0
End Select
'miercoles
Select Case .miercoles.List(.miercoles.ListIndex)
 Case (4)
 Check1(2).Value = 1
 Case (0)
 Check1(2).Value = 0
End Select
'jueves
Select Case .jueves.List(.jueves.ListIndex)
 Case (5)
 Check1(3).Value = 1
 Case (0)
 Check1(3).Value = 0
End Select
'viernes
Select Case .viernes.List(.viernes.ListIndex)
 Case (6)
 Check1(4).Value = 1
 Case (0)
 Check1(4).Value = 0
End Select
'sabado
Select Case .sabado.List(.sabado.ListIndex)
 Case (7)
 Check1(5).Value = 1
 Case (0)
 Check1(5).Value = 0
End Select
'domingo
Select Case .domingo.List(.domingo.ListIndex)
 Case (1)
 Check1(6).Value = 1
 Case (0)
 Check1(6).Value = 0
End Select
End With
Next dev
End Sub

Private Sub set_dias()
 With frmprograma
 'lunes
 Select Case Check1(0).Value
  Case (1)
  .lunes(0).List(.lunes(0).ListIndex) = 2
  Case (0)
  .lunes(0).List(.lunes(0).ListIndex) = 0
 End Select
 'martes
 Select Case Check1(1).Value
  Case (1)
  .martes.List(.martes.ListIndex) = 3
  Case (0)
  .martes.List(.martes.ListIndex) = 0
  End Select
 'miercoles
Select Case Check1(2).Value
 Case (1)
 .miercoles.List(.miercoles.ListIndex) = 4
 Case (0)
 .miercoles.List(.miercoles.ListIndex) = 0
End Select
'jueves
Select Case Check1(3).Value
 Case (1)
 .jueves.List(.jueves.ListIndex) = 5
 Case (0)
 .jueves.List(.jueves.ListIndex) = 0
End Select
'viernes
Select Case Check1(4).Value
 Case (1)
 .viernes.List(.viernes.ListIndex) = 6
 Case (0)
 .viernes.List(.viernes.ListIndex) = 0
End Select
'sabado
Select Case Check1(5).Value
 Case (1)
 .sabado.List(.sabado.ListIndex) = 7
 Case (0)
 .sabado.List(.sabado.ListIndex) = 0
End Select
'domingo
Select Case Check1(6).Value
 Case (1)
 .domingo.List(.domingo.ListIndex) = 1
 Case (0)
 .domingo.List(.domingo.ListIndex) = 0
End Select
End With
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmnuevoevento
End Sub

Private Sub mc_Click(Index As Integer)
 Text1.Text = Text1.Text + mc.Item(Index).Caption + vbNewLine
End Sub
