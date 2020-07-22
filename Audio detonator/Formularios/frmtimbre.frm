VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmtimbre 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evento en ejecución"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   Icon            =   "frmtimbre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tiempo 
      Interval        =   1000
      Left            =   600
      Top             =   5640
   End
   Begin VB.Timer timreloj 
      Interval        =   1000
      Left            =   120
      Top             =   5640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Comentarios"
      ForeColor       =   &H000000FF&
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8295
      Begin VB.Frame frmsolo_hora 
         BackColor       =   &H00000000&
         Height          =   735
         Left            =   6360
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label labhora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Solo Hora."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   195
            Left            =   150
            TabIndex        =   20
            ToolTipText     =   "Se ejecuta siempre sin importar el Día de la Semana"
            Top             =   270
            Width           =   915
         End
      End
      Begin VB.Frame fram_dias 
         BackColor       =   &H00000000&
         Height          =   2175
         Left            =   6360
         TabIndex        =   10
         ToolTipText     =   "Listado de Progrmación de los dias o el dia que queres activar el timbre."
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CheckBox Check1 
            BackColor       =   &H0000FF00&
            Caption         =   "domingo"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   1560
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000080FF&
            Caption         =   "Sabado"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Viernes"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Jueves"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Miercoles"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Martes"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Lunes"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FF00FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DIAS"
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
            Index           =   5
            Left            =   360
            TabIndex        =   18
            Top             =   1850
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   5520
         Picture         =   "frmtimbre.frx":0CCA
         ScaleHeight     =   975
         ScaleWidth      =   2655
         TabIndex        =   7
         Top             =   240
         Width           =   2655
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cargando..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   435
            Left            =   150
            TabIndex        =   8
            Top             =   170
            Width           =   2220
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Tipo :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   1250
         TabIndex        =   9
         Top             =   4680
         Width           =   4485
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Hora que se Activo :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   200
         TabIndex        =   6
         Top             =   4440
         Width           =   5370
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Timpo Restante :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   440
         TabIndex        =   5
         Top             =   4200
         Width           =   5370
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Timpo Trascurrido :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   290
         TabIndex        =   4
         Top             =   3960
         Width           =   5370
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Timpo Total :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   3720
         Width           =   5370
      End
   End
   Begin Virtual_Martin_temporize.ChameleonBtn Command1 
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cerrar"
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
      MICON           =   "frmtimbre.frx":10C80
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   5520
      Width           =   8295
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   14631
      _cy             =   1085
   End
   Begin VB.Label labinfo 
      BackStyle       =   0  'Transparent
      Caption         =   "El Timbre se esta ejecutando..."
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
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   2880
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmtimbre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Detonador de  Martin temporize v1.0
'*
'*
'***************************************************************************
Public timpo_programado, restante, trascurrido As Integer
Public comentario_general As String

Private Sub Command1_Click()
 Finalizar
 Unload Me
End Sub

Private Sub Form_Load()
 On Error GoTo nose
 Me.Icon = frmprograma.Icon
 restante = timpo_programado
 Text1.Text = comentario_general
 frmProgramacon.Show
 'Mopuerto.disparar_bit ' ejecuta el puerto de salida del pc
 wmp1.URL = frmProgramacon.wmp.URL
 frmProgramacon.wmp.Close
 Unload frmProgramacon
 wmp1.settings.volume = 100
 wmp1.Controls.play
 wmp1.settings.playCount = 1000
nose:
End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error GoTo no_se
 Shell frmprograma.liscomando.List(frmprograma.liscomando.ListIndex), _
 vbNormalNoFocus
no_se:
 Command1_Click
 timpo_programado = 0
 restante = 0
 trascurrido = 0
 'Finalizar no dispara al puerto
 frmprograma.guardard_Click
 Unload frmProgramacon
End Sub

Public Sub Finalizar()
'puerto.apagar_puertos ' apaga todos los puertos de LTP1
End Sub

Private Sub tiempo_Timer()
 trascurrido = trascurrido + 1: restante = restante - 1
 Label1(1).Caption = "Tiempo Trascurrido :" & " " & trascurrido & " " & " seg."
 Label1(2).Caption = "Tiempo Restante :" & " " & restante & " " & "seg."
 Command1.Caption = "&Cerrar" & " " & "(" & restante & ")"
 funcin_cerrar
End Sub

Private Sub timreloj_Timer()
 Label2.Caption = Time & " " & "Hs"
End Sub

Private Sub funcin_cerrar()
 If timpo_programado = trascurrido Then
 trascurrido = 0 ' destrullo la hora
 fram_dias.Visible = False
 frmsolo_hora.Visible = False
 'apagarTodo_Puerto
 Unload Me
 End If
End Sub
