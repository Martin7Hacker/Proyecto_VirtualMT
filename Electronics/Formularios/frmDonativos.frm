VERSION 5.00
Begin VB.Form frmDonativos 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Donativos"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "frmDonativos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pdonar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2160
      MouseIcon       =   "frmDonativos.frx":0CCA
      Picture         =   "frmDonativos.frx":0FD4
      ScaleHeight     =   1185
      ScaleWidth      =   2940
      TabIndex        =   4
      Top             =   1680
      Width           =   2970
   End
   Begin VB.PictureBox ptargeta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2520
      Picture         =   "frmDonativos.frx":10206
      ScaleHeight     =   225
      ScaleWidth      =   2175
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcolaborar 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Colaborar"
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
      MICON           =   "frmDonativos.frx":10A01
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAceptar 
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      MICON           =   "frmDonativos.frx":10A1D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "THANK YOU YERY MUCH ALWAYSLOVE EE:UU"
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
      Height          =   555
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   3165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "para cumplir mi sueño de ir a EE:UU"
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
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   3165
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "con cuenta propia..."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Label lblcard 
      BackStyle       =   0  'Transparent
      Caption         =   "Con tarjetas de créditos"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2760
      TabIndex        =   2
      Top             =   3000
      Width           =   1665
   End
   Begin VB.Label lblsoftware 
      BackStyle       =   0  'Transparent
      Caption         =   "to fulfill my dream of going to EE:UU."
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
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "frmDonativos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Para realizar donacíones para el proyecto Virtual Martin temporize v1.0
'*
'*
'***************************************************************************

Private Declare Function ShellExecute Lib _
 "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, _
 ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAceptar_Click()
 Unload Me
End Sub

Private Sub cmdcolaborar_Click()
 ptargeta_Click
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
End Sub

Private Sub Label1_Click()
 ptargeta_Click
End Sub

Private Sub lblcard_Click()
 ptargeta_Click
End Sub

Private Sub pdonar_Click()
 ptargeta_Click
End Sub

Private Sub ptargeta_Click()
 Dim x As String
 x = ShellExecute(Me.hwnd, "Open" _
 , "http://martinsoft0.blogspot.com/p/donar.html", _
 &O0, &O0, 0)
 Unload Me
End Sub
