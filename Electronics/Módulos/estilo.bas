Attribute VB_Name = "estilo"
'***************************************************************************
'*
'*
'* estilo grafico para  Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Option Explicit

Private Type tagInitCommonControlsEx
 lngSize As Long
 lngICC As Long
End Type

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" _
(iccex As tagInitCommonControlsEx) As Boolean
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" ( _
ByVal hLibModule As Long) As Long
Public Declare Function SetErrorMode Lib "kernel32" ( _
ByVal wMode As Long) As Long
Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Global Const ICC_USEREX_CLASSES = &H200
Global Const SEM_NOGPFAULTERRORBOX = &H2&
Global m_bInIDE As Boolean

Public Sub UnloadApp()
 If Not InIDE() Then SetErrorMode SEM_NOGPFAULTERRORBOX
End Sub

Public Function InIDE() As Boolean
 Debug.Assert (IsInIDE())
 InIDE = m_bInIDE
End Function

Private Function IsInIDE() As Boolean
 m_bInIDE = True
 IsInIDE = m_bInIDE
End Function
 
Public Function InitCommonControlsVB() As Boolean
 On Error Resume Next
 Dim iccex As tagInitCommonControlsEx
 With iccex
 .lngSize = LenB(iccex)
 .lngICC = ICC_USEREX_CLASSES
 End With
 InitCommonControlsEx iccex
 InitCommonControlsVB = (Err.Number = 0)
 On Error GoTo 0
End Function
