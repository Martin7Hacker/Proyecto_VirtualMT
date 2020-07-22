VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "controles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Funcion Mover Datos Controles con Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Dim c_n As Boolean
Dim d_c(1) As String
Dim i_c(1) As Integer

Function moverDato(ByVal lista1 As ListBox)
 Select Case (c_n)
  Case False
  i_c(0) = lista1.ListIndex
  d_c(0) = lista1.List(i_c(0))
  lista1.List(i_c(0)) = d_c(0)
  c_n = True
  Case True
  i_c(1) = lista1.ListIndex
  d_c(1) = lista1.List(i_c(1))
  lista1.List(i_c(1)) = d_c(0)
  lista1.List(i_c(0)) = d_c(1)
  c_n = False
 End Select
End Function
