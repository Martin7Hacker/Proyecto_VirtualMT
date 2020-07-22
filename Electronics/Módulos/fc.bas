Attribute VB_Name = "fc"
'***************************************************************************
'*
'*
'* Procedimiento Abreviado  Virtual Martin temporize v1.0
'*
'*
'***************************************************************************

Public Sub comp_clave_fSalir(ByVal camb As Boolean, ByVal cla_num _
 As Byte, ByVal cla_hex As String, ByVal comp_num As Byte, com_hex _
 As String, ByVal ventana As Form)
 'modo numerico
 Select Case (camb)
 Case (False)
 If cla_num = comp_num Then
 f_salir ventana
 End If
 If cla_hex = com_hex Then
 f_salir ventana
 End If
 Case (True)
 If cla_num = comp_num And cla_hex = com_hex Then
 f_salir ventana
 End If
End Select
End Sub

Private Sub f_salir(ByVal vent As Form)
Unload vent
End Sub
