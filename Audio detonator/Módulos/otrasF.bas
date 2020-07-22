Attribute VB_Name = "OtrasF"
'***************************************************************************
'*
'*
'* Comparar Fechas con Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Option Explicit

Public Function ComparaFechas(ByVal F1 As Date, ByVal F2 As Date) As Boolean
 ComparaFechas = False
 If (CDate(F1) < CDate(F2)) Then Exit Function
 ComparaFechas = True
End Function

Public Function LastError$()
 LastError$ = "Se ha generado el Código de Error No. " _
 & NumErr& & vbCrLf & vbCrLf & _
 "Motivo: " & MsgErr$
End Function

