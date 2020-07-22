Attribute VB_Name = "puertof"
'***************************************************************************
'*
'*
'* puerto paralelo Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Private Declare Sub PortOut Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Data As Byte)
 Private Declare Sub PortWordOut Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Data As Integer)
 Private Declare Sub PortDWordOut Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Data As Long)
 Private Declare Function PortIn Lib "IO.DLL" (ByVal Port _
 As Integer) As Byte
 Private Declare Function PortWordIn Lib "IO.DLL" (ByVal Port _
 As Integer) As Integer
 Private Declare Function PortDWordIn Lib "IO.DLL" (ByVal Port _
 As Integer) As Long
 Private Declare Sub SetPortBit Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Bit As Byte)
 Private Declare Sub ClrPortBit Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Bit As Byte)
 Private Declare Sub NotPortBit Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Bit As Byte)
 Private Declare Function GetPortBit Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Bit As Byte) As Boolean
 Private Declare Function RightPortShift Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Val As Boolean) As Boolean
 Private Declare Function LeftPortShift Lib "IO.DLL" (ByVal Port _
 As Integer, ByVal Val As Boolean) As Boolean
 Private Declare Function IsDriverInstalled Lib "IO.DLL" () As Boolean
 Dim estadopin As Boolean
'Puerto de salisda del pc
Public pu1, pu2, pu3, pu4, pu5, pu6, pu7, pu8 As Byte
'----------------------------------------------------------------------
Public puerto1, puerto2, puerto3, puerto4, _
puerto5, puerto6, puerto7, puerto8 As Boolean
'----------------------------------------------------------------------

Public Sub disparar_bit()
Select Case pu1
 Case (0)
 puerto1 = False
 Case (1)
 puerto1 = True
End Select

Select Case pu2
 Case (0)
 puerto2 = False
 Case (1)
 puerto2 = True
End Select

Select Case pu3
 Case (0)
 puerto3 = False
 Case (1)
 puerto3 = True
End Select

Select Case pu4
 Case (0)
 puerto4 = False
 Case (1)
 puerto4 = True
End Select

Select Case pu5
 Case (0)
 puerto5 = False
 Case (1)
 puerto5 = True
End Select

Select Case pu6
 Case (0)
 puerto6 = False
 Case (1)
 puerto6 = True
End Select

Select Case pu7
 Case (0)
 puerto7 = False
 Case (1)
 puerto7 = True
End Select

Select Case pu8
 Case (0)
 puerto8 = False
 Case (1)
 puerto8 = True
End Select
encender puerto1, puerto2, puerto3, puerto4 _
, puerto5, puerto6, puerto7, puerto8
End Sub

Private Sub encend(ByVal puerto As Byte)
 estadopin = GetPortBit(&H378, puerto)
 If (estadopin = False) Then
 outpuerto = &H378
 SetPortBit outpuerto, puerto
 End If
End Sub

Private Sub apagar(ByVal puerto As Byte)
 estadopin = GetPortBit(&H378, puerto)
 If (estadopin = True) Then
 outpuerto = &H378
 ClrPortBit outpuerto, puerto
 End If
End Sub

Public Sub apagar_puertos()
 outpuerto = &H378
 PortOut outpuerto, 0
End Sub

Private Sub encender(ByVal pin1 As Boolean, ByVal pin2 As Boolean, ByVal pin3 _
As Boolean, ByVal pin4 As Boolean, ByVal pin5 As Boolean, ByVal pin6 As Boolean, _
ByVal pin7 As Boolean, ByVal pin8 As Boolean)
Select Case pin1
 Case (True)
 encend 0
 Case (False)
 apagar 0
End Select
Select Case pin2
 Case (True)
 encend 1
 Case (False)
 apagar 1
End Select
Select Case pin3
 Case (True)
 encend 2
 Case (False)
 apagar 2
End Select
Select Case pin4
 Case (True)
 encend 3
 Case (False)
 apagar 3
End Select
Select Case pin5
 Case (True)
 encend 4
 Case (False)
 apagar 4
End Select
Select Case pin6
 Case (True)
 encend 5
 Case (False)
 apagar 5
End Select
Select Case pin7
 Case (True)
 encend 6
 Case (False)
 apagar 6
End Select
Select Case pin8
 Case (True)
 encend 7
 Case (False)
 apagar 7
End Select
End Sub
