Attribute VB_Name = "Lenguage"
'***************************************************************************
'*
'*
'* Lenguaje con Virtual Martin temporize v1.0
'*
'*
'***************************************************************************
Public lenguage_opciones(23) As String           ' representa un vector de cadena de texto.
Public lenguage_opciones_generador(40) As String ' representa un vector de cadena de texto.
Public lenguage_rutas(10) As String              ' representa un vector de caracter para rutas
Public lenguage_iniciarwindows(7) As String      ' representa un vector de inicio
Public lenguage_circuito(3) As String            ' reperesenta un vector de lenguage del circuito
Public lenguage_estVentana(6) As String          ' representa un vector de estado
Public lenguage_datosCreador(7) As String        ' representa un vector de datos
Public lenguage_fichaCreador(20) As String       ' representa un vector de datos creado
Public lenguage_estFunciones(17) As String       ' representa un vector de datos
Public lenguage_estcopias(5)     As String       ' de funciones al sistema
Public lenguage_memoria(10)      As String
Public lenguage_crearModificar(25)   As String

'Const c_opc As Byte = 16

Public Sub definir_lenguage_opciones() 'estructura para ejecutar el lenguage del programa
                                       'es desir hace posible que el programa dentro
                                       'del menú opciónes allan diferentes opciones.
                                       'de idioma.
 lenguage_opciones(0) = "Opciónes de modificado"
 lenguage_opciones(1) = "&" & "Modificado de datos"
 lenguage_opciones(2) = "&" & "Oprimiendo los bótones"
 lenguage_opciones(3) = "&" & "Hora"
 lenguage_opciones(4) = "&" & "Tipo"
 lenguage_opciones(5) = "&" & "Filtro"
 lenguage_opciones(6) = "&" & "Int Entrada"
 lenguage_opciones(7) = "&" & "Int Salida"
 lenguage_opciones(8) = "&" & "Texto Entrada"
 lenguage_opciones(9) = "&" & "Texto Salida"
 lenguage_opciones(10) = "&" & "Lunes"
 lenguage_opciones(11) = "&" & "Martes"
 lenguage_opciones(12) = "&" & "Miercoles"
 lenguage_opciones(13) = "&" & "Jueves"
 lenguage_opciones(14) = "&" & "Viernes"
 lenguage_opciones(15) = "&" & "Sabados"
 lenguage_opciones(16) = "&" & "Domingos"
 lenguage_opciones(17) = "&" & "Tipo de Aplicado:"
 lenguage_opciones(18) = "Se niegan las opciones oprimidas"
 lenguage_opciones(19) = "Se niegan las opciónes no oprimidas"
 lenguage_opciones(20) = "&" & "Aplicar"
 lenguage_opciones(21) = "&" & "Salir"
 lenguage_opciones(22) = "&" & "C: on/off"
 lenguage_opciones(23) = "&" & "Restaurar"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'generador de timbres
 '################################################################
 lenguage_opciones_generador(0) = "Generador de Rutinas"
 lenguage_opciones_generador(1) = "&" & "Lista Desplegable"
 lenguage_opciones_generador(2) = "&" & "Programa Desplegable"
 lenguage_opciones_generador(3) = "&" & "Opciones de Modificado"
 lenguage_opciones_generador(4) = "&" & "Hora:"
 lenguage_opciones_generador(5) = "&" & "Tipo:"
 lenguage_opciones_generador(6) = "&" & "Filtro:"
 lenguage_opciones_generador(7) = "&" & "Int:                                  [Entrada]"
 lenguage_opciones_generador(8) = "&" & "Int:                                  [Salida]"
 lenguage_opciones_generador(9) = "&" & "DIAS"
 lenguage_opciones_generador(10) = "&" & "Lunes"
 lenguage_opciones_generador(11) = "&" & "Martes"
 lenguage_opciones_generador(12) = "&" & "Miercoles"
 lenguage_opciones_generador(13) = "&" & "Jueves"
 lenguage_opciones_generador(14) = "&" & "Viernes"
 lenguage_opciones_generador(15) = "&" & "Sabados"
 lenguage_opciones_generador(16) = "&" & "Domingos"
 lenguage_opciones_generador(17) = "&" & "No existe ningun elemento de evento."
 lenguage_opciones_generador(18) = "&" & "Existen actualmente"
 lenguage_opciones_generador(19) = "&" & "elementos de evento"
 lenguage_opciones_generador(20) = "&" & "elemento de evento"
 lenguage_opciones_generador(21) = "&" & "[" & "ENTRADA" & "]"
 lenguage_opciones_generador(22) = "&" & "[" & "SALIDA" & "]"
 lenguage_opciones_generador(23) = "&" & " Cancelar"
 lenguage_opciones_generador(24) = "&" & "Crear Evento:              "
 lenguage_opciones_generador(25) = "&" & "Crear Eventos:             "
 lenguage_opciones_generador(26) = "&" & "Modificar"
 lenguage_opciones_generador(27) = "No hacer nada *"
 lenguage_opciones_generador(28) = "Apagar el Equipo"
 lenguage_opciones_generador(29) = "Apagar y reiniciar el equipo"
 lenguage_opciones_generador(30) = "Anular el Apagado de equipo"
 lenguage_opciones_generador(31) = "Equipo que se apagara / reiniciara / anulara"
 lenguage_opciones_generador(32) = "Establecer el tiempo de espera de apagado."
 lenguage_opciones_generador(33) = "Comentario de apagado máximo, 127 caracteres"
 lenguage_opciones_generador(34) = "Forzar el cierre de todas las aplicaciones sin advertir"
 lenguage_opciones_generador(35) = "Ingrese descripción máximo, 127 caracteres"
 lenguage_opciones_generador(36) = "&" & "Sin dialogo..."
 lenguage_opciones_generador(37) = "&" & "Tiempo ="
 lenguage_opciones_generador(38) = "encendido."
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'rutas de Archivos
 '################################################################
 lenguage_rutas(0) = "Historial de Rutas"
 lenguage_rutas(1) = "&" & "Historial de Archivos definidos"
 lenguage_rutas(2) = "&" & "Cancelar"
 lenguage_rutas(3) = "&" & "Cargar"
 lenguage_rutas(4) = "&" & "Borrar Selección"
 lenguage_rutas(5) = "&" & "Borrar Todo"
 lenguage_rutas(6) = "&" & "Usar Archivo"
 lenguage_rutas(7) = "&" & "Des Usar Archivo"
 lenguage_rutas(8) = "&" & "Aceptar"
 lenguage_rutas(9) = "Quieres utilizar este archivo con  Microtime" & " "
 lenguage_rutas(10) = "¿ Quieres eliminar el Archivo usado de Memoria ?"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'Inicar Windows
 '################################################################
 lenguage_iniciarwindows(0) = "Virtual Martin temporize: Iniciar con Windows"
 lenguage_iniciarwindows(1) = "&" & "¿ Arrancar con Windows ?"
 lenguage_iniciarwindows(2) = "&" & "Arrancar"
 lenguage_iniciarwindows(3) = "&" & "No Arrancar"
 lenguage_iniciarwindows(4) = "&" & "Aceptar"
 lenguage_iniciarwindows(5) = "Cuando Inicie o Reinicie Windows Virtual Martin temporize Arrancara con Windows"
 lenguage_iniciarwindows(6) = " Hubo un error, Para Iniciar Con Windows S.O"
 lenguage_iniciarwindows(7) = "Se elimino el Arranque Automatico en Windows S.O"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'circuito impreso
 '################################################################
 lenguage_circuito(0) = "Circuito Electrónico"
 lenguage_circuito(1) = "Esqemas"
 lenguage_circuito(2) = "&" & "Imprimir"
 lenguage_circuito(3) = "&" & "Aceptar"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'estados de la ventana principal
 '################################################################
 lenguage_estVentana(0) = "Estado de La Ventana Principal"
 lenguage_estVentana(1) = "Estado:"
 lenguage_estVentana(2) = "&" & "Cancelar"
 lenguage_estVentana(3) = "&" & "Aplicar"
 lenguage_estVentana(4) = "Ventana Restaurada"
 lenguage_estVentana(5) = "Ventana Minimizada"
 lenguage_estVentana(6) = "Ventana Maximizada"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'datos del creador
 '################################################################
 lenguage_datosCreador(0) = "Datos del Creador del Software"
 lenguage_datosCreador(1) = "&" & "Teléfono : 43327818"
 lenguage_datosCreador(2) = "&" & "Cel : 091432320"
 lenguage_datosCreador(3) = "&" & "Email : CirculoWeb@hotmail.com"
 lenguage_datosCreador(4) = "&" & "Facebook : Martin Grasso ."
 lenguage_datosCreador(5) = "&" & "Localidad: Canelones, Uruguay."
 lenguage_datosCreador(6) = "&" & "Localidad: Tala, Uruguay."
 lenguage_datosCreador(7) = "&" & "Aceptar"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'datos del creador
 '################################################################
 lenguage_fichaCreador(0) = "Personalizar datos"
 lenguage_fichaCreador(1) = "Datos del Creador de los Timbres"
 lenguage_fichaCreador(2) = "Aceptar"
 lenguage_fichaCreador(3) = "Nombre :"
 lenguage_fichaCreador(4) = "Segundo Nombre :"
 lenguage_fichaCreador(5) = "Apellido :"
 lenguage_fichaCreador(6) = "Segundo Apellido :"
 lenguage_fichaCreador(7) = "Dirección :"
 lenguage_fichaCreador(8) = "Segunda Dirección :"
 lenguage_fichaCreador(9) = "Localidad :"
 lenguage_fichaCreador(10) = "Pais :"
 lenguage_fichaCreador(11) = "Teléfono :"
 lenguage_fichaCreador(12) = "Celular :"
 lenguage_fichaCreador(13) = "Correo Electrónico :"
 lenguage_fichaCreador(14) = "Facebook :"
 lenguage_fichaCreador(15) = "Comentario General :"
 lenguage_fichaCreador(16) = "Cancelar"
 lenguage_fichaCreador(17) = "Limpiar"
 lenguage_fichaCreador(18) = "&Aceptar"
 lenguage_fichaCreador(19) = " ¿ Quieres Limpiar Todos los Datos en Pantalla ?"
 lenguage_fichaCreador(20) = "Los datos se guardaron en memoria con éxito"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'funciones
 '################################################################
 lenguage_estFunciones(0) = "Funciones al Sistema"
 lenguage_estFunciones(1) = "Funciones de Sistema Operables"
 lenguage_estFunciones(2) = "comentarios:"
 lenguage_estFunciones(3) = "No hacer nada *"
 lenguage_estFunciones(4) = "Apagar el Equipo"
 lenguage_estFunciones(5) = "Apagar y reiniciar el equipo"
 lenguage_estFunciones(6) = "Anular el Apagado de equipo"
 lenguage_estFunciones(7) = "Equipo que se apagara / reiniciara / anulara"
 lenguage_estFunciones(8) = "Establecer el tiempo de espera de apagado."
 lenguage_estFunciones(9) = "Comentario de apagado máximo, 127 caracteres"
 lenguage_estFunciones(10) = "Forzar el cierre de todas las aplicaciones sin advertir"
 lenguage_estFunciones(11) = "Ingrese descripción máximo, 127 caracteres"
 lenguage_estFunciones(12) = "&" & "Sin dialogo..."
 lenguage_estFunciones(13) = "&" & "Tiempo ="
 lenguage_estFunciones(14) = "encendido."
 lenguage_estFunciones(15) = "Cancelar"
 lenguage_estFunciones(16) = "Aplicar"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'Impresor por Cantidad
 '################################################################
 lenguage_estcopias(0) = "Impresor por Cantidad"
 lenguage_estcopias(1) = "Copias:"
 lenguage_estcopias(2) = "-"
 lenguage_estcopias(3) = "+"
 lenguage_estcopias(4) = "Cancelar"
 lenguage_estcopias(5) = "Mandar a Imprimir las Copias"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'archivos en Memoria
 '################################################################
 lenguage_memoria(0) = "Archivos actuales en la Memoria del Software"
 lenguage_memoria(1) = "¿Existen Archivos en memoria que desea Hacer?"
 lenguage_memoria(2) = "id"
 lenguage_memoria(3) = "Hora"
 lenguage_memoria(4) = "Tipo"
 lenguage_memoria(5) = "Segundos"
 lenguage_memoria(6) = "Comentario"
 lenguage_memoria(7) = "Salir"
 lenguage_memoria(8) = "Guardar y Salir"
 lenguage_memoria(9) = "Cancelar"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'crear modificar
 '################################################################
 lenguage_crearModificar(0) = "titulo programa"
 lenguage_crearModificar(1) = "Agregar Nuevo Evento"
 lenguage_crearModificar(2) = "Funciones al sistema"
 lenguage_crearModificar(3) = "Hora :"
 lenguage_crearModificar(4) = "Tipo :"
 lenguage_crearModificar(5) = "Intervalo :"
 lenguage_crearModificar(6) = "Filtro :"
 lenguage_crearModificar(7) = "Lunes"
 lenguage_crearModificar(8) = "Martes"
 lenguage_crearModificar(9) = "Miercoles"
 lenguage_crearModificar(10) = "Jueves"
 lenguage_crearModificar(11) = "Viernes"
 lenguage_crearModificar(12) = "Sabados"
 lenguage_crearModificar(13) = "Domingos"
 lenguage_crearModificar(14) = "Comentario:"
 lenguage_crearModificar(15) = "comentarios:"
 lenguage_crearModificar(16) = "Cancelar"
 lenguage_crearModificar(17) = "color de:Pintagrama"
 lenguage_crearModificar(18) = "Crear"
 lenguage_crearModificar(19) = "Modificar"
 lenguage_crearModificar(20) = "Modificar Evento."
 lenguage_crearModificar(21) = "Entrada"
 lenguage_crearModificar(22) = "Salir"
 lenguage_crearModificar(23) = "Solo Hora"
 lenguage_crearModificar(24) = "Hora y Dia"
 lenguage_crearModificar(25) = "Quieres Aplicar las Modificaciones del Evento."
End Sub
