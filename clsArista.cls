VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : GRAFO
'* CONTENIDO     : CLASE ARISTA
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO
'* INICIO        : 19 DE FEBRERO DE 2017
'* ACTUALIZACION : 19 DE FEBRERO DE 2017
'****************************************************************************************
Option Explicit

'****************************************************************************************
' CREAR CLASE VISUAL BASIC
'****************************************************************************************

' AREA DE DATOS
'*******************************************************************
' Declaracion de los Atributos privados para encapsulamiento
Private miIdentificacion As String
Private miNumero As Integer
Private miVertice1 As Integer
Private miVertice2 As Integer
Private miColor As Integer
'Private miPropiedad As String

' Metodos Get y Let para manipulacion de datos de la clase
' Get
Public Property Get Identificacion() As String
  Identificacion = miIdentificacion
End Property
' Let
Public Property Let Identificacion(ByVal valor As String)
  miIdentificacion = valor
End Property

' Get
Public Property Get Numero() As Integer
  Numero = miNumero
End Property
' Let
Public Property Let Numero(ByVal valor As Integer)
  miNumero = valor
End Property

' Get
Public Property Get Vertice1() As Integer
  Vertice1 = miVertice1
End Property
' Let
Public Property Let Vertice1(ByVal valor As Integer)
  miVertice1 = valor
End Property

' Get
Public Property Get Vertice2() As Integer
  Vertice2 = miVertice2
End Property
' Let
Public Property Let Vertice2(ByVal valor As Integer)
  miVertice2 = valor
End Property

' Get
Public Property Get Color() As String
  Color = miColor
End Property
' Let
Public Property Let Color(ByVal valor As String)
  miColor = valor
End Property

' AREA DE METODOS Y FUNCIONALIDAD
'*******************************************************************
' Metodo constructor y destructor de la clase
' Constructor
Private Sub Class_Initialize()
' Que hacer al instanciar el objeto
  miIdentificacion = ""
  miNumero = 0
  'miVertice1 = 0
  'miVertice2 = 0
  miColor = 0
End Sub
' Destructor
Sub Class_Terminate()
' Que hacer al destruir el objeto
End Sub

' Metodos publicos para la funcionalidad de la clase
Public Sub miSubrutina()
' Instrucciones
End Sub
Public Function miFuncion() As String
' Instrucciones
End Function

' Metodos privados auxiliares para la funcionalidad de la clase
Private Sub miSubrutinaAuxiliar()
' Instrucciones
End Sub
Private Function miFuncionAuxiliar() As String
' Instrucciones
End Function
'*******************************************************************

