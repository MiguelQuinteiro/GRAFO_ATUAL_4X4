Attribute VB_Name = "Declaraciones"
'****************************************************************************************
'* PROYECTO      : GRAFO
'* CONTENIDO     : DECLARACIONES DE VARIABLES PARA PROYECTO GRAFOS DEL SUDOKU
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO
'* INICIO        : 19 DE FEBRERO DE 2017
'* ACTUALIZACION : 19 DE FEBRERO DE 2017
'****************************************************************************************
Option Explicit

'PARA PODER GUARDAR IMAGEN DEL FORMULARIO
Public Declare Sub keybd_event _
                    Lib "user32" ( _
                        ByVal bVk As Byte, _
                        ByVal bScan As Byte, _
                        ByVal dwFlags As Long, _
                        ByVal dwExtraInfo As Long)

' VARIABLES PUBLICAS
Public oV() As clsVertice
Public oA() As clsArista

Public miNumeroVertices As Integer
Public miNumeroAristas As Integer

Public miPi As Double
Public miFactorCircular As Double
Public miFactorExterno As Double
Public miFactorInterno As Double
Public miFactorProfundidad As Double

Public miRadio As Integer
Public miColorFondo As Integer
Public miColorAristas As Integer
Public miColorVertices As Integer
Public miTamañoVertices As Integer
Public miFormaVertices As Integer

Public miVec(16) As Integer

