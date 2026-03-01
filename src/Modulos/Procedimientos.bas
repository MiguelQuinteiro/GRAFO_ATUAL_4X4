Attribute VB_Name = "Procedimientos"
'****************************************************************************************
'* PROYECTO      : GRAFO
'* CONTENIDO     : PROCEDIMIENTOS PARA PROYECTO GRAFOS DEL SUDOKU
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO
'* INICIO        : 19 DE FEBRERO DE 2017
'* ACTUALIZACION : 19 DE FEBRERO DE 2017
'****************************************************************************************
Option Explicit

' REINICIAR LOS VERTICES Y LAS ARISTAS
Public Sub ReiniciaVerticesAristas()
' Instancia los vertices
'Dim oV() As clsVertice
  ReDim oV(miNumeroVertices)
  Dim v As Integer
  For v = 1 To miNumeroVertices
    Set oV(v) = New clsVertice
    With oV(v)
      .Identificacion = "V. " & Trim(Str(v))
      .Numero = v
      .Color = miColorVertices
      .Tamaño = miTamañoVertices
      .Forma = miFormaVertices
    End With
  Next v
  ' Instancia los aristas
  'Dim oA() As clsArista
  ReDim oA(miNumeroAristas)
  Dim a As Integer
  Dim v1 As Integer
  Dim v2 As Integer
  a = 0
  For v1 = 1 To miNumeroVertices
    For v2 = v1 + 1 To miNumeroVertices
      a = a + 1
      Set oA(a) = New clsArista
      With oA(a)
        .Identificacion = "A. " & Trim(Str(a))
        .Numero = a
        .Color = miColorAristas
        .Vertice1 = v1
        .Vertice2 = v2
      End With
    Next v2
  Next v1
End Sub

' CALCULAR CIRCUNFERENCIA
Public Sub CarcularCircunferencia(ByRef pV As clsVertice, ByVal pN As Integer)
  pV.PosX = miRadio * Cos((360 / miNumeroVertices) * (miPi / 180) * pN) * miFactorCircular
  pV.PosY = miRadio * Sin((360 / miNumeroVertices) * (miPi / 180) * pN) * miFactorCircular
End Sub

' CALCULAR EL TESERACTO
Public Sub CalcularTeseracto(ByRef pVector() As Integer)
  Dim dX As Integer
  Dim dY As Integer
  ' Ubicar los Vertices
  dX = -3514
  dY = -3136
  'Vértices X Y
  '1
  oV(pVector(1)).PosX = (6650 + dX) * miFactorExterno
  oV(pVector(1)).PosY = (4718 + dY) * miFactorExterno
  '2
  oV(pVector(2)).PosX = (4550 + dX) * miFactorExterno
  oV(pVector(2)).PosY = (5782 + dY) * miFactorExterno
  oV(pVector(2)).Tamaño = miTamañoVertices - 90
  '3
  oV(pVector(3)).PosX = (350 + dX) * miFactorExterno
  oV(pVector(3)).PosY = (5782 + dY) * miFactorExterno
  oV(pVector(3)).Tamaño = miTamañoVertices - 90
  '4
  oV(pVector(4)).PosX = (2436 + dX) * miFactorExterno
  oV(pVector(4)).PosY = (4715 + dY) * miFactorExterno
  '5
  oV(pVector(5)).PosX = (6650 + dX) * miFactorExterno
  oV(pVector(5)).PosY = (504 + dY) * miFactorExterno
  '6
  oV(pVector(6)).PosX = (4550 + dX) * miFactorExterno
  oV(pVector(6)).PosY = (1890 + dY) * miFactorExterno
  oV(pVector(6)).Tamaño = miTamañoVertices - 90
  '7
  oV(pVector(7)).PosX = (350 + dX) * miFactorExterno
  oV(pVector(7)).PosY = (1890 + dY) * miFactorExterno
  oV(pVector(7)).Tamaño = miTamañoVertices - 90
  '8
  oV(pVector(8)).PosX = (2436 + dX) * miFactorExterno
  oV(pVector(8)).PosY = (504 + dY) * miFactorExterno
  '9
  oV(pVector(9)).PosX = (4284 + dX) * miFactorInterno
  oV(pVector(9)).PosY = (3528 + dY) * miFactorInterno
  oV(pVector(9)).Tamaño = miTamañoVertices - 40
  '10
  oV(pVector(10)).PosX = (3766 + dX) * miFactorInterno
  oV(pVector(10)).PosY = (3794 + dY) * miFactorInterno
  oV(pVector(10)).Tamaño = miTamañoVertices - 60
  '11
  oV(pVector(11)).PosX = (2716 + dX) * miFactorInterno
  oV(pVector(11)).PosY = (3794 + dY) * miFactorInterno
  oV(pVector(11)).Tamaño = miTamañoVertices - 60
  '12
  oV(pVector(12)).PosX = (3234 + dX) * miFactorInterno
  oV(pVector(12)).PosY = (3528 + dY) * miFactorInterno
  oV(pVector(12)).Tamaño = miTamañoVertices - 40
  '13
  oV(pVector(13)).PosX = (4284 + dX) * miFactorInterno
  oV(pVector(13)).PosY = (2478 + dY) * miFactorInterno
  oV(pVector(13)).Tamaño = miTamañoVertices - 40
  '14
  oV(pVector(14)).PosX = (3766 + dX) * miFactorInterno
  'oV(pVector(14)).PosY = (2744 + dY) * miFactorInterno
  oV(pVector(14)).PosY = (2824 + dY) * miFactorInterno
  oV(pVector(14)).Tamaño = miTamañoVertices - 60
  '15
  oV(pVector(15)).PosX = (2716 + dX) * miFactorInterno
  'oV(pVector(15)).PosY = (2744 + dY) * miFactorInterno
  oV(pVector(15)).PosY = (2824 + dY) * miFactorInterno
  oV(pVector(15)).Tamaño = miTamañoVertices - 60
  '16
  oV(pVector(16)).PosX = (3234 + dX) * miFactorInterno
  oV(pVector(16)).PosY = (2478 + dY) * miFactorInterno
  oV(pVector(16)).Tamaño = miTamañoVertices - 40
End Sub

' RESALTAR EL MODELO SELECCIONADO
Public Sub ResaltarModelo(ByVal pM As Integer, ByVal pV1 As Integer, ByVal pV2 As Integer, ByVal pV3 As Integer, ByVal pV4 As Integer)
' Cambia color vertices
  oV(pV1).Color = 9
  oV(pV2).Color = 14
  oV(pV3).Color = 12
  oV(pV4).Color = 10
  ' Cambia color aristas
  Dim a As Integer
  For a = 1 To miNumeroAristas
    If (oA(a).Vertice1 = pV1) Or (oA(a).Vertice1 = pV2) Or (oA(a).Vertice1 = pV3) Or (oA(a).Vertice1 = pV4) Then
      If (oA(a).Vertice2 = pV1) Or (oA(a).Vertice2 = pV2) Or (oA(a).Vertice2 = pV3) Or (oA(a).Vertice2 = pV4) Then
        If (pM = 1) Or (pM = 5) Or (pM = 8) Or (pM = 12) Then
          oA(a).Color = 12
        End If
        If (pM = 2) Or (pM = 6) Or (pM = 7) Or (pM = 11) Then
          oA(a).Color = 14
        End If
        If (pM = 3) Or (pM = 4) Or (pM = 9) Or (pM = 10) Then
          oA(a).Color = 9
        End If
      End If
    End If
  Next a
End Sub

' CALCULAR ROTACION X
Public Sub CalcularRotacionX()
  Dim Aux As Integer
  Aux = miVec(4)
  miVec(4) = miVec(3)
  miVec(3) = miVec(2)
  miVec(2) = miVec(1)
  miVec(1) = Aux
  Aux = miVec(8)
  miVec(8) = miVec(7)
  miVec(7) = miVec(6)
  miVec(6) = miVec(5)
  miVec(5) = Aux
  Aux = miVec(12)
  miVec(12) = miVec(11)
  miVec(11) = miVec(10)
  miVec(10) = miVec(9)
  miVec(9) = Aux
  Aux = miVec(16)
  miVec(16) = miVec(15)
  miVec(15) = miVec(14)
  miVec(14) = miVec(13)
  miVec(13) = Aux
End Sub

' CALCULAR ROTACION Y
Public Sub CalcularRotacionY()
  Dim Aux As Integer
  Aux = miVec(5)
  miVec(5) = miVec(6)
  miVec(6) = miVec(2)
  miVec(2) = miVec(1)
  miVec(1) = Aux
  Aux = miVec(3)
  miVec(3) = miVec(4)
  miVec(4) = miVec(8)
  miVec(8) = miVec(7)
  miVec(7) = Aux
  Aux = miVec(13)
  miVec(13) = miVec(14)
  miVec(14) = miVec(10)
  miVec(10) = miVec(9)
  miVec(9) = Aux
  Aux = miVec(11)
  miVec(11) = miVec(12)
  miVec(12) = miVec(16)
  miVec(16) = miVec(15)
  miVec(15) = Aux
End Sub

' CALCULAR ROTACION Z
Public Sub CalcularRotacionZ()
  Dim Aux As Integer
  Aux = miVec(5)
  miVec(5) = miVec(8)
  miVec(8) = miVec(4)
  miVec(4) = miVec(1)
  miVec(1) = Aux
  Aux = miVec(3)
  miVec(3) = miVec(2)
  miVec(2) = miVec(6)
  miVec(6) = miVec(7)
  miVec(7) = Aux
  Aux = miVec(13)
  miVec(13) = miVec(16)
  miVec(16) = miVec(12)
  miVec(12) = miVec(9)
  miVec(9) = Aux
  Aux = miVec(11)
  miVec(11) = miVec(10)
  miVec(10) = miVec(14)
  miVec(14) = miVec(15)
  miVec(15) = Aux
End Sub

' CALCULAR DESPLAZAMIENTO X
Public Sub CalcularDespalzamientoX()
  Dim Aux As Integer
  Aux = miVec(4)
  miVec(4) = miVec(12)
  miVec(12) = miVec(9)
  miVec(9) = miVec(1)
  miVec(1) = Aux
  Aux = miVec(3)
  miVec(3) = miVec(11)
  miVec(11) = miVec(10)
  miVec(10) = miVec(2)
  miVec(2) = Aux
  Aux = miVec(8)
  miVec(8) = miVec(16)
  miVec(16) = miVec(13)
  miVec(13) = miVec(5)
  miVec(5) = Aux
  Aux = miVec(7)
  miVec(7) = miVec(15)
  miVec(15) = miVec(14)
  miVec(14) = miVec(6)
  miVec(6) = Aux
End Sub

' CALCULAR DESPLAZAMIENTO Y
Public Sub CalcularDespalzamientoY()
  Dim Aux As Integer
  Aux = miVec(5)
  miVec(5) = miVec(13)
  miVec(13) = miVec(9)
  miVec(9) = miVec(1)
  miVec(1) = Aux
  Aux = miVec(6)
  miVec(6) = miVec(14)
  miVec(14) = miVec(10)
  miVec(10) = miVec(2)
  miVec(2) = Aux
  Aux = miVec(8)
  miVec(8) = miVec(16)
  miVec(16) = miVec(12)
  miVec(12) = miVec(4)
  miVec(4) = Aux
  Aux = miVec(7)
  miVec(7) = miVec(15)
  miVec(15) = miVec(11)
  miVec(11) = miVec(3)
  miVec(3) = Aux
End Sub

' CALCULAR DESPLAZAMIENTO Z
Public Sub CalcularDespalzamientoZ()
  Dim Aux As Integer
  Aux = miVec(2)
  miVec(2) = miVec(10)
  miVec(10) = miVec(9)
  miVec(9) = miVec(1)
  miVec(1) = Aux
  Aux = miVec(3)
  miVec(3) = miVec(11)
  miVec(11) = miVec(12)
  miVec(12) = miVec(4)
  miVec(4) = Aux
  Aux = miVec(6)
  miVec(6) = miVec(14)
  miVec(14) = miVec(13)
  miVec(13) = miVec(5)
  miVec(5) = Aux
  Aux = miVec(7)
  miVec(7) = miVec(15)
  miVec(15) = miVec(16)
  miVec(16) = miVec(8)
  miVec(8) = Aux
End Sub


