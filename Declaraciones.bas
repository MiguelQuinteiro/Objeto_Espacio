Attribute VB_Name = "Declaraciones"

Option Explicit

' ************ Espacio
Public Type Coordenadas2D
  X As Double
  Y As Double
End Type

Public Type Coordenadas3D
  X As Double
  Y As Double
  Z As Double
End Type

' ************ Carcteristicas
Public Type Atributos
  Tamaño As Double
  Color As Double
  Visible As Boolean
End Type

' ************ Elementos Iniciales
Public Type Puntos
  C3 As Coordenadas3D
  A As Atributos
  C2 As Coordenadas2D
End Type

Public Type Lineas
  P1 As Puntos
  P2 As Puntos
  A As Atributos
End Type

' ************ Ojetos 3D
Public Type Cubo
  P(8) As Puntos
  L(12) As Lineas
End Type

' Declaración de variables
Public ppX As Double
Public ppY As Double
Public ang As Double

' Declaración de variables
Public CentroX As Double
Public CentroY As Double
Public CentroZ As Double

' Declaración de Constantes
Public Const miPi = 3.14159265358979
Public Const grados = (miPi / 180)

Public Resultado(1 To 1, 1 To 4) As Double

