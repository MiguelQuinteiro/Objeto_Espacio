Attribute VB_Name = "Funciones"

Option Explicit

' Grafica un punto en espacio vectorial de tres dimensiones 3D
Public Function CoordenadaXPantalla(ByVal pX As Double, ByVal pY As Double, ByVal pZ As Double) As Double
' Coordenadas de Pantalla de X
  CoordenadaXPantalla = CentroY + (-pX * Cos(ang)) + (pY) + (0)
End Function

' Grafica un punto en espacio vectorial de tres dimensiones 3D
Public Function CoordenadaYPantalla(ByVal pX As Double, ByVal pY As Double, ByVal pZ As Double) As Double
' Coordenadas de Pantalla de Y
  CoordenadaYPantalla = CentroZ + (pX * Sin(ang)) + (0) + (-pZ)
End Function

