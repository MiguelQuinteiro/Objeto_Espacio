VERSION 5.00
Begin VB.Form frmObjetoEspacio 
   BackColor       =   &H8000000E&
   Caption         =   " .:. Objeto Espacio .:."
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEscalado 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Text            =   "1"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtPz 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Text            =   "0"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtPy 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   16
      Text            =   "0"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtPx 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Text            =   "0"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtDz 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Text            =   "0"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtDy 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Text            =   "0"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtDx 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Text            =   "0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtTeta 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Text            =   "0"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtBeta 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Text            =   "0"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtAlfa 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Text            =   "0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdTransformacion 
      Caption         =   "Transformacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox txtAngulo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "45"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrueba 
      Caption         =   "Prueba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblEscalado 
      Caption         =   "Escalado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblPerspectiva 
      Caption         =   "Perspectiva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblTraslacion 
      Caption         =   "Traslación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblAngulo 
      Caption         =   "Ángulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblTZ 
      Caption         =   "Trans. Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblTY 
      Caption         =   "Trans. Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblTX 
      Caption         =   "Trans. X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmObjetoEspacio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim miCubo As Cubo
Dim miTamañoCubo As Double
Dim i As Integer

' Declaración de Punto
Dim Punto(1 To 1, 1 To 4) As Double
' Declaración de Transformacion
Dim Tx(1 To 4, 1 To 4) As Double
Dim Ty(1 To 4, 1 To 4) As Double
Dim Tz(1 To 4, 1 To 4) As Double
Dim alfa As Double
Dim beta As Double
Dim teta As Double
Dim Traslacion(1 To 1, 1 To 3) As Double
Dim Perspectiva(1 To 3) As Double
Dim Escala As Double

Private Sub cmdPrueba_Click()
' Ajuste del Angulo
  ang = Val(txtAngulo.Text) * (miPi / 180)
  miTamañoCubo = 1000

  ' Borrar la pantalla
  Cls

  ' Coordenadas del Centro
  CentroX = -6000 + 0
  CentroY = 1000 + frmObjetoEspacio.Width / 2
  CentroZ = -1000 + frmObjetoEspacio.Height / 2

  With miCubo
    ' Punto 1
    .P(1).C3.X = 0 * miTamañoCubo
    .P(1).C3.Y = 0 * miTamañoCubo
    .P(1).C3.Z = 0 * miTamañoCubo
    .P(1).A.Color = vbGreen
    .P(1).A.Tamaño = 100
    .P(1).A.Visible = True
    .P(1).C2.X = CoordenadaXPantalla(.P(1).C3.X, .P(1).C3.Y, .P(1).C3.Z)
    .P(1).C2.Y = CoordenadaYPantalla(.P(1).C3.X, .P(1).C3.Y, .P(1).C3.Z)

    ' Punto 2
    .P(2).C3.X = 0 * miTamañoCubo
    .P(2).C3.Y = 1 * miTamañoCubo
    .P(2).C3.Z = 0 * miTamañoCubo
    .P(2).A.Color = vbBlue
    .P(2).A.Tamaño = 100
    .P(2).A.Visible = True
    .P(2).C2.X = CoordenadaXPantalla(.P(2).C3.X, .P(2).C3.Y, .P(2).C3.Z)
    .P(2).C2.Y = CoordenadaYPantalla(.P(2).C3.X, .P(2).C3.Y, .P(2).C3.Z)

    ' Punto 3
    .P(3).C3.X = 1 * miTamañoCubo
    .P(3).C3.Y = 0 * miTamañoCubo
    .P(3).C3.Z = 0 * miTamañoCubo
    .P(3).A.Color = vbRed
    .P(3).A.Tamaño = 100
    .P(3).A.Visible = True
    .P(3).C2.X = CoordenadaXPantalla(.P(3).C3.X, .P(3).C3.Y, .P(3).C3.Z)
    .P(3).C2.Y = CoordenadaYPantalla(.P(3).C3.X, .P(3).C3.Y, .P(3).C3.Z)

    ' Punto 4
    .P(4).C3.X = 1 * miTamañoCubo
    .P(4).C3.Y = 1 * miTamañoCubo
    .P(4).C3.Z = 0 * miTamañoCubo
    .P(4).A.Color = vbYellow
    .P(4).A.Tamaño = 100
    .P(4).A.Visible = True
    .P(4).C2.X = CoordenadaXPantalla(.P(4).C3.X, .P(4).C3.Y, .P(4).C3.Z)
    .P(4).C2.Y = CoordenadaYPantalla(.P(4).C3.X, .P(4).C3.Y, .P(4).C3.Z)

    ' Punto 5
    .P(5).C3.X = 0 * miTamañoCubo
    .P(5).C3.Y = 0 * miTamañoCubo
    .P(5).C3.Z = 1 * miTamañoCubo
    .P(5).A.Color = vbYellow
    .P(5).A.Tamaño = 100
    .P(5).A.Visible = True
    .P(5).C2.X = CoordenadaXPantalla(.P(5).C3.X, .P(5).C3.Y, .P(5).C3.Z)
    .P(5).C2.Y = CoordenadaYPantalla(.P(5).C3.X, .P(5).C3.Y, .P(5).C3.Z)

    ' Punto 6
    .P(6).C3.X = 0 * miTamañoCubo
    .P(6).C3.Y = 1 * miTamañoCubo
    .P(6).C3.Z = 1 * miTamañoCubo
    .P(6).A.Color = vbGreen
    .P(6).A.Tamaño = 100
    .P(6).A.Visible = True
    .P(6).C2.X = CoordenadaXPantalla(.P(6).C3.X, .P(6).C3.Y, .P(6).C3.Z)
    .P(6).C2.Y = CoordenadaYPantalla(.P(6).C3.X, .P(6).C3.Y, .P(6).C3.Z)

    ' Punto 7
    .P(7).C3.X = 1 * miTamañoCubo
    .P(7).C3.Y = 0 * miTamañoCubo
    .P(7).C3.Z = 1 * miTamañoCubo
    .P(7).A.Color = vbBlue
    .P(7).A.Tamaño = 100
    .P(7).A.Visible = True
    .P(7).C2.X = CoordenadaXPantalla(.P(7).C3.X, .P(7).C3.Y, .P(7).C3.Z)
    .P(7).C2.Y = CoordenadaYPantalla(.P(7).C3.X, .P(7).C3.Y, .P(7).C3.Z)

    ' Punto 8
    .P(8).C3.X = 1 * miTamañoCubo
    .P(8).C3.Y = 1 * miTamañoCubo
    .P(8).C3.Z = 1 * miTamañoCubo
    .P(8).A.Color = vbRed
    .P(8).A.Tamaño = 100
    .P(8).A.Visible = True
    .P(8).C2.X = CoordenadaXPantalla(.P(8).C3.X, .P(8).C3.Y, .P(8).C3.Z)
    .P(8).C2.Y = CoordenadaYPantalla(.P(8).C3.X, .P(8).C3.Y, .P(8).C3.Z)

    ' Linea 01
    .L(1).P1.C2.X = .P(1).C2.X
    .L(1).P1.C2.Y = .P(1).C2.Y
    .L(1).P2.C2.X = .P(2).C2.X
    .L(1).P2.C2.Y = .P(2).C2.Y
    .L(1).A.Visible = True
    .L(1).A.Color = vbBlack

    ' Linea 02
    .L(2).P1.C2.X = .P(1).C2.X
    .L(2).P1.C2.Y = .P(1).C2.Y
    .L(2).P2.C2.X = .P(3).C2.X
    .L(2).P2.C2.Y = .P(3).C2.Y
    .L(2).A.Visible = True
    .L(2).A.Color = vbBlack

    ' Linea 03
    .L(3).P1.C2.X = .P(1).C2.X
    .L(3).P1.C2.Y = .P(1).C2.Y
    .L(3).P2.C2.X = .P(5).C2.X
    .L(3).P2.C2.Y = .P(5).C2.Y
    .L(3).A.Visible = True
    .L(3).A.Color = vbBlack

    ' Linea 04
    .L(4).P1.C2.X = .P(4).C2.X
    .L(4).P1.C2.Y = .P(4).C2.Y
    .L(4).P2.C2.X = .P(2).C2.X
    .L(4).P2.C2.Y = .P(2).C2.Y
    .L(4).A.Visible = True
    .L(4).A.Color = vbBlack

    ' Linea 05
    .L(5).P1.C2.X = .P(4).C2.X
    .L(5).P1.C2.Y = .P(4).C2.Y
    .L(5).P2.C2.X = .P(3).C2.X
    .L(5).P2.C2.Y = .P(3).C2.Y
    .L(5).A.Visible = True
    .L(5).A.Color = vbBlack

    ' Linea 06
    .L(6).P1.C2.X = .P(4).C2.X
    .L(6).P1.C2.Y = .P(4).C2.Y
    .L(6).P2.C2.X = .P(8).C2.X
    .L(6).P2.C2.Y = .P(8).C2.Y
    .L(6).A.Visible = True
    .L(6).A.Color = vbBlack

    ' Linea 07
    .L(7).P1.C2.X = .P(6).C2.X
    .L(7).P1.C2.Y = .P(6).C2.Y
    .L(7).P2.C2.X = .P(5).C2.X
    .L(7).P2.C2.Y = .P(5).C2.Y
    .L(7).A.Visible = True
    .L(7).A.Color = vbBlack

    ' Linea 08
    .L(8).P1.C2.X = .P(6).C2.X
    .L(8).P1.C2.Y = .P(6).C2.Y
    .L(8).P2.C2.X = .P(8).C2.X
    .L(8).P2.C2.Y = .P(8).C2.Y
    .L(8).A.Visible = True
    .L(8).A.Color = vbBlack

    ' Linea 09
    .L(9).P1.C2.X = .P(6).C2.X
    .L(9).P1.C2.Y = .P(6).C2.Y
    .L(9).P2.C2.X = .P(2).C2.X
    .L(9).P2.C2.Y = .P(2).C2.Y
    .L(9).A.Visible = True
    .L(9).A.Color = vbBlack

    ' Linea 10
    .L(10).P1.C2.X = .P(7).C2.X
    .L(10).P1.C2.Y = .P(7).C2.Y
    .L(10).P2.C2.X = .P(5).C2.X
    .L(10).P2.C2.Y = .P(5).C2.Y
    .L(10).A.Visible = True
    .L(10).A.Color = vbBlack

    ' Linea 11
    .L(11).P1.C2.X = .P(7).C2.X
    .L(11).P1.C2.Y = .P(7).C2.Y
    .L(11).P2.C2.X = .P(8).C2.X
    .L(11).P2.C2.Y = .P(8).C2.Y
    .L(11).A.Visible = True
    .L(11).A.Color = vbBlack

    ' Linea 12
    .L(12).P1.C2.X = .P(7).C2.X
    .L(12).P1.C2.Y = .P(7).C2.Y
    .L(12).P2.C2.X = .P(3).C2.X
    .L(12).P2.C2.Y = .P(3).C2.Y
    .L(12).A.Visible = True
    .L(12).A.Color = vbBlack

  End With


  ' Mostrar los puntos
  For i = 1 To 8
    Call MuestraPunto(miCubo.P(i))
  Next
  For i = 1 To 12
    Call MuestraLinea(miCubo.L(i))
  Next


  '''
  '''  ' Ajuste del Angulo
  '''  ang = Val(txtAngulo.Text) * (miPi / 180)
  '''  miTamañoCubo = 4000
  '''
  '''  ' Borrar la pantalla
  '''  ' Cls
  '''
  '''  ' Coordenadas del Centro
  '''  CentroX = 0
  '''  CentroY = frmObjetoEspacio.Width / 2
  '''  CentroZ = frmObjetoEspacio.Height / 2
  '''
  '''  With miCubo
  '''    ' Punto 1
  '''    .P(1).C3.X = 0 * miTamañoCubo
  '''    .P(1).C3.Y = 0 * miTamañoCubo
  '''    .P(1).C3.Z = 0 * miTamañoCubo
  '''    .P(1).A.Color = vbGreen
  '''    .P(1).A.Tamaño = 100
  '''    .P(1).A.Visible = True
  '''    .P(1).C2.X = CoordenadaXPantalla(.P(1).C3.X, .P(1).C3.Y, .P(1).C3.Z)
  '''    .P(1).C2.Y = CoordenadaYPantalla(.P(1).C3.X, .P(1).C3.Y, .P(1).C3.Z)
  '''
  '''    ' Punto 2
  '''    .P(2).C3.X = 0 * miTamañoCubo
  '''    .P(2).C3.Y = 1 * miTamañoCubo
  '''    .P(2).C3.Z = 0 * miTamañoCubo
  '''    .P(2).A.Color = vbBlue
  '''    .P(2).A.Tamaño = 100
  '''    .P(2).A.Visible = True
  '''    .P(2).C2.X = CoordenadaXPantalla(.P(2).C3.X, .P(2).C3.Y, .P(2).C3.Z)
  '''    .P(2).C2.Y = CoordenadaYPantalla(.P(2).C3.X, .P(2).C3.Y, .P(2).C3.Z)
  '''
  '''    ' Punto 3
  '''    .P(3).C3.X = 1 * miTamañoCubo
  '''    .P(3).C3.Y = 0 * miTamañoCubo
  '''    .P(3).C3.Z = 0 * miTamañoCubo
  '''    .P(3).A.Color = vbRed
  '''    .P(3).A.Tamaño = 100
  '''    .P(3).A.Visible = True
  '''    .P(3).C2.X = CoordenadaXPantalla(.P(3).C3.X, .P(3).C3.Y, .P(3).C3.Z)
  '''    .P(3).C2.Y = CoordenadaYPantalla(.P(3).C3.X, .P(3).C3.Y, .P(3).C3.Z)
  '''
  '''    ' Punto 4
  '''    .P(4).C3.X = 1 * miTamañoCubo
  '''    .P(4).C3.Y = 1 * miTamañoCubo
  '''    .P(4).C3.Z = 0 * miTamañoCubo
  '''    .P(4).A.Color = vbYellow
  '''    .P(4).A.Tamaño = 100
  '''    .P(4).A.Visible = True
  '''    .P(4).C2.X = CoordenadaXPantalla(.P(4).C3.X, .P(4).C3.Y, .P(4).C3.Z)
  '''    .P(4).C2.Y = CoordenadaYPantalla(.P(4).C3.X, .P(4).C3.Y, .P(4).C3.Z)
  '''
  '''    ' Punto 5
  '''    .P(5).C3.X = 0 * miTamañoCubo
  '''    .P(5).C3.Y = 0 * miTamañoCubo
  '''    .P(5).C3.Z = 1 * miTamañoCubo
  '''    .P(5).A.Color = vbYellow
  '''    .P(5).A.Tamaño = 100
  '''    .P(5).A.Visible = True
  '''    .P(5).C2.X = CoordenadaXPantalla(.P(5).C3.X, .P(5).C3.Y, .P(5).C3.Z)
  '''    .P(5).C2.Y = CoordenadaYPantalla(.P(5).C3.X, .P(5).C3.Y, .P(5).C3.Z)
  '''
  '''    ' Punto 6
  '''    .P(6).C3.X = 0 * miTamañoCubo
  '''    .P(6).C3.Y = 1 * miTamañoCubo
  '''    .P(6).C3.Z = 1 * miTamañoCubo
  '''    .P(6).A.Color = vbGreen
  '''    .P(6).A.Tamaño = 100
  '''    .P(6).A.Visible = True
  '''    .P(6).C2.X = CoordenadaXPantalla(.P(6).C3.X, .P(6).C3.Y, .P(6).C3.Z)
  '''    .P(6).C2.Y = CoordenadaYPantalla(.P(6).C3.X, .P(6).C3.Y, .P(6).C3.Z)
  '''
  '''    ' Punto 7
  '''    .P(7).C3.X = 1 * miTamañoCubo
  '''    .P(7).C3.Y = 0 * miTamañoCubo
  '''    .P(7).C3.Z = 1 * miTamañoCubo
  '''    .P(7).A.Color = vbBlue
  '''    .P(7).A.Tamaño = 100
  '''    .P(7).A.Visible = True
  '''    .P(7).C2.X = CoordenadaXPantalla(.P(7).C3.X, .P(7).C3.Y, .P(7).C3.Z)
  '''    .P(7).C2.Y = CoordenadaYPantalla(.P(7).C3.X, .P(7).C3.Y, .P(7).C3.Z)
  '''
  '''    ' Punto 8
  '''    .P(8).C3.X = 1 * miTamañoCubo
  '''    .P(8).C3.Y = 1 * miTamañoCubo
  '''    .P(8).C3.Z = 1 * miTamañoCubo
  '''    .P(8).A.Color = vbRed
  '''    .P(8).A.Tamaño = 100
  '''    .P(8).A.Visible = True
  '''    .P(8).C2.X = CoordenadaXPantalla(.P(8).C3.X, .P(8).C3.Y, .P(8).C3.Z)
  '''    .P(8).C2.Y = CoordenadaYPantalla(.P(8).C3.X, .P(8).C3.Y, .P(8).C3.Z)
  '''
  '''    ' Linea 01
  '''    .L(1).P1.C2.X = .P(1).C2.X
  '''    .L(1).P1.C2.Y = .P(1).C2.Y
  '''    .L(1).P2.C2.X = .P(2).C2.X
  '''    .L(1).P2.C2.Y = .P(2).C2.Y
  '''    .L(1).A.Visible = True
  '''    .L(1).A.Color = vbBlack
  '''
  '''    ' Linea 02
  '''    .L(2).P1.C2.X = .P(1).C2.X
  '''    .L(2).P1.C2.Y = .P(1).C2.Y
  '''    .L(2).P2.C2.X = .P(3).C2.X
  '''    .L(2).P2.C2.Y = .P(3).C2.Y
  '''    .L(2).A.Visible = True
  '''    .L(2).A.Color = vbBlack
  '''
  '''    ' Linea 03
  '''    .L(3).P1.C2.X = .P(1).C2.X
  '''    .L(3).P1.C2.Y = .P(1).C2.Y
  '''    .L(3).P2.C2.X = .P(5).C2.X
  '''    .L(3).P2.C2.Y = .P(5).C2.Y
  '''    .L(3).A.Visible = True
  '''    .L(3).A.Color = vbBlack
  '''
  '''    ' Linea 04
  '''    .L(4).P1.C2.X = .P(4).C2.X
  '''    .L(4).P1.C2.Y = .P(4).C2.Y
  '''    .L(4).P2.C2.X = .P(2).C2.X
  '''    .L(4).P2.C2.Y = .P(2).C2.Y
  '''    .L(4).A.Visible = True
  '''    .L(4).A.Color = vbBlack
  '''
  '''    ' Linea 05
  '''    .L(5).P1.C2.X = .P(4).C2.X
  '''    .L(5).P1.C2.Y = .P(4).C2.Y
  '''    .L(5).P2.C2.X = .P(3).C2.X
  '''    .L(5).P2.C2.Y = .P(3).C2.Y
  '''    .L(5).A.Visible = True
  '''    .L(5).A.Color = vbBlack
  '''
  '''    ' Linea 06
  '''    .L(6).P1.C2.X = .P(4).C2.X
  '''    .L(6).P1.C2.Y = .P(4).C2.Y
  '''    .L(6).P2.C2.X = .P(8).C2.X
  '''    .L(6).P2.C2.Y = .P(8).C2.Y
  '''    .L(6).A.Visible = True
  '''    .L(6).A.Color = vbBlack
  '''
  '''    ' Linea 07
  '''    .L(7).P1.C2.X = .P(6).C2.X
  '''    .L(7).P1.C2.Y = .P(6).C2.Y
  '''    .L(7).P2.C2.X = .P(5).C2.X
  '''    .L(7).P2.C2.Y = .P(5).C2.Y
  '''    .L(7).A.Visible = True
  '''    .L(7).A.Color = vbBlack
  '''
  '''    ' Linea 08
  '''    .L(8).P1.C2.X = .P(6).C2.X
  '''    .L(8).P1.C2.Y = .P(6).C2.Y
  '''    .L(8).P2.C2.X = .P(8).C2.X
  '''    .L(8).P2.C2.Y = .P(8).C2.Y
  '''    .L(8).A.Visible = True
  '''    .L(8).A.Color = vbBlack
  '''
  '''    ' Linea 09
  '''    .L(9).P1.C2.X = .P(6).C2.X
  '''    .L(9).P1.C2.Y = .P(6).C2.Y
  '''    .L(9).P2.C2.X = .P(2).C2.X
  '''    .L(9).P2.C2.Y = .P(2).C2.Y
  '''    .L(9).A.Visible = True
  '''    .L(9).A.Color = vbBlack
  '''
  '''    ' Linea 10
  '''    .L(10).P1.C2.X = .P(7).C2.X
  '''    .L(10).P1.C2.Y = .P(7).C2.Y
  '''    .L(10).P2.C2.X = .P(5).C2.X
  '''    .L(10).P2.C2.Y = .P(5).C2.Y
  '''    .L(10).A.Visible = True
  '''    .L(10).A.Color = vbBlack
  '''
  '''    ' Linea 11
  '''    .L(11).P1.C2.X = .P(7).C2.X
  '''    .L(11).P1.C2.Y = .P(7).C2.Y
  '''    .L(11).P2.C2.X = .P(8).C2.X
  '''    .L(11).P2.C2.Y = .P(8).C2.Y
  '''    .L(11).A.Visible = True
  '''    .L(11).A.Color = vbBlack
  '''
  '''    ' Linea 12
  '''    .L(12).P1.C2.X = .P(7).C2.X
  '''    .L(12).P1.C2.Y = .P(7).C2.Y
  '''    .L(12).P2.C2.X = .P(3).C2.X
  '''    .L(12).P2.C2.Y = .P(3).C2.Y
  '''    .L(12).A.Visible = True
  '''    .L(12).A.Color = vbBlack
  '''
  '''  End With
  '''
  '''
  '''  ' Mostrar los puntos
  '''  For i = 1 To 8
  '''    Call MuestraPunto(miCubo.P(i))
  '''  Next
  '''  For i = 1 To 12
  '''    Call MuestraLinea(miCubo.L(i))
  '''  Next
  '''

  For i = 1 To 8

    With miCubo
      ' Punto 1
      .P(i).C3.X = 0 * miTamañoCubo
      .P(i).C3.Y = 0 * miTamañoCubo
      .P(i).C3.Z = 0 * miTamañoCubo
      .P(i).C2.X = CoordenadaXPantalla(.P(1).C3.X, .P(1).C3.Y, .P(1).C3.Z)
      .P(i).C2.Y = CoordenadaYPantalla(.P(1).C3.X, .P(1).C3.Y, .P(1).C3.Z)

      ' Punto de prueba
      Punto(1, 1) = .P(i).C3.X
      Punto(1, 2) = .P(i).C3.Y
      Punto(1, 3) = .P(i).C3.Z
      Punto(1, 4) = 1
    End With

    ' Aplica la Transformacion
    Call AplicaTransformacion(Tz, Punto)

    With miCubo
      ' Punto 1
      .P(i).C3.X = Resultado(1, 1) * miTamañoCubo
      .P(i).C3.Y = Resultado(1, 2) * miTamañoCubo
      .P(i).C3.Z = Resultado(1, 3) * miTamañoCubo
      .P(i).C2.X = CoordenadaXPantalla(.P(1).C3.X, .P(1).C3.Y, .P(1).C3.Z)
      .P(i).C2.Y = CoordenadaYPantalla(.P(1).C3.X, .P(1).C3.Y, .P(1).C3.Z)
    End With

  Next

  'Cls
  ' Mostrar los puntos
  For i = 1 To 8
    Call MuestraPunto(miCubo.P(i))
  Next
  For i = 1 To 12
    Call MuestraLinea(miCubo.L(i))
  Next


End Sub

' Muestra un punto en Pantalla
Private Sub MuestraPunto(ByRef pP As Puntos)
' Mostrar el punto con control del Tamaño y Color
  Dim radio As Double
  If pP.A.Visible = True Then
    If pP.A.Tamaño <= 0 Then
      PSet (pP.C2.X, pP.C2.Y), pP.A.Color
    Else
      For radio = 1 To pP.A.Tamaño
        Circle (pP.C2.X, pP.C2.Y), radio, pP.A.Color
      Next
    End If
  End If
End Sub

' Muestra una linea en pantalla
Private Sub MuestraLinea(ByRef pL As Lineas)
  Line (pL.P1.C2.X, pL.P1.C2.Y)-(pL.P2.C2.X, pL.P2.C2.Y)
End Sub

' Aplica la tranaformación a un punto
Private Sub cmdTransformacion_Click()
' Lectura de Angulos de rotación
  alfa = Val(txtAlfa.Text) * grados
  beta = Val(txtBeta.Text) * grados
  teta = Val(txtTeta.Text) * grados

  ' Lectura Traslación
  Traslacion(1, 1) = Val(txtDx.Text)
  Traslacion(1, 2) = Val(txtDy.Text)
  Traslacion(1, 3) = Val(txtDz.Text)

  ' Lectura Perspectiva
  Perspectiva(1) = Val(txtPx.Text)
  Perspectiva(2) = Val(txtPy.Text)
  Perspectiva(3) = Val(txtPz.Text)

  ' Lectura Escalado
  Escala = Val(txtEscalado.Text)

  ' Transformacion Rotación X con Alfa
  Tx(1, 1) = 1: Tx(1, 2) = 0: Tx(1, 3) = 0: Tx(1, 4) = Traslacion(1, 1)
  Tx(2, 1) = 0: Tx(2, 2) = Cos(alfa): Tx(2, 3) = -Sin(alfa): Tx(2, 4) = Traslacion(1, 2)
  Tx(3, 1) = 0: Tx(3, 2) = Sin(alfa): Tx(3, 3) = Cos(alfa): Tx(3, 4) = Traslacion(1, 3)
  Tx(4, 1) = Perspectiva(1): Tx(4, 2) = Perspectiva(2): Tx(4, 3) = Perspectiva(3): Tx(4, 4) = Escala

  ' Transformacion Rotación Y con Beta
  Ty(1, 1) = Cos(beta): Ty(1, 2) = 0: Ty(1, 3) = Sin(beta): Ty(1, 4) = Traslacion(1, 1)
  Ty(2, 1) = 0: Ty(2, 2) = 1: Ty(2, 3) = 0: Ty(2, 4) = Traslacion(1, 2)
  Ty(3, 1) = -Sin(beta): Ty(3, 2) = 0: Ty(3, 3) = Cos(beta): Ty(3, 4) = Traslacion(1, 3)
  Ty(4, 1) = Perspectiva(1): Ty(4, 2) = Perspectiva(2): Ty(4, 3) = Perspectiva(3): Ty(4, 4) = Escala

  ' Transformacion Rotación Z con Teta
  Tz(1, 1) = Cos(teta): Tz(1, 2) = -Sin(teta): Tz(1, 3) = 0: Tz(1, 4) = Traslacion(1, 1)
  Tz(2, 1) = Sin(teta): Tz(2, 2) = Cos(teta): Tz(2, 3) = 0: Tz(2, 4) = Traslacion(1, 2)
  Tz(3, 1) = 0: Tz(3, 2) = 0: Tz(3, 3) = 1: Tz(3, 4) = Traslacion(1, 3)
  Tz(4, 1) = Perspectiva(1): Tz(4, 2) = Perspectiva(2): Tz(4, 3) = Perspectiva(3): Tz(4, 4) = Escala

  ' Punto de prueba
  Punto(1, 1) = 1
  Punto(1, 2) = 2
  Punto(1, 3) = 3
  Punto(1, 4) = 1

  ' Aplica la Transformacion
  Call AplicaTransformacion(Tz, Punto)

End Sub




Public Sub AplicaTransformacion(ByRef pT() As Double, ByRef pP() As Double)
  Dim r As Integer, igual As Integer
  Dim acumula As Double
  ' Recorre la Matriz Producto
  For r = 1 To 4
    For igual = 1 To 4
      acumula = acumula + (pT(r, igual) * pP(1, igual))
    Next igual
    Resultado(1, r) = acumula
    acumula = 0
  Next r

  Print ""
  Print Resultado(1, 1)
  Print Resultado(1, 2)
  Print Resultado(1, 3)
  Print Resultado(1, 4)
End Sub

