VERSION 5.00
Begin VB.Form frmEspacio3D 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   " .:. ESPACIO EN 3D .:."
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   14175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGira 
      Caption         =   "Gira"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.CheckBox chkPlanoXY 
      BackColor       =   &H00000000&
      Caption         =   "Plano XY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CheckBox chkGrillasCubicas 
      BackColor       =   &H00000000&
      Caption         =   "Grillas Cúbicas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CheckBox chkEsfera 
      BackColor       =   &H00000000&
      Caption         =   "Esfera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CheckBox chkEjes 
      BackColor       =   &H00000000&
      Caption         =   "Ejes de Coordenadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox chkVectorPosicion 
      BackColor       =   &H00000000&
      Caption         =   "Vector Posición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CheckBox chkCuadricula 
      BackColor       =   &H00000000&
      Caption         =   "Cuadricula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CheckBox ckkEjesParalelos 
      BackColor       =   &H00000000&
      Caption         =   "Ejes Paralelos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CheckBox chkMarcasEjes 
      BackColor       =   &H00000000&
      Caption         =   "Marcas en Ejes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdMoverPunto 
      Caption         =   "Mover Punto"
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
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdGirar 
      Caption         =   "Gira 0 - 360"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
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
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrafica 
      Caption         =   "Graficar"
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
End
Attribute VB_Name = "frmEspacio3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Declaración de variables
Dim CentroX As Double
Dim CentroY As Double
Dim CentroZ As Double

' Declaración de Constantes
Const miPi = 3.14159265358979

Private Sub cmdGira_Click()
' Declaración de variables
  Dim a As Double
  ' Rotación de 360 grados
  txtAngulo.Text = Val(txtAngulo.Text) + 1
  ' Muestra la Grafica
  Call cmdGrafica_Click
  ' Muestra la CuadriculaXY
  If chkCuadricula.Value = 1 Then
    Call CuadriculaXY
  End If
  ' Muestra el Plano XY
  If chkPlanoXY.Value = 1 Then
    Call PlanoXY
  End If
  DoEvents
End Sub

' Al cargar el formulario
Private Sub Form_Load()

End Sub

' Genera la Grafica
Private Sub cmdGrafica_Click()
' Generar el gráfico
  Call Graficar
  DoEvents
End Sub

' Gira el Eje X
Private Sub cmdGirar_Click()
' Declaración de variables
  Dim a As Double
  ' Rotación de 360 grados
  For a = 1 To 360
    txtAngulo = a
    ' Muestra la Grafica
    Call cmdGrafica_Click
    ' Muestra la CuadriculaXY
    If chkCuadricula.Value = 1 Then
      Call CuadriculaXY
    End If
    ' Muestra el Plano XY
    If chkPlanoXY.Value = 1 Then
      Call PlanoXY
    End If
    DoEvents
  Next
End Sub

' Mover un punto por la gáfica
Private Sub cmdMoverPunto_Click()
' Declaración de variables
  Dim posX As Double
  Dim posY As Double
  Dim posZ As Double
  Dim x As Double
  Dim ang As Double
  ' Inicializacón de variables
  posX = 2500
  posY = 2500
  posZ = 2500
  ' Movimiento circular uniforme
  For ang = 0 To 360 * 5 Step 0.5
    ' Borrar la pantalla
    Cls
    ' Coordenadas del Centro
    CentroX = 0
    CentroY = frmEspacio3D.Width / 2
    CentroZ = frmEspacio3D.Height / 2
    ' Ejes
    If chkEjes.Value = 1 Then
      Call PL3D(-5000, 0, 0, 5000, 0, 0, 15, vbGreen)
      Call PL3D(-5000, 0, 20, 5000, 0, 20, 15, vbGreen)
      Call PL3D(0, -5000, 0, 0, 5000, 0, 15, vbRed)
      Call PL3D(20, -5000, 0, 20, 5000, 0, 15, vbRed)
      Call PL3D(0, 0, -5000, 0, 0, 5000, 15, vbBlue)
      Call PL3D(0, 20, -5000, 0, 20, 5000, 15, vbBlue)
    End If
    ' Puntos de prueba marca en los ejes
    If chkMarcasEjes.Value = 1 Then
      Call PP3D(0, 0, 0, 50, vbWhite)
      Call PP3D(posX * Sin(ang * miPi / 180), 0, 0, 50, vbGreen)
      Call PP3D(0, posY * Cos(ang * miPi / 180), 0, 50, vbRed)
      Call PP3D(0, 0, posZ * -Sin(ang * miPi / 180), 50, vbBlue)
    End If
    ' Muestra la CuadriculaXY
    If chkCuadricula.Value = 1 Then
      Call CuadriculaXY
    End If
    ' Dibuja el plano xy
    If chkPlanoXY.Value = 1 Then
      Call PlanoXY
    End If

    ' El punto graficado
    If posZ * -Sin(ang * miPi / 180) > 0 Then
      Call PP3D(posX * Sin(ang * miPi / 180), posY * Cos(ang * miPi / 180), posZ * -Sin(ang * miPi / 180), 250, vbWhite)
    Else
      Call PP3D(posX * Sin(ang * miPi / 180), posY * Cos(ang * miPi / 180), posZ * -Sin(ang * miPi / 180), 250, RGB(100, 100, 100))
    End If

    ' Paralelos a los ejes Punteados
    If ckkEjesParalelos.Value = 1 Then
      Call PL3D(-5000, posY * Cos(ang * miPi / 180), 0, 5000, posY * Cos(ang * miPi / 180), 0, 15, vbYellow)
      Call PL3D(posX * Sin(ang * miPi / 180), -5000, 0, posX * Sin(ang * miPi / 180), 5000, 0, 15, vbYellow)
      Call PL3D(posX * Sin(ang * miPi / 180), posY * Cos(ang * miPi / 180), -5000, posX * Sin(ang * miPi / 180), posY * Cos(ang * miPi / 180), 5000, 15, vbYellow)
    End If
    DoEvents
  Next
End Sub

' Crea una grafica
Private Sub Graficar()
' Declaración de variable
  Dim x As Double
  Dim y As Double
  Dim z As Double
  Dim d As Double
  ' Borrar la pantalla
  Cls
  ' Coordenadas del Centro
  CentroX = 0
  CentroY = frmEspacio3D.Width / 2
  CentroZ = frmEspacio3D.Height / 2
  ' Ejes
  If chkEjes.Value = 1 Then
    Call PL3D(-5000, 0, 0, 5000, 0, 0, 15, vbGreen)
    Call PL3D(0, -5000, 0, 0, 5000, 0, 15, vbRed)
    Call PL3D(0, 0, -5000, 0, 0, 5000, 15, vbBlue)
  End If
  ' Muestra la CuadriculaXY
  If chkCuadricula.Value = 1 Then
    Call CuadriculaXY
  End If
  ' Muestra el Plano XY
  If chkPlanoXY.Value = 1 Then
    Call PlanoXY
  End If
  ' Paralelos a los ejes Punteados
  If ckkEjesParalelos.Value = 1 Then
    For x = -5000 To 5000 Step 100
      Call PP3D(x, 2000, 0, 0, vbGreen)
      Call PP3D(1000, x, 0, 0, vbRed)
      Call PP3D(1000, 2000, x, 0, vbBlue)
    Next
  End If
  ' Puntos de prueba marca en los ejes
  If chkMarcasEjes.Value = 1 Then
    Call PP3D(0, 0, 0, 50, vbWhite)
    Call PP3D(1000, 0, 0, 50, vbGreen)
    Call PP3D(0, 2000, 0, 50, vbRed)
    Call PP3D(0, 0, 3000, 50, vbBlue)
  End If
  ' El punto graficado
  Call PP3D(1000, 2000, 3000, 50, vbYellow)

  Dim orbitas As Integer
  orbitas = 20
  ' Esfera
  If chkEsfera.Value = 1 Then
    For x = -orbitas To orbitas
      For y = -orbitas To orbitas
        For z = -orbitas To orbitas
          If z Mod 2 = 0 Then
            d = Sqr((x * 150 * 1.9) ^ 2 + (y * 150) ^ 2 + (z * 150) ^ 2)
            If d > 3000 And d < 3150 Then

              Call PP3D(x * 255, y * 255, ((z + 1) * 255) / 1.1, 15, QBColor(Int(Abs(z + 1) Mod 15)))  'vbWhite
              'Call PP3D(x * 300, y * 300, ((z + 1) * 300) / 1.1, 20, vbGreen) 'vbWhite

            End If
          End If
        Next z
      Next y
    Next x
  End If

  ' Grillas cúbicas
  If chkGrillasCubicas.Value = 1 Then
    Dim escala As Double
    Dim tamaño As Integer
    Dim densidad As Integer
    ' Ajuste de Figura
    escala = 250
    tamaño = 0
    densidad = 10
    ' Dibuja la figura
    For x = 1 To densidad
      For y = 1 To densidad
        For z = 1 To densidad
          Call PP3D(x * escala, y * escala, z * escala, tamaño, vbWhite)
        Next z
      Next y
    Next x
  End If
End Sub

' Grafica un punto en espacio vectorial de tres dimensiones 3D
Public Sub PP3D(ByVal pX As Double, ByVal pY As Double, ByVal pZ As Double, pTamaño As Integer, ByVal pColor As Long)
' Declaración de variables
  Dim ppX As Double
  Dim ppY As Double
  Dim ang As Double
  Dim Radio As Double
  ' Ajuste del Angulo
  ang = Val(txtAngulo.Text) * (miPi / 180)
  ' Coordenadas de Pantalla del Punto
  ppX = CentroY + (-pX * Cos(ang)) + (pY) + (0)
  ppY = CentroZ + (pX * Sin(ang)) + (0) + (-pZ)
  ' Mostrar el punto con control del Tamaño y Color
  If pTamaño <= 0 Then
    PSet (ppX, ppY), pColor
  Else
    For Radio = 1 To pTamaño
      Circle (ppX, ppY), Radio, pColor
    Next
  End If
  ' Muestra el vector posición
  If pColor <> vbRed And pColor <> vbGreen And pColor <> vbBlue Then
    Call VectorPosicion(CentroY, CentroZ, ppX, ppY)
  End If
End Sub

' Muestra el vector posición
Public Sub VectorPosicion(ByVal pCX As Double, ByVal pCY As Double, ByVal pX As Double, ByVal pY As Double)
' Muestra el vector posición
  If chkVectorPosicion.Value = 1 Then
    Line (pCX, pCY)-(pX, pY), vbWhite
  End If
End Sub

' Grafica una línea en espacio vectorial de tres dimensiones 3D
Public Sub PL3D(ByVal pX1 As Double, ByVal pY1 As Double, ByVal pZ1 As Double, ByVal pX2 As Double, ByVal pY2 As Double, ByVal pZ2 As Double, pTamaño As Integer, ByVal pColor As Long)
' Declaración de variables
  Dim ppX1 As Double
  Dim ppY1 As Double
  Dim ppX2 As Double
  Dim ppY2 As Double
  Dim ang As Double
  Dim Radio As Double
  ' Ajuste del Angulo
  ang = Val(txtAngulo.Text) * (miPi / 180)
  ' Coordenadas de Pantalla del Punto1
  ppX1 = CentroY + (-pX1 * Cos(ang)) + (pY1) + (0)
  ppY1 = CentroZ + (pX1 * Sin(ang)) + (0) + (-pZ1)
  ' Coordenadas de Pantalla del Punto2
  ppX2 = CentroY + (-pX2 * Cos(ang)) + (pY2) + (0)
  ppY2 = CentroZ + (pX2 * Sin(ang)) + (0) + (-pZ2)
  ' Mostrar el punto con control del Tamaño y Color
  Line (ppX1, ppY1)-(ppX2, ppY2), pColor
End Sub

' Muestra el plano XY
Private Sub PlanoXY()
' Declaración de variables
  Dim x As Double
  Dim y As Double
  Dim z As Double
  Dim cuadricula As Double
  ' Coordenadas del Centro
  CentroX = 0
  CentroY = frmEspacio3D.Width / 2
  CentroZ = frmEspacio3D.Height / 2
  ' Parametros
  cuadricula = 1000
  ' Muestra lineas del Plano XY
  For y = -5000 To 5000 Step cuadricula
    Call PL3D(-5000, y, 0, 5000, y, 0, 0, vbGreen)
  Next
  ' Muestra lineas del Plano XY
  For x = -5000 To 5000 Step cuadricula
    Call PL3D(x, -5000, 0, x, 5000, 0, 0, vbRed)
  Next
End Sub

' Muestra el cuadricula XY
Private Sub CuadriculaXY()
' Declaración de variables
  Dim x As Double
  Dim y As Double
  Dim z As Double
  Dim cuadricula As Double
  ' Coordenadas del Centro
  CentroX = 0
  CentroY = frmEspacio3D.Width / 2
  CentroZ = frmEspacio3D.Height / 2
  ' Parametros
  cuadricula = 200
  ' Muestra lineas del Plano XY
  For y = -5000 To 5000 Step cuadricula
    Call PL3D(-5000, y, 0, 5000, y, 0, 0, RGB(100, 100, 100))
  Next
  ' Muestra lineas del Plano XY
  For x = -5000 To 5000 Step cuadricula
    Call PL3D(x, -5000, 0, x, 5000, 0, 0, RGB(100, 100, 100))
  Next
End Sub

