VERSION 5.00
Begin VB.Form frmGrafo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   Caption         =   "GRAFO"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1052
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      BackColor       =   &H8000000D&
      Caption         =   "Guardar imagen "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8640
      TabIndex        =   60
      Top             =   6600
      Width           =   2175
      Begin VB.CommandButton cmdGuardarGrafo 
         Caption         =   "Guardar Grafo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtArchivo 
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
         TabIndex        =   61
         Text            =   "Grafo.bmp"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H8000000D&
      Caption         =   "Movimientos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   11040
      TabIndex        =   53
      Top             =   6600
      Width           =   2175
      Begin VB.CommandButton cmdDZ 
         Caption         =   "D Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   59
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdDY 
         Caption         =   "D Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   58
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdDX 
         Caption         =   "D X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   57
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdRZ 
         Caption         =   "R Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   56
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdRY 
         Caption         =   "R Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   55
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdRX 
         Caption         =   "R X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H8000000D&
      Caption         =   "Comandos "
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
      Left            =   13440
      TabIndex        =   49
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdCalcularVertices 
         Caption         =   "Calcular Vértices"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   50
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H8000000D&
      Caption         =   "Modelos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   13440
      TabIndex        =   36
      Top             =   4560
      Width           =   2175
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H008080FF&
         Caption         =   "Mod. 01"
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
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mod. 02"
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
         Index           =   2
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mod. 03"
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
         Index           =   3
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mod. 04"
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
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H008080FF&
         Caption         =   "Mod. 05"
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
         Index           =   5
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mod. 06"
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
         Index           =   6
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mod. 07"
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
         Index           =   7
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H008080FF&
         Caption         =   "Mod. 08"
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
         Index           =   8
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mod. 09"
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
         Index           =   9
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mod. 10"
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
         Index           =   10
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mod. 11"
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
         Index           =   11
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdModelo 
         BackColor       =   &H008080FF&
         Caption         =   "Mod. 12"
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
         Index           =   12
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3360
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000D&
      Caption         =   "Comandos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8640
      TabIndex        =   33
      Top             =   4560
      Width           =   4575
      Begin VB.CommandButton cmdGrafoTeseractoModelos 
         Caption         =   "Grafo Teseracto Modelos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   52
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrafoCircularModelos 
         Caption         =   "Grafo Circular Modelos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   51
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrafoCircularVacio 
         Caption         =   "Grafo Circular Vacío"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrafoTeseractoVacio 
         Caption         =   "Grafo Teseracto Vacío"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.TextBox txtEtiqueta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   15360
      TabIndex        =   30
      Top             =   8400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Área de Grafo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   8295
      Begin VB.PictureBox picGrafo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   8000
         Left            =   120
         ScaleHeight     =   7965
         ScaleWidth      =   7965
         TabIndex        =   29
         Top             =   240
         Width           =   8000
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000D&
      Caption         =   "Otros Parametros "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   13440
      TabIndex        =   24
      Top             =   1440
      Width           =   2175
      Begin VB.CheckBox chcCircunferencia 
         BackColor       =   &H8000000D&
         Caption         =   "Circunferencia"
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox chcEjes 
         BackColor       =   &H8000000D&
         Caption         =   "Ejes"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chcEtiquetas 
         BackColor       =   &H8000000D&
         Caption         =   "Etiquetas"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtColorFondo 
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
         Left            =   1320
         TabIndex        =   25
         Text            =   "7"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Caption         =   "Color de Fondo:"
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
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000D&
      Caption         =   "Factores "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   11040
      TabIndex        =   15
      Top             =   1440
      Width           =   2175
      Begin VB.TextBox txtFactorCircular 
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
         Left            =   1320
         TabIndex        =   19
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtFactorExterno 
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
         Left            =   1320
         TabIndex        =   18
         Text            =   "1.2"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtFactorInterno 
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
         Left            =   1320
         TabIndex        =   17
         Text            =   "1.2"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtFactorProfundidad 
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
         Left            =   1320
         TabIndex        =   16
         Text            =   "200"
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         Caption         =   "Factor Circular:"
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
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   "Factor Externo:"
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
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "Factor  Interno:"
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
         TabIndex        =   21
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "Factor Profundidad:"
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
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Caption         =   "Vértices y Aristas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   8640
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
      Begin VB.TextBox txtColorAristas 
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
         Left            =   1320
         TabIndex        =   13
         Text            =   "7"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtColorVertices 
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
         Left            =   1320
         TabIndex        =   11
         Text            =   "15"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtFormaVertices 
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
         Left            =   1320
         TabIndex        =   9
         Text            =   "2"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtTamañoVertices 
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
         Left            =   1320
         TabIndex        =   7
         Text            =   "180"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "Color de Aristas: "
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
         TabIndex        =   14
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   "Color de Vértices: "
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
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "Forma de Vértices: "
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
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Tamaño de Vértices: "
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
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      Caption         =   "Número de Aristas "
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
      Left            =   11040
      TabIndex        =   3
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtNumeroAristas 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Número de Aristas: "
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
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Caption         =   "Número de Vértices "
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
      Left            =   8640
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtNumeroVertices 
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
         Left            =   1320
         TabIndex        =   1
         Text            =   "16"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Número de Vértices: "
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
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGrafo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : GRAFO
'* CONTENIDO     : PROYECTO PARA CREAR GRAFOS DEL SUDOKU
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO
'* INICIO        : 19 DE FEBRERO DE 2017
'* ACTUALIZACION : 19 DE FEBRERO DE 2017
'****************************************************************************************
Option Explicit

'AL CARGAR EL FORMULARIO
Private Sub Form_Load()
' Carga todos los valores iniciales
  miPi = 3.1415926535
  miRadio = 3750
  ' Mueve el centro de coordenadas al centro del picturebox
  picGrafo.ScaleTop = picGrafo.ScaleHeight / 2
  picGrafo.ScaleLeft = picGrafo.ScaleWidth / 2
  picGrafo.ScaleHeight = picGrafo.ScaleHeight * (-1)
  picGrafo.ScaleWidth = picGrafo.ScaleWidth * (-1)
End Sub

' CALCULAR VERTICES Y ARISTAS
Private Sub cmdCalcularVertices_Click()
' Ajusta los Parametros
  Call AjustarParametros

  ' Calcula la cantidad de aristas
  miNumeroVertices = Val(txtNumeroVertices.Text)
  miNumeroAristas = (miNumeroVertices * (miNumeroVertices - 1)) / 2
  txtNumeroAristas.Text = miNumeroAristas

  ' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
End Sub

' MOSTRAR GRAFO CIRCULAR VACIO
Private Sub cmdGrafoCircularVacio_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  ' Ubicar los Vertices
  Dim v As Integer
  For v = 1 To miNumeroVertices
    Call CarcularCircunferencia(oV(v), v)
  Next v
  ' Mostrar el Grafo
  picGrafo.Cls
  Call MostrarGrafo
End Sub

' MOSTRAR GRAFO TESERACTO VACIO
Private Sub cmdGrafoTeseractoVacio_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  miVec(1) = 1: miVec(2) = 7: miVec(3) = 14: miVec(4) = 12
  miVec(5) = 10: miVec(6) = 16: miVec(7) = 3: miVec(8) = 5
  miVec(9) = 8: miVec(10) = 2: miVec(11) = 11: miVec(12) = 13
  miVec(13) = 15: miVec(14) = 9: miVec(15) = 6: miVec(16) = 4
  Call CalcularTeseracto(miVec())
  ' Mostrar el Grafo
  picGrafo.Cls
  Call MostrarGrafo
End Sub

' MOSTRAR UN MODELO
Private Sub cmdModelo_Click(Index As Integer)
  Select Case Index
  Case 1
    Call ResaltarModelo(1, 1, 7, 10, 16)
  Case 2
    Call ResaltarModelo(2, 1, 7, 12, 14)
  Case 3
    Call ResaltarModelo(3, 1, 8, 10, 15)
  Case 4
    Call ResaltarModelo(4, 2, 7, 9, 16)
  Case 5
    Call ResaltarModelo(5, 2, 8, 9, 15)
  Case 6
    Call ResaltarModelo(6, 2, 8, 11, 13)
  Case 7
    Call ResaltarModelo(7, 3, 5, 10, 16)
  Case 8
    Call ResaltarModelo(8, 3, 5, 12, 14)
  Case 9
    Call ResaltarModelo(9, 3, 6, 11, 14)
  Case 10
    Call ResaltarModelo(10, 4, 5, 12, 13)
  Case 11
    Call ResaltarModelo(11, 4, 6, 9, 15)
  Case 12
    Call ResaltarModelo(12, 4, 6, 11, 13)
  End Select
  ' Mostrar el Grafo
  Call MostrarGrafo
End Sub

' MUESTRA LOS MODELOS EN FORMA CIRCULAR
Private Sub cmdGrafoCircularModelos_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  ' Ubicar los Vertices
  Dim v As Integer
  For v = 1 To miNumeroVertices
    Call CarcularCircunferencia(oV(v), v)
  Next v
  ' Resalta todos los modelos
  Call ResaltaTodo
  ' Mostrar el Grafo
  picGrafo.Cls
  Call MostrarGrafo
End Sub

' MUESTRA LOS MODELOS EN EL TESERACTO
Private Sub cmdGrafoTeseractoModelos_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  'Dim miVec(16) As Integer
  miVec(1) = 1: miVec(2) = 7: miVec(3) = 14: miVec(4) = 12
  miVec(5) = 10: miVec(6) = 16: miVec(7) = 3: miVec(8) = 5
  miVec(9) = 8: miVec(10) = 2: miVec(11) = 11: miVec(12) = 13
  miVec(13) = 15: miVec(14) = 9: miVec(15) = 6: miVec(16) = 4
  Call CalcularTeseracto(miVec())
  ' Mostrar el Grafo
  picGrafo.Cls
  ' Resalta todos los modelos
  Call ResaltaTodo
End Sub

' GUARDAR EL GRAFO
Private Sub cmdGuardarGrafo_Click()
' Ajusta el tamaño del formulario
  Me.Height = 9120
  Me.Width = 8670
  ' Borra el portapapeles
  Clipboard.Clear
  DoEvents
  ' Manda la pulsación de teclas para capturar la imagen de la pantalla
  On Error Resume Next
  Call keybd_event(&H2C, 1, 0, 0)
  DoEvents
  SavePicture Clipboard.GetData(vbCFBitmap), txtArchivo.Text
  DoEvents
  'SavePicture Clipboard.GetData(vbCFBitmap), Trim(txtRuta.Text & txtArchivo.Text)
  'DoEvents

  ' Ajusta el tamaño del formulario
  Me.Height = 9120
  Me.Width = 15900
End Sub

' ROTACION EN X
Private Sub cmdRX_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  Call CalcularRotacionX
  Call CalcularTeseracto(miVec())
  ' Mostrar el Grafo
  picGrafo.Cls
  ' Resalta todos los modelos
  Call ResaltaTodo
  Call chcEtiquetas_Click
  Call chcEtiquetas_Click
  ' Mostrar el Grafo
  'picGrafo.Cls
  'Call MostrarGrafo
End Sub

' ROTACION EN Y
Private Sub cmdRY_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  Call CalcularRotacionY
  Call CalcularTeseracto(miVec())
  ' Mostrar el Grafo
  picGrafo.Cls
  ' Resalta todos los modelos
  Call ResaltaTodo
  Call chcEtiquetas_Click
  Call chcEtiquetas_Click
  ' Mostrar el Grafo
  'picGrafo.Cls
  'Call MostrarGrafo
End Sub

' ROTACION EN Z
Private Sub cmdRZ_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  Call CalcularRotacionZ
  Call CalcularTeseracto(miVec())
  ' Mostrar el Grafo
  picGrafo.Cls
  ' Resalta todos los modelos
  Call ResaltaTodo
  Call chcEtiquetas_Click
  Call chcEtiquetas_Click
  ' Mostrar el Grafo
  'picGrafo.Cls
  'Call MostrarGrafo
End Sub

' DESPLAZAMIENTO EN X
Private Sub cmdDX_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  Call CalcularDespalzamientoX
  Call CalcularTeseracto(miVec())
  ' Mostrar el Grafo
  picGrafo.Cls
  ' Resalta todos los modelos
  Call ResaltaTodo
  Call chcEtiquetas_Click
  Call chcEtiquetas_Click
  ' Mostrar el Grafo
  'picGrafo.Cls
  'Call MostrarGrafo
End Sub

' DESPALZAMIENTO EN Y
Private Sub cmdDY_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  Call CalcularDespalzamientoY
  Call CalcularTeseracto(miVec())
  ' Mostrar el Grafo
  picGrafo.Cls
  ' Resalta todos los modelos
  Call ResaltaTodo
  Call chcEtiquetas_Click
  Call chcEtiquetas_Click
  ' Mostrar el Grafo
  'picGrafo.Cls
  'Call MostrarGrafo
End Sub

' DESPLAZAMIENTO EN Z
Private Sub cmdDZ_Click()
' Reinicia Vertices y Aristas
  Call ReiniciaVerticesAristas
  Call CalcularDespalzamientoZ
  Call CalcularTeseracto(miVec())
  ' Mostrar el Grafo
  picGrafo.Cls
  ' Resalta todos los modelos
  Call ResaltaTodo
  Call chcEtiquetas_Click
  Call chcEtiquetas_Click
  ' Mostrar el Grafo
  'picGrafo.Cls
  'Call MostrarGrafo
End Sub

' MOSTRAR Y OCULTAR LAS ETIQUETAS
Private Sub chcEtiquetas_Click()
' Declaracion de variables
  Dim v As Integer
  If chcEtiquetas.Value = 1 Then
    ' Mostrar las etiquetas
    If txtEtiqueta.Count = 1 Then
      For v = 1 To miNumeroVertices
        Load txtEtiqueta(v)
      Next v
    End If
    For v = 1 To miNumeroVertices
      txtEtiqueta(v).ZOrder (0)
      txtEtiqueta(v).Top = 280 - oV(v).PosY / 16
      txtEtiqueta(v).Left = 270 - oV(v).PosX / 16
      txtEtiqueta(v).BackColor = QBColor(oV(v).Color)
      txtEtiqueta(v).Text = oV(v).Numero
      txtEtiqueta(v).Visible = True
    Next v
    DoEvents
  Else
    ' Ocultar las etiquetas
    If txtEtiqueta.Count > 1 Then
      For v = 1 To miNumeroVertices
        txtEtiqueta(v).Visible = False
        Unload txtEtiqueta(v)
      Next v
    End If
    DoEvents
  End If
End Sub

' MOSTRAR U OCULTAR EJES
Private Sub chcEjes_Click()
  If chcEjes.Value = 1 Then
    ' Pinta los ejes de Coordenadas
    picGrafo.Line (0, 4000)-(0, -4000)
    picGrafo.Line (-4000, 0)-(4000, 0)
    picGrafo.Line (-4000, 4000)-(4000, -4000)
    picGrafo.Line (-4000, -4000)-(4000, 4000)
  Else
    picGrafo.Cls
    Call MostrarGrafo
  End If
End Sub

' MOSTRAR U OCULTAR CIRCUNFERENCIA
Private Sub chcCircunferencia_Click()
  If chcCircunferencia = 1 Then
    ' Pinta los ejes de Coordenadas
    picGrafo.Circle (0, 0), miRadio, vbBlack
  Else
    picGrafo.Cls
    Call MostrarGrafo
  End If
End Sub

' AJUSTAR PARAMETROS PARA EL GRAFO
Private Sub AjustarParametros()
' Ajusta segun los parametros
  miFactorCircular = Val(txtFactorCircular.Text)
  miFactorExterno = Val(txtFactorExterno.Text)
  miFactorInterno = Val(txtFactorInterno.Text)
  miFactorProfundidad = Val(txtFactorProfundidad.Text)
  miColorFondo = Val(txtColorFondo.Text)
  miColorAristas = Val(txtColorAristas.Text)
  miColorVertices = Val(txtColorVertices.Text)
  miTamañoVertices = Val(txtTamañoVertices.Text)
  miFormaVertices = Val(txtFormaVertices.Text)
  picGrafo.BackColor = QBColor(miColorFondo)
End Sub

' MOSTRAR EL GRAFO CIRCULAR
Private Sub MostrarCirculo()
' Ubicar los vertices en el orden en el que seran llamados
' Asignar a cada vertice su posicion X,Y
End Sub

' MOSTRAR EL GRAFO TESERACTO
Private Sub MostrarTeseracto()
' Ubicar los vertices en el orden en el que seran llamados
' Asignar a cada vertice su posicion X,Y
End Sub

' PINTAR LAS ARISTAS
Private Sub PintarAristas()
  Dim a As Integer
  Dim d As Integer
  'd = 50
  For a = 1 To miNumeroAristas
    If oA(a).Color <> miColorFondo Then
      If oA(a).Color = 12 Then
        If oV(oA(a).Vertice1).PosX <> oV(oA(a).Vertice2).PosX Then
          d = 30
        Else
          d = 50
        End If
        picGrafo.Line (oV(oA(a).Vertice1).PosX - d, oV(oA(a).Vertice1).PosY - d)-(oV(oA(a).Vertice2).PosX - d, oV(oA(a).Vertice2).PosY - d), QBColor(oA(a).Color)
      Else
        picGrafo.Line (oV(oA(a).Vertice1).PosX, oV(oA(a).Vertice1).PosY)-(oV(oA(a).Vertice2).PosX, oV(oA(a).Vertice2).PosY), QBColor(oA(a).Color)
      End If
    End If
  Next a
End Sub

' PINTAR LOS VERTICES
Private Sub PintarVertices()
  Dim v As Integer
  Dim r As Integer
  For v = 1 To miNumeroVertices
    ' Circular
    If oV(v).Forma = 2 Then
      For r = 1 To oV(v).Tamaño
        If r = oV(v).Tamaño Then
          picGrafo.Circle (oV(v).PosX, oV(v).PosY), r, vbBlack
        Else
          picGrafo.Circle (oV(v).PosX, oV(v).PosY), r, QBColor(oV(v).Color)
        End If

      Next r
    End If
    ' Rectangular
    If oV(v).Forma = 1 Then
      picGrafo.Line (oV(v).PosX - (oV(v).Tamaño), oV(v).PosY - (oV(v).Tamaño))-(oV(v).PosX + (oV(v).Tamaño), oV(v).PosY + (oV(v).Tamaño)), QBColor(oV(v).Color), BF
    End If
  Next v
End Sub

' RESALTA TODOS LOS MODELOS
Private Sub ResaltaTodo()
' Resalta todos los modelos
  Call ResaltarModelo(1, 1, 7, 10, 16)
  Call MostrarGrafo
  Call ResaltarModelo(2, 1, 7, 12, 14)
  Call MostrarGrafo
  Call ResaltarModelo(3, 1, 8, 10, 15)
  Call MostrarGrafo
  Call ResaltarModelo(4, 2, 7, 9, 16)
  Call MostrarGrafo
  Call ResaltarModelo(5, 2, 8, 9, 15)
  Call MostrarGrafo
  Call ResaltarModelo(6, 2, 8, 11, 13)
  Call MostrarGrafo
  Call ResaltarModelo(7, 3, 5, 10, 16)
  Call MostrarGrafo
  Call ResaltarModelo(8, 3, 5, 12, 14)
  Call MostrarGrafo
  Call ResaltarModelo(9, 3, 6, 11, 14)
  Call MostrarGrafo
  Call ResaltarModelo(10, 4, 5, 12, 13)
  Call MostrarGrafo
  Call ResaltarModelo(11, 4, 6, 9, 15)
  Call MostrarGrafo
  Call ResaltarModelo(12, 4, 6, 11, 13)
  Call MostrarGrafo
End Sub

' MOSTRAR EL GRAFO CIRCULAR
Public Sub MostrarGrafo()
' Borrar el grafo
'picGrafo.Cls
' Mostar las aristas
  Call PintarAristas
  ' Mostrar los vertices
  Call PintarVertices
  ' Verifica estado de las check box
  If chcEjes.Value = 1 Then
    Call chcEjes_Click
  End If
  If chcCircunferencia.Value = 1 Then
    Call chcCircunferencia_Click
  End If
End Sub

