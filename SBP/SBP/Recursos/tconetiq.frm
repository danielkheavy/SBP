VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form tconetiq 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion de Etiquetas"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   11340
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Configurar Etiquetas"
      Height          =   4215
      Left            =   1080
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton cmdCancelar 
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
         Left            =   6120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tconetiq.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   7440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tconetiq.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Columnas 
         Height          =   375
         Left            =   960
         MaxLength       =   3
         TabIndex        =   58
         Text            =   "3"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox msuperior 
         Height          =   375
         Left            =   960
         MaxLength       =   5
         TabIndex        =   57
         Text            =   "300"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox sepfila 
         Height          =   375
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   56
         Text            =   "800"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox sepcolumna 
         Height          =   375
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   55
         Text            =   "250"
         Top             =   2880
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   2550
         Left            =   120
         Picture         =   "tconetiq.frx":0F5C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   8685
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Columnas"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   2880
         Width           =   690
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "MargenSup"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   3240
         Width           =   825
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Sep.Fila"
         Height          =   195
         Left            =   2040
         TabIndex        =   62
         Top             =   3240
         Width           =   570
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Sep.Col."
         Height          =   195
         Left            =   2040
         TabIndex        =   61
         Top             =   2880
         Width           =   600
      End
   End
   Begin VB.TextBox presenta 
      Height          =   495
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   53
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox factor 
      Height          =   495
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   52
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox linea 
      Height          =   495
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   51
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox nlinea 
      Height          =   495
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   50
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox unidad 
      Height          =   495
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   49
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox tarjeta 
      Height          =   495
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   48
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox plano 
      Height          =   495
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   47
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox descripcio 
      Height          =   495
      Left            =   1440
      MaxLength       =   53
      TabIndex        =   46
      Top             =   1200
      Width           =   3855
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11280
      TabIndex        =   40
      Top             =   0
      Width           =   11340
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tconetiq.frx":2E4C6
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Consulta"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tconetiq.frx":2F6D8
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tconetiq.frx":308EA
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox cantidad 
      Height          =   495
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   38
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox t16 
      Height          =   495
      Left            =   8040
      MaxLength       =   3
      TabIndex        =   37
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox t15 
      Height          =   495
      Left            =   8040
      MaxLength       =   3
      TabIndex        =   35
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox t14 
      Height          =   495
      Left            =   8040
      MaxLength       =   3
      TabIndex        =   33
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox t13 
      Height          =   495
      Left            =   8040
      MaxLength       =   3
      TabIndex        =   31
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox t12 
      Height          =   495
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   29
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox t11 
      Height          =   495
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   27
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox t10 
      Height          =   495
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   25
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox t9 
      Height          =   495
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   23
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox t8 
      Height          =   495
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   21
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox t7 
      Height          =   495
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   19
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox t6 
      Height          =   495
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   17
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox t5 
      Height          =   495
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   15
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox t4 
      Height          =   495
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox t3 
      Height          =   495
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   11
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox t2 
      Height          =   495
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   9
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox t1 
      Height          =   495
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox producto 
      Height          =   495
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tallas"
      Height          =   375
      Left            =   3480
      TabIndex        =   66
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Presentac."
      Height          =   495
      Left            =   120
      TabIndex        =   65
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label29 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tarjeta"
      Height          =   495
      Left            =   120
      TabIndex        =   45
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plano Nro."
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Linea"
      Height          =   495
      Left            =   120
      TabIndex        =   39
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label xt16 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7440
      TabIndex        =   36
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label xt15 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7440
      TabIndex        =   34
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label xt14 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7440
      TabIndex        =   32
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label xt13 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7440
      TabIndex        =   30
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label xt12 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6120
      TabIndex        =   28
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label xt11 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6120
      TabIndex        =   26
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label xt10 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6120
      TabIndex        =   24
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label xt9 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6120
      TabIndex        =   22
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label xt8 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4800
      TabIndex        =   20
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label xt7 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4800
      TabIndex        =   18
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label xt6 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label xt5 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label xt4 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label xt3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label xt2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label xt1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Factor1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factor"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unidad"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcio"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Menu dki232 
      Caption         =   "&Config"
   End
   Begin VB.Menu kdfi3434 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu leo343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tconetiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim gzbuf(500)  As String

Dim gzbuf1(500) As String

Dim gzbuf2(500) As String

Dim gzbuf3(500) As String

Dim gzbuf4(500) As String

Private Sub cmdCancelar_Click()
    leo343_Click

End Sub

Private Sub cmdGrabar_Click()
    leo343_Click

End Sub

Private Sub cmdPrint_Click()
    kdfi3434_Click

End Sub

Sub ahora_imprime()

    Dim found   As Integer

    Dim FitSize As Integer

    Dim tamano  As Integer

    Dim j       As Integer

    Dim I       As Integer

    Dim X1      As Integer

    Dim max     As Integer

    Dim mx1     As Integer

    Dim a1      As Integer

    Dim k       As Integer

    Dim columna As Integer

    Dim sw      As Integer

    If MsgBox("Desea Imprimir", 1, "Aviso") <> 1 Then Exit Sub
    sw = 0
    Printer.Orientation = 1
    Printer.ScaleMode = 6
    Printer.ScaleWidth = 297
    Printer.ScaleHeight = 210
    Printer.FontSize = 10
    Printer.FontBold = True
    'cargar los 100 primeros registros en el gzbuffer
    
    For I = 1 To Val(t1)

        If Val(t1) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt1
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t2)

        If Val(t2) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt2
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t3)

        If Val(t3) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt3
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t4)

        If Val(t4) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt4
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t5)

        If Val(t5) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt5
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t6)

        If Val(t6) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt6
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I
    
    For I = 1 To Val(t7)

        If Val(t7) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt7
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t8)

        If Val(t8) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt8
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t9)

        If Val(t9) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt9
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t10)

        If Val(t10) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt10
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t11)

        If Val(t11) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt11
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t12)

        If Val(t12) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt12
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t13)

        If Val(t13) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt13
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t14)

        If Val(t14) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt14
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t15)

        If Val(t15) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt15
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I

    For I = 1 To Val(t16)

        If Val(t16) > 0 Then
            k = k + 1
            gzbuf(k) = Mid$("" & descripcio, 1, 30)
            gzbuf1(k) = Mid$("" & descripcio, 31, 20)
            gzbuf2(k) = "Talla:" & xt16
            gzbuf3(k) = "Modelo:" & presenta & " " & unidad
            gzbuf4(k) = "Tar:" & tarjeta & " CodProd." & producto

        End If

    Next I
    
    If k <= 0 Then Exit Sub
    'a1 primera columna
    mx1 = 0
    X1 = 0  'primera fila
    max = k 'nro etiquetas
    columna = Val(columnas)
    j = 1

    If columna > max Then
        columna = max

    End If

    'aqui viene la impresion
    
    Do
    
        If j > max Then Exit Do
        a1 = 5
        tamano = 98
        imprime_columnas X1, a1, mx1, j, tamano, max, columna
    Loop
    
    'colocar "prueba del sistema ", 15, 27
    'colocar "prueba del sistema ", 115, 27
    'colocar "prueba del sistema ", 205, 27
    Printer.EndDoc

End Sub

Sub imprime_columnas(X1 As Integer, _
                     a1 As Integer, _
                     mx1 As Integer, _
                     j As Integer, _
                     tamano As Integer, _
                     max As Integer, _
                     columna)

    Dim I As Integer

    mx1 = X1

    For I = 1 To columna
        X1 = mx1

        If j > max Then Exit For
        dibujar_etiqueta "", a1, X1
        X1 = X1 + 3
        dibujar_etiqueta Mid$(gzbuf(j), 1, 33), a1, X1
        X1 = X1 + 3
        dibujar_etiqueta Mid$(gzbuf1(j), 1, 33), a1, X1
        X1 = X1 + 3
        dibujar_etiqueta Mid$(gzbuf2(j), 1, 33), a1, X1
        X1 = X1 + 3
        dibujar_etiqueta Mid$(gzbuf3(j), 1, 33), a1, X1
        X1 = X1 + 3
        dibujar_etiqueta Mid$(gzbuf4(j), 1, 33), a1, X1
        X1 = X1 + 3
        dibujar_etiqueta "", a1, X1
        'Printer.PaintPicture Picture1.Picture, 0,0,picture1.picture.Width , picture1.picture.Height
        X1 = X1 + 3
        a1 = a1 + tamano
        j = j + 1
    Next I

    If X1 >= 210 Then  'debe haber cambio de pagia
        'MsgBox "Otra Pagina :x1=" & x1 & " a1=" & a1
        Printer.NewPage
        X1 = 0
        mx1 = 0
        a1 = 5

    End If

End Sub

Sub dibujar_etiqueta(buf1 As String, a1 As Integer, X1 As Integer)
    colocar buf1, a1, X1

End Sub

Function colocar(Texto As String, X As Integer, Y As Integer) 'COL,FILA
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print Texto

End Function

Private Sub Command1_Click()

End Sub

Private Sub dki232_Click()
    Frame2.Visible = True

End Sub

Private Sub Form_Activate()

    Dim found As Integer

    found = busca_linea("" & linea)
    found = busca_producto("" & producto)

End Sub

Private Sub kdfi3434_Click()

    Dim found As Integer

    Dim sFile As String

    On Error GoTo cmd89_err

    'If Command1.Visible = True Then Exit Sub
    CommonDialog1.CancelError = True

    On Error Resume Next

    CommonDialog1.ShowPrinter

    If Err.Number = 32755 Then
        Exit Sub

    End If

    On Error GoTo 0

    If CommonDialog1.Orientation = cdlLandscape Then
        Printer.Orientation = cdlLandscape

    End If

    'opcion3 = 0
    'Command1.Visible = True
    'sfile = globaldir & "\temporal\" & gusuario & ".txt"
    'found = imprime_archivoj(sfile, 0)
    'Command1.Visible = False
    'opcion3 = 0
    'cmdPrint_Click
    ahora_imprime
    Exit Sub
cmd89_err:
    MsgBox "Error " & error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub leo343_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False

    End If

    tconetiq.Hide
    Unload tconetiq

End Sub

Function busca_linea(buf As String)

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("linea")
    mytablex.Index = "linea"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_linea = 1
        nlinea = "" & mytablex.Fields("descripcio")
        xt1 = "" & mytablex.Fields("t1")
        xt2 = "" & mytablex.Fields("t2")
        xt3 = "" & mytablex.Fields("t3")
        xt4 = "" & mytablex.Fields("t4")
        xt5 = "" & mytablex.Fields("t5")
        xt6 = "" & mytablex.Fields("t6")
        xt7 = "" & mytablex.Fields("t7")
        xt8 = "" & mytablex.Fields("t8")
        xt9 = "" & mytablex.Fields("t9")
        xt10 = "" & mytablex.Fields("t10")
        xt11 = "" & mytablex.Fields("t11")
        xt12 = "" & mytablex.Fields("t12")
        xt13 = "" & mytablex.Fields("t13")
        xt14 = "" & mytablex.Fields("t14")
        xt15 = "" & mytablex.Fields("t15")
        xt16 = "" & mytablex.Fields("t16")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_producto(buf As String)

    Dim mytablex As Table

    presenta = ""

    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        presenta = "" & mytablex.Fields("presenta")

    End If

    mytablex.Close

End Function

