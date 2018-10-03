VERSION 5.00
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Begin VB.Form tcuadrc1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuadre de Caja"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8625
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox localc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cajeroc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox turnoc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cajac 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check3d4 
      BackColor       =   &H00808080&
      Caption         =   "Solo Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3945
      TabIndex        =   32
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox numcuadre 
      Height          =   375
      Left            =   2400
      MaxLength       =   11
      TabIndex        =   30
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox local1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   27
      Text            =   "%"
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox check3d3 
      BackColor       =   &H00808080&
      Caption         =   "Secciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CheckBox check3d2 
      BackColor       =   &H00808080&
      Caption         =   "Familias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   22
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox check3d1 
      BackColor       =   &H00808080&
      Caption         =   "Incluye.Cod Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   21
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox horai 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox horaf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox turno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   15
      Text            =   "%"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox caja 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   13
      Text            =   "%"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox cajero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "%"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox titulo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   8
      Top             =   4800
      Width           =   3855
   End
   Begin VB.TextBox nrolineas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   7
      Text            =   "45"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8565
      TabIndex        =   2
      Top             =   0
      Width           =   8625
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcuadrc1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcuadrc1.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta"
         Top             =   0
         Width           =   735
      End
   End
   Begin GorditoButton.Boton Command1 
      Height          =   825
      Left            =   6360
      TabIndex        =   35
      Top             =   4800
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1455
      PicturePosition =   0
      Caption         =   "PROCESAR"
      BackColor       =   255
      ResalteColor    =   255
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin GorditoButton.Boton CmdImprimir 
      Height          =   825
      Left            =   6960
      TabIndex        =   36
      Top             =   2880
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1455
      PicturePosition =   0
      Caption         =   "IMPRIMIR"
      BackColor       =   255
      ResalteColor    =   255
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin VB.Label opcioncierres 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "opcioncierres"
      Height          =   195
      Left            =   7080
      TabIndex        =   37
      Top             =   2640
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label nregistro 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   34
      Top             =   3855
      Width           =   1575
   End
   Begin VB.Label flag_proveedor 
      Height          =   255
      Left            =   3360
      TabIndex        =   33
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label pantalla 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   270
      TabIndex        =   31
      Top             =   4530
      Width           =   105
   End
   Begin VB.Label tipoexterno 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Correlativo Cuadre"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   5520
      Width           =   2145
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label flagdiario 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label fecha 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label todos 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3480
      TabIndex        =   24
      Top             =   3960
      Width           =   75
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HoraInicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HoraFinal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cajero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo reporte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Menu dju232 
      Caption         =   "&Buscar"
   End
   Begin VB.Menu flo3 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcuadrc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim opcionimprimir As String

Dim tradiario      As String

Dim tipo_impresion As Integer

Dim sum1           As Double

Dim sum2           As Double

Dim sum3           As Double

Dim sum4           As Double

Dim mytable2       As Table

Dim mytable1       As Table

Dim mytable3       As Table

'Constantes para SendMessage enviar click a otra ventana
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const BM_SETSTATE = &HF3

Private Const WM_LBUTTONDOWN = &H201 ' botón izquierdo abajo

Private Const WM_LBUTTONUP = &H202 ' izquierdo arriba

Private Const WM_LBUTTONDBLCLK As Long = &H203 ' izquierdo doble click

' enviar pulsación de mouse al Hwnd indicado
Sub Enviar_Pulsacion(handle As Long)
    Call SendMessage(handle, BM_SETSTATE, 0, ByVal 0&)
    Call SendMessage(handle, WM_LBUTTONDOWN, 0, ByVal 0&)
    Call SendMessage(handle, WM_LBUTTONUP, 0, ByVal 0&)
    Call SendMessage(handle, BM_SETSTATE, 1, ByVal 0&)

End Sub

Sub borrar_cuadres()

    Dim mytablex As Table

    Dim sw       As String

    On Error GoTo cmd4561_err
   
    sw = "1"
    Set mytablex = mydbxglo.OpenTable(usuariopos & "01")  'cuadre 01
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
    sw = "2"
   
    Set mytablex = mydbxglo.OpenTable(usuariopos & "02")  'cuadre 02
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
   
    sw = "3"
   
    Set mytablex = mydbxglo.OpenTable(usuariopos & "03")      'cuadre 03
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
   
    sw = "4"
   
    Set mytablex = mydbxglo.OpenTable(usuariopos & "04")           'cuadre 04
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
   
    Exit Sub
cmd4561_err:
    MsgBox "Error en borra cuadres " & error & " " & sw, 24, "Aviso"
    mytablex.Close
    Exit Sub

End Sub

Sub borrar_temporal(buf As String)

    Dim mytablex As Table

    On Error GoTo cmd6662_err
   
    Set mytablex = mydbxglo.OpenTable(buf & "D")
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
   
    Set mytablex = mydbxglo.OpenTable(buf & "F")
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
   
    Exit Sub
cmd6662_err:
    MsgBox "Mensaje,Error en Borrar Temporal " & error$, 24, "Aviso"
    mytablex.Close

    Exit Sub

End Sub

Function busca_cierre(buf As String)

    Dim mysnapx As New ADODB.Recordset

    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open "select * from parameca where caja='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mysnapx.RecordCount > 0 Then
        busca_cierre = "" & mysnapx.Fields("cierres")

    End If

    mysnapx.Close

End Function

Function busca_clase(buf As String) As String

    Dim mysnapx As New ADODB.Recordset

    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open "select * from clasesunat where clasesunat='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mysnapx.RecordCount > 0 Then
        busca_clase = Mid$(Trim("" & mysnapx.Fields("descripcio")), 1, 5) & "(" & Trim("" & mysnapx.Fields("percepcion")) & ")"

    End If

    mysnapx.Close

End Function

Function busca_config(sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd6711_err

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If sw = 0 Then
            busca_config = "" & mytablex.Fields("centraliza")

        End If

        If sw = 1 Then
            busca_config = "" & mytablex.Fields("vdolar")

        End If

        If sw = 2 Then
            busca_config = "" & mytablex.Fields("tipo5")

        End If

    End If

    mytablex.Close
    Exit Function
  
cmd6711_err:
    mytablex.Close
    MsgBox "Error en busca_config " + error, 48, "Aviso"
    Exit Function
   
End Function

Function busca_empresa() As String

    On Error GoTo cmd_34emp

    Dim mytablex As New ADODB.Recordset

    'MsgBox empresapos
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from empresa where codigo='" & empresapos & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If Val("" & mytablex.Fields("nro")) = 2 Then
            busca_empresa = "S"

        End If

    End If

    mytablex.Close
    
    Exit Function
cmd_34emp:
    MsgBox "ERROR EN .. EMPRESA .." & error, 24, "AVISO"
    mytablex.Close
    
    Exit Function

End Function

' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
' Se ingreso signo
Function busca_fpago(buf1 As String, sdx As Double, sdx1 As Double)

    Dim buf      As String

    Dim buf3     As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd7811_err

    buf3 = ""

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from fpago where fpago='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf3 = "" & mytablex.Fields("descripcio")
          
        If "" & mytablex.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("total"))
            sdx1 = 0

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            sdx1 = Val("" & mytablex.Fields("total"))
            sdx = 0

        End If

        busca_fpago = 1

    End If

    mytablex.Close
    found = formateaa(buf3, 8, 0, 0)
    Exit Function
cmd7811_err:
    MsgBox "Aviso en busca_fpago " + error$, 48, "Aviso"
    Exit Function

End Function

' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018

Function busca_igv() As Double

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd666_err

    busca_igv = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_igv = Val("" & mytablex.Fields("parivta"))
    Else
        busca_igv = 1

    End If

    mytablex.Close
    Exit Function
cmd666_err:
    MsgBox "Mensaje,Error en moneda " & error$
    mytablex.Close
    Exit Function

End Function

Function busca_inicio(buf2 As String, buf3 As String, buf4 As String) As String

    Dim mysnapx As New ADODB.Recordset

    Dim buf     As String

    '-------------------------

    buf = "select * from " & dbca & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    buf = buf & " and tipo ='" & buf2 & "'"
    buf = buf & " order by fecha,str(numero)"

    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    If mysnapx.RecordCount = 0 Then
        buf3 = ""
        buf4 = ""
    Else
        buf3 = "" & mysnapx.Fields("numero")
        mysnapx.MoveLast
        buf4 = "" & mysnapx.Fields("numero")

    End If

    mysnapx.Close

End Function

Function busca_linea(buf1 As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim buf3     As String

    Dim found    As Integer

    buf3 = ""

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from familia where familia='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        ''''13/09/2017 kenyo Mejor Reporte Ticket
        'buf3 = Mid$("" & mytablex.Fields("descripcio"), 1, 15)
        buf3 = Mid$("" & mytablex.Fields("descripcio"), 1, 25)
        ''''13/09/2017 kenyo Mejor Reporte Ticket
      
        busca_linea = 1

    End If

    mytablex.Close
   
    ''''13/09/2017 kenyo Mejor Reporte Ticket
    'found = formateaa(buf3, 12, 0, 0)
    found = formateaa(buf3, 25, 0, 0)
    ''''13/09/2017 kenyo Mejor Reporte Ticket

End Function

''''14/09/2017 kenyo Mejor Reporte Ticket
Function busca_nombreTicket(buf1 As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim buf3     As String

    Dim found    As Integer

    buf3 = ""

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from tipo where tipo='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf3 = Mid$("" & mytablex.Fields("descripcio"), 1, 16)
 
        busca_nombreTicket = 1

    End If

    mytablex.Close
    found = formateaa("*" & buf3, 16, 0, 0)

End Function

''''14/09/2017 kenyo Mejor Reporte Ticket

Function busca_nombre(buf1 As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim buf3     As String

    Dim found    As Integer

    buf3 = ""

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from tipo where tipo='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf3 = Mid$("" & mytablex.Fields("descripcio"), 1, 6)
        busca_nombre = 1

    End If

    mytablex.Close
    found = formateaa(buf3, 6, 0, 0)

End Function

Function busca_param()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_param = Val("" & mytablex.Fields("imp_und"))

    End If

    mytablex.Close

End Function

Function busca_productoc(buf As String, sw As Integer) As String

    Dim mytablex As New ADODB.Recordset
  
    If sw = 0 Then
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
            busca_productoc = Mid$(Trim("" & mytablex.Fields("descripcio")), 1, 35)
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
      
        Else
            busca_productoc = buf

        End If

        mytablex.Close

    End If

    If sw = 1 Then
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from familia where familia='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
            busca_productoc = Mid$(Trim("" & mytablex.Fields("descripcio")), 1, 35)
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
        Else
            busca_productoc = buf
   
        End If

        mytablex.Close

    End If

    If sw = 2 Then
   
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from ccosto where ccosto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
      
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
            busca_productoc = Mid$(Trim("" & mytablex.Fields("descripcio")), 1, 35)
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
     
            'MsgBox "" & mytablex.Fields("descripcio")
        Else
            busca_productoc = buf

        End If

        mytablex.Close
      
    End If

    Exit Function

End Function
 
'''24/08/2017  Kenyo descripcion larga en reportes ticket
Function busca_productocExtra(buf As String, sw As Integer) As String

    Dim mytablex As New ADODB.Recordset
  
    If sw = 0 Then
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select SUBSTRING(DESCRIPCIO,32,50) as DESCRIPCIO from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
            busca_productocExtra = Mid$(Trim(" " & mytablex.Fields("descripcio")), 1, 35)
            '''24/08/2017  Kenyo descripcion larga en reportes ticket
        Else
            busca_productocExtra = buf

        End If

        mytablex.Close

    End If

End Function

'''24/08/2017  Kenyo descripcion larga en reportes ticket

' Testing Proyecto Facturacion Electronica 01/03/2018
Function busca_tipo2(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select SUBSTRING(DESCRIPCIO,1,14) as descripcio from tipo where tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipo2 = mytablex.Fields("DESCRIPCIO")

        For I = 1 To 14 - Len((mytablex.Fields("DESCRIPCIO")))
            busca_tipo2 = busca_tipo2 & " "
        Next
      
    End If

    mytablex.Close

End Function

Public Function Completa_Espacios(hastaCuanto As Integer, _
                                  myDato As String, _
                                  nuevoDato As String)

    Dim I As Integer

    For I = 1 To hastaCuanto
        myDato = "0" & myDato
    Next
    nuevoDato = myDato

End Function

' Testing Proyecto Facturacion Electronica 01/03/2018

Function busca_tipo(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from tipo where tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Function busca_usuario(xuser As String) As String

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    xnpuerto1 = "1"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from vendedor where codigo='" & xuser & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_usuario = "" '& mytablex.Fields("puerto")
        xnpuerto1 = "" '& mytablex.Fields("tipoca")

    End If

    mytablex.Close
   
End Function

Function busca_puerto_caja(xuser As String) As String

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    xnpuerto1 = "1"
   
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parameca where caja='" & xuser & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If opcionimprimir = "I" Then
            busca_puerto_caja = "" & mytablex.Fields("puertocie")
        Else
            busca_puerto_caja = "" & mytablex.Fields("puertocua")

        End If
      
        xnpuerto1 = "" & mytablex.Fields("tipocie")

    End If

    mytablex.Close
   
End Function

Sub cabecera(bufd As String)

    Dim buf      As String

    Dim titulo   As String

    Dim I        As Integer

    Dim sdx      As Double

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd4433_err

    'MsgBox glocal
    mytablex.Open "select * from TLOCAL where codigo='" & glocal & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = Trim("" & mytablex.Fields("cabecera"))
        found = formateaa(buf, 200, 2, 0)

        'found = formateaa("" & mytablex.Fields("cabecera2"), 30, 2, 0)
        'found = formateaa("", 1, 2, 0)
    End If

    mytablex.Close

    '-----------------------
    'If existearchivo(globalpath & "\001d\01\empresa.txt") = 1 Then
    'Open globalpath & "\001d\01\empresa.txt" For Input As #8
    'buf = ""
    'Do
    ' If EOF(8) Then Exit Do
    '   buf = Input$(1, #8)
    '   found = formateaa(buf, Len(buf), 0, 0)
    'Loop
    'Close #8
    'found = formateaa("", 1, 2, 0)
    'End If
    
    '-----------------------
    'titulo = Mid$(menuipos!nempresa, 1, 15) & "-" & Mid$(menuipos!nlocal, 1, 15)
    'i = (36 - Len(titulo)) / 2
    'found = formateaa(" ", i, 0, 0)
    'found = formateaa(titulo, Len(titulo), 2, 0)
    
    If opcion1 = "5" Then

        '--------- busca correlativo
        If numcuadre.Visible = False Then
            sdx = graba_cierres("" & caja)
            titulo = "CIERRE DEL DIA NRO: " & Format(sdx, "000000")

        End If

        If numcuadre.Visible = True Then
            sdx = graba_cierres("" & caja)
            titulo = "CIERRE DEL DIA NRO: " & Format(Val(numcuadre), "000000")

        End If

        '---------
    End If
    
    If opcion1 <> "5" Then
        titulo = "CUADRE PARCIAL NRO: " & Day(Now) & "-" & Month(Now)

    End If

    buf = titulo
    I = (36 - Len(titulo)) / 2
    found = formateaa(" ", I, 0, 0)
    found = formateaa(titulo, Len(titulo), 2, 0)

    titulo = bufd
    buf = titulo
    I = (36 - Len(titulo)) / 2
    found = formateaa(" ", I, 0, 0)
    found = formateaa(titulo, Len(titulo), 2, 0)
    '-------
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    found = formateaa("CAJERO." & cajero & " CAJA." & caja & " TNO." & turno, 35, 2, 0)
    ver_cajeros
    
    found = formateaa("FECHAI." & fechai & " FECHAF." & fechaf, 35, 2, 0)
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)

    'buf = "usuario    " & menucua!usuario
    'found = formateaa(buf, Len(buf), 2, 0)
    '
    'buf = "Cuadre Dia " & fechai & " - " & FECHAF
    'found = formateaa(buf, Len(buf), 2, 0)

    'buf = String(35, "-")
    'found = formateaa(buf, 35, 2, 0)
    Exit Sub
cmd4433_err:
    MsgBox "Aviso en cabecera " + 48, "Aviso"
    Exit Sub

End Sub

Sub cabeza_divisas()

    Dim buf   As String

    Dim found As Integer

    cabecera "divisas"
    buf = "ES"
    found = formateaa(buf, 2, 0, 0)
    buf = "Numero "
    found = formateaa(buf, 9, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fecha"
    found = formateaa(buf, 8, 0, 0)
    buf = "Hora"
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "M"
    found = formateaa(buf, 1, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "VALOR"
    found = formateaa(buf, 7, 2, 0)
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)

End Sub

Sub cabeza_documento()

    Dim buf   As String

    Dim found As Integer

    ''''14/09/2017 kenyo Mejor Reporte Ticket
    'cabecera "CABEZA DOCUMENTOS"
    'buf = "ES"
    'found = formateaa(buf, 2, 0, 0)
    'buf = "Numero "
    'found = formateaa(buf, 9, 0, 0)
    'found = formateaa("", 1, 0, 0)
    'buf = "Fecha"
    'found = formateaa(buf, 8, 0, 0)
    'buf = "Hora"
    'found = formateaa(buf, 5, 0, 0)
    'found = formateaa("", 1, 0, 0)
    'buf = "M"
    'found = formateaa(buf, 1, 0, 0)
    'found = formateaa("", 1, 0, 0)

    'buf = "VALOR"
    'found = formateaa(buf, 7, 2, 0)
    'buf = String(35, "-")
    'found = formateaa(buf, 35, 2, 0)
       
    cabecera "DOCUMENTOS EMITIDOS TICKETS"
    buf = "E"
    found = formateaa(buf, 2, 0, 0)
    buf = "S"
    found = formateaa(buf, 2, 0, 0)
    buf = "Numero "
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fecha"
    found = formateaa(buf, 6, 0, 0)
       
    buf = "Hora"
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "M"
    found = formateaa(buf, 1, 0, 0)
    found = formateaa("", 1, 0, 0)
       
    buf = "VALOR"
    found = formateaa(buf, 7, 2, 0)
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
       
    ''''14/09/2017 kenyo Mejor Reporte Ticket
End Sub

Private Sub cajac_Change()

    'caja = extra_loquesea(cajac)
    'cajac.SetFocus
End Sub

Private Sub cajac_Click()

    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja..Crear combo cajac en tcuadrc1
    'caja = extra_loquesea(cajac)
    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja..Crear combo cajac en tcuadrc1
End Sub

Private Sub cajac_Validate(Cancel As Boolean)
    caja = extra_loquesea(cajac)
    MsgBox ("AQUI")
    cajac.SetFocus

End Sub

Private Sub cajeroc_Click()
    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja..Crear combo cajac en tcuadrc1
    'cajero = extra_loquesea(cajeroc)
    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja..Crear combo cajac en tcuadrc1

End Sub

Private Sub CmdImprimir_Click()
    opcionimprimir = "I"
    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja
    caja = tptovta.caja
    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja
    MuestraParcial

End Sub

Sub CAJA_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    turno.SetFocus

End Sub

Sub caja_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = &H26 Then
    '   cajero.SetFocus
    '   Exit Sub
    'End If

End Sub

Sub cajero_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    caja.SetFocus

End Sub

Sub cerrar_archivos()
    Exit Sub

End Sub

Sub cerrar_glaboles(I As Integer)

    On Error GoTo cmd3455_err

    Select Case I

        Case 0
         
            Exit Sub

        Case 1
        
            Exit Sub

    End Select

    Exit Sub
cmd3455_err:
    Exit Sub

End Sub

Sub cierre_dia()

    Dim found As Integer

    Dim buf   As String

    On Error GoTo cmd6_err

    '----------------------
    cn.Execute ("DELETE FROM apertura where  caja='" & caja & "'")
    '----------------------
    found = proceso_diario_maestro()

    If found = 1 Then
        borra_cadiario

    End If
   
    MsgBox "Mensaje,Proceso de cierre Terminado", 48, "Aviso"
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error en cierre dia" & error$, 24, "Aviso"
    Exit Sub

End Sub

Private Sub cmdExit_Click()
    flo3_Click

End Sub

Private Sub cmdSort_Click()
    Command1_Click

End Sub

Sub MuestraParcial()

    Dim sw1      As Integer

    Dim sw2      As Integer

    Dim sw3      As Integer

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    If numcuadre.Visible = True Then
        If Not IsNumeric(numcuadre) Then
            numcuadre.SetFocus
            Exit Sub

        End If

    End If

    'MsgBox opcion1
    If opcion1 = "5" Then  'si es cierre el numero de caja debe ser valida
        If parametro_mesas() = "S" Then
            If mesa_abiertas() > 0 Then
                MsgBox "Existen Mesas Abiertas ", 48, "Aviso"
                Exit Sub

            End If

        End If

        If caja = "%" Then
            MsgBox "Debe Digitar al menos una Caja ", 48, "Aviso"
            caja.SetFocus
            Exit Sub

        End If

        turno = "%"
        cajero = "%"

        If CVDate(fechaf) > CVDate(fechai) Then
            fechaf = fechai
            Exit Sub

        End If

        If tcuadrc1.Caption <> "COPIA CIERRE DEL DIA" Then
            found = valida_caja()

            If found = 0 Then
                MsgBox "Caja No Valida", 48, "Aviso"
                fechai.SetFocus
                Exit Sub

            End If

        End If

        'mirar si quiere que centralzen primero
        mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            If "" & mytablex.Fields("centralizacierre") = "S" Then
                tload.Label12.Caption = "CIERRE"
                tload.fechai = fechai
                tload.Show 1
                MsgBox "Presione Enter para Continuar con el Cierre ", 48, "Aviso"

            End If

        End If

        mytablex.Close
   
    End If

    sw1 = 0
    sw2 = 0
    sw3 = 0
    found = creando_cuadres("" & usuariopos)

    If found = 0 Then
        MsgBox "Por favor vuelva ingresar al Programa", 24, "Aviso"
        End
        Exit Sub

    End If

    If Len(fechai) = 0 Then
        fechai = Format(Now, "dd/mm/yyyy")
        Exit Sub

    End If

    If Not IsDate(fechai) Then
        fechai = ""
        fechai.SetFocus
        Exit Sub

    End If

    fechai = Format(fechai, "dd/mm/yyyy")

    If Not IsDate(fechaf) Then
        fechaf = ""
        fechaf.SetFocus
        Exit Sub

    End If

    If Len(fechaf) = 0 Then
        fechaf = Format(Now, "dd/mm/yyyy")
        Exit Sub

    End If

    fechaf = Format(fechaf, "dd/mm/yyyy")

    '--------------
    If opcion1 = "1" Then  'si es cuadre parcial

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from vendedor where codigo='" & cajero & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            sw1 = 1
            xnpuerto1 = "" '& mytablex.Fields("tipoca")

        End If

        mytablex.Close
   
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from parameca where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            sw2 = 1

        End If

        mytablex.Close
   
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from turno where turno='" & turno & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            sw3 = 1

        End If

        mytablex.Close

        '--------------
        'If sw1 = 1 And sw2 = 1 And sw3 = 1 Then
        '---------------- LA IBERICA
        'data1.Connect = "FOXPRO 2.5;"
        'data1.DatabaseName = globaldir
        'data1.RecordSource = "select fpago,descripcio,moneda,TOTAL from fpago where fpago='1' order by val(fpago) "
        'data1.Refresh
        'If data1.Recordset.EOF = True And data1.Recordset.BOF = True Then
        '   data1.Recordset.Close
        '   Exit Sub
        'End If
        'panel3D1.Visible = True
        'table1.SetFocus
        'Exit Sub
        '---------------
        'End If
    End If

    '--------------
    procesar_cuadre 0

End Sub

Sub Command1_Click()
    opcionimprimir = "N"
    MuestraParcial

End Sub

Sub Command1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        sa11_Click
        Exit Sub

    End If

    fechaf_KeyPress (13)

End Sub

Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechaf.SetFocus
        Exit Sub

    End If

End Sub

Sub cuadre_parcial(sw As Integer, sw1 As Integer)
    proceso_impresion sw, sw1

End Sub

Sub cuerpo_programa(sw As Integer)

    Dim buf   As String

    Dim tsw   As Integer

    Dim found As Integer

    Dim I     As Integer

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim sdx3  As Double

    Dim vr    As Integer

    On Error GoTo cmd23_err

    sum1 = 0
    sum2 = 0
    sum3 = 0
    suma5 = 0
    suma6 = 0

    borrar_cuadres
    fecha = "Poniendo Cajeros"

    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    found = formateaa("", 1, 2, 0)
    buf = "          RANGO DE TICKETS"
    found = formateaa(buf, 25, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

    visualiza_cajeros

    'buf = String(35, "-")
    'found = formateaa(buf, 35, 2, 0)
    fecha = "Poniendo " & dicigv
    sdx = busca_igv()

    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    found = formateaa("", 1, 2, 0)
    buf = "     VENTAS POR USUARIO/TURNO"
    found = formateaa(buf, 29, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

    buf = "T/CAMBIO :" & Format(sdx, "0.000")
    found = formateaa(buf, Len(buf), 2, 0)
    fecha = "Acumulando..espere"

    servicio_realizado
    buf = "CAJA/TURNO/CAJERO"
    found = formateaa(buf, Len(buf), 2, 0)
    imprime_cajaturno

    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    'buf = "SERVICIOS"
    'found = formateaa(buf, Len(buf), 2, 0)

    found = formateaa("", 1, 2, 0)
    buf = "        VENTAS POR SERVICIO"
    found = formateaa(buf, 27, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

    imprime_servicio

    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    'buf = "DOCUMENTOS VALORADOS"
    'found = formateaa(buf, Len(buf), 2, 0)

    found = formateaa("", 1, 2, 0)
    buf = "       DOCUMENTOS VALORADOS"
    found = formateaa(buf, 27, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

    sum1 = 0
    sum2 = 0
    sum3 = 0
    sum4 = 0
    imprime_doctos 0

    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    ' buf = "RESUMEN DE VENTAS"
    'found = formateaa(buf, Len(buf), 2, 0)

    found = formateaa("", 1, 2, 0)
    buf = "          RESUMEN DE VENTAS"
    found = formateaa(buf, 27, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

    imprime_valorv

    If todos = "S" Then

        ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
        'buf = "OTROS DOCUMENTOS "
        'found = formateaa(buf, Len(buf), 2, 0)

        found = formateaa("", 1, 2, 0)
        buf = "          OTROS DOCUMENTOS"
        found = formateaa(buf, 26, 2, 0)
        buf = String(35, "=")
        found = formateaa(buf, 35, 2, 0)
        ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

        imprime_doctos 1
        found = formateaa("NETO VENTAS", 14, 0, 0)
        buf = Format(sum1, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(sum2, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

    End If

    'imprime pedidos
    'TOTAL OTROS
    'MsgBox "x"
    If todos = "S" Then

        ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
        'buf = "ORDEN TRABAJO "
        'found = formateaa(buf, Len(buf), 2, 0)

        found = formateaa("", 1, 2, 0)
        buf = "         ORDEN DE TRABAJO"
        found = formateaa(buf, 25, 2, 0)
        buf = String(35, "=")
        found = formateaa(buf, 35, 2, 0)
        ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

        buf = "Nro              Total  Acuenta "
        found = formateaa(buf, Len(buf), 2, 0)
        imprime_orden_trabajo

        ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
        ' buf = "INGRESOS/EGRESOS"
        'found = formateaa(buf, Len(buf), 2, 0)

        found = formateaa("", 1, 2, 0)
        buf = "         INGRESOS/EGRESOS"
        found = formateaa(buf, 25, 2, 0)
        buf = String(35, "=")
        found = formateaa(buf, 35, 2, 0)
        ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

        imprime_recibos

    End If

    'buf = "ORDEN TRABAJO-ABONOS "
    'found = formateaa(buf, Len(buf), 2, 0)
    'imprime_ordenes

    '
    sdx = busca_igv()

    If sdx = 0 Then
        sdx = 1

    End If

    sdx1 = (sum1 + sum3 + suma5) + (sum2 + sum4 + suma6) * sdx
    sdx1 = Format(sdx1, "0.00")
    sdx2 = sdx1 / sdx
    sdx2 = Format(sdx2, "0.00")
    '---------------------------------------------------
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    buf = "TOT.EFE.CAJA "
    found = formateaa(buf, 14, 0, 0)
    found = formateaa("", 1, 0, 0)

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    If busca_config(1) = "N" Then
        sdx2 = 0

    End If

    buf = Format(sdx2, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    '---------------------------------------------------
    fecha = "POR FAVOR ESPERE ...."
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)

    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    found = formateaa("", 1, 2, 0)
    buf = "          FORMAS DE PAGO"
    found = formateaa(buf, 24, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

    forma_pago
    imprime_fpago
    tsw = 8

    If sw = 1 Then
        tsw = 2

    End If

    For I = 1 To tsw
        found = formateaa("", 1, 2, 0)
    Next I

    fecha = "TERMINANDO PROCESO ...."
    Exit Sub
cmd23_err:
    MsgBox "Error en cuerpo programa.." & error$, 48, "Aviso"
    Exit Sub

End Sub

'Sub cuerpo_programa(sw As Integer)
'    Dim buf As String
'    Dim tsw As Integer
'    Dim found As Integer
'    Dim i As Integer
'    Dim sdx As Double
'    Dim sdx1 As Double
'    Dim sdx2 As Double
'    Dim sdx3 As Double
'    Dim vr As Integer
'    On Error GoTo cmd23_err
'    sum1 = 0
'    sum2 = 0
'    sum3 = 0
'    suma5 = 0
'    suma6 = 0
'    borrar_cuadres
'    fecha = "Poniendo Cajeros"
'
'
'      ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
'    found = formateaa("", 1, 2, 0)
'    buf = "          RANGO DE TICKETS"
'    found = formateaa(buf, 25, 2, 0)
'    buf = String(35, "=")
'    found = formateaa(buf, 35, 2, 0)
'    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
'
'
'    visualiza_cajeros
'    'buf = String(35, "-")
'    'found = formateaa(buf, 35, 2, 0)
'    fecha = "Poniendo Igv"
'    sdx = busca_igv()
'
'
'    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
'    found = formateaa("", 1, 2, 0)
'    buf = "     VENTAS POR USUARIO/TURNO"
'    found = formateaa(buf, 29, 2, 0)
'    buf = String(35, "=")
'    found = formateaa(buf, 35, 2, 0)
'    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
'
'    buf = "T/CAMBIO :" & Format(sdx, "0.000")
'    found = formateaa(buf, Len(buf), 2, 0)
'    fecha = "Acumulando..espere"
'    buf = "SERVICIOS"
'    found = formateaa(buf, Len(buf), 2, 0)
'    servicio_realizado
'    imprime_servicio
'    buf = "DOCUMENTOS VALORADOS"
'    found = formateaa(buf, Len(buf), 2, 0)
'    sum1 = 0
'    sum2 = 0
'    sum3 = 0
'    sum4 = 0
'    imprime_doctos 0
'    buf = "RESUMEN DE VENTAS"
'    found = formateaa(buf, Len(buf), 2, 0)
'    imprime_valorv
'    If todos = "S" Then
'       buf = "OTROS DOCUMENTOS "
'       found = formateaa(buf, Len(buf), 2, 0)
'       imprime_doctos 1
'       found = formateaa("NETO VENTAS", 14, 0, 0)
'       buf = Format(sum1, "0.00")
'       found = formateaa(buf, 8, 0, 1)
'       found = formateaa("", 1, 0, 0)
'       buf = Format(sum2, "0.00")
'       found = formateaa(buf, 8, 0, 1)
'       found = formateaa("", 1, 2, 0)
'    End If
'    'imprime pedidos
'    'TOTAL OTROS
'    'MsgBox "x"
'    buf = "ORDEN TRABAJO "
'    found = formateaa(buf, Len(buf), 2, 0)
'    imprime_orden_trabajo
'    buf = "INGRESOS/EGRESOS"
'    found = formateaa(buf, Len(buf), 2, 0)
'    imprime_recibos
'
'    'buf = "ORDEN TRABAJO-ABONOS "
'    'found = formateaa(buf, Len(buf), 2, 0)
'    'imprime_ordenes
'
'    '
'    sdx = busca_igv()
'    If sdx = 0 Then
'       sdx = 1
'    End If
'
'    sdx1 = (sum1 + sum3 + suma5) + (sum2 + sum4 + suma6) * sdx
'    sdx1 = Format(sdx1, "0.00")
'    sdx2 = sdx1 / sdx
'    sdx2 = Format(sdx2, "0.00")
'    '---------------------------------------------------
'    buf = String(35, "-")
'    found = formateaa(buf, 35, 2, 0)
'    buf = "TOT.EFE.CAJA "
'    found = formateaa(buf, 14, 0, 0)
'    found = formateaa("", 1, 0, 0)
'
'    buf = Format(sdx1, "0.00")
'    found = formateaa(buf, 8, 0, 1)
'    found = formateaa("", 1, 0, 0)
'    If busca_config(1) = "N" Then
'       sdx2 = 0
'    End If
'    buf = Format(sdx2, "0.00")
'    found = formateaa(buf, 8, 0, 1)
'    found = formateaa("", 1, 2, 0)
'    '---------------------------------------------------
'    fecha = "POR FAVOR ESPERE ...."
'    buf = String(35, "-")
'    found = formateaa(buf, 35, 2, 0)
'    buf = "FORMA DE PAGO/INGRESOS"
'    found = formateaa(buf, Len(buf), 2, 0)
'    forma_pago
'    imprime_fpago
'    tsw = 8
'    If sw = 1 Then
'       tsw = 2
'    End If
'    For i = 1 To tsw
'    found = formateaa("", 1, 2, 0)
'    Next i
'    fecha = "TERMINANDO PROCESO ...."
'    Exit Sub
'cmd23_err:
'    MsgBox "Error en cuerpo programa.." & error$, 48, "Aviso"
'    Exit Sub
'
'
'
'
'End Sub
'''18/09/2017 KENYO Formato Simple Cierre X & Y
Sub cuerpo_programaSimple(sw As Integer)

    Dim buf   As String

    Dim tsw   As Integer

    Dim found As Integer

    Dim I     As Integer

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim sdx3  As Double

    Dim vr    As Integer

    On Error GoTo cmd23_err

    sum1 = 0
    sum2 = 0
    sum3 = 0
    suma5 = 0
    suma6 = 0
    
    borrar_cuadres
    fecha = "Poniendo Cajeros"
    
    '    found = formateaa("", 1, 2, 0)
    '    buf = "         RESUMEN DE VENTAS"
    '    found = formateaa(buf, 26, 2, 0)
    '
    '    buf = String(35, "=")
    '    found = formateaa(buf, 35, 2, 0)
    '
    '     found = formateaa("", 1, 2, 0)
    '    found = formateaa(".", 1, 2, 0)
    
    found = formateaa("", 1, 2, 0)
    buf = "          RANGO DE TICKETS"
    found = formateaa(buf, 25, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    
    visualiza_cajerosSimple

    fecha = "Poniendo " & dicigv
    sdx = busca_igv()

    fecha = "Acumulando..espere"
    
    servicio_realizado
  
    found = formateaa("", 1, 2, 0)
    buf = "     VENTAS POR USUARIO/TURNO"
    found = formateaa(buf, 29, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
   
    imprime_cajaturno
 
    found = formateaa("", 1, 2, 0)
    buf = "       VENTAS POR DOCUMENTO"
    found = formateaa(buf, 27, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    
    '
    sdx = busca_igv()

    If sdx = 0 Then
        sdx = 1

    End If
    
    sdx1 = (sum1 + sum3 + suma5) + (sum2 + sum4 + suma6) * sdx
    sdx1 = Format(sdx1, "0.00")
    sdx2 = sdx1 / sdx
    sdx2 = Format(sdx2, "0.00")
    '---------------------------------------------------

    '---------------------------------------------------
    fecha = "POR FAVOR ESPERE ...."
    
    '''18/09/2017 KENYO Formato Simple Cierre X & Y
    'found = formateaa(buf, Len(buf), 2, 0)
    sum1 = 0
    sum2 = 0
    sum3 = 0
    sum4 = 0
    imprime_doctos 0
     
    If todos = "S" Then
      
        buf = String(35, "-")
        found = formateaa(buf, 35, 2, 0)
       
        imprime_doctos 1

        '       found = formateaa("NETO VENTAS", 14, 0, 0)
        '       buf = Format(sum1, "0.00")
        '       found = formateaa(buf, 8, 0, 1)
        '       found = formateaa("", 1, 0, 0)
        '       buf = Format(sum2, "0.00")
        '       found = formateaa(buf, 8, 0, 1)
        '       found = formateaa("", 1, 2, 0)
    End If
    
    found = formateaa("", 1, 2, 0)
    buf = "          FORMA DE PAGO"
    found = formateaa(buf, 23, 2, 0)
    buf = String(35, "=")
    found = formateaa(buf, 35, 2, 0)
    
    forma_pagoSimple
    imprime_fpagoSimple
    '''18/09/2017 KENYO Formato Simple Cierre X & Y
    
    tsw = 8

    If sw = 1 Then
        tsw = 2

    End If

    For I = 1 To tsw
        found = formateaa("", 1, 2, 0)
    Next I

    fecha = "TERMINANDO PROCESO ...."
    Exit Sub
cmd23_err:
    MsgBox "Error en cuerpo programa.." & error$, 48, "Aviso"
    Exit Sub
    
End Sub

'''18/09/2017 KENYO Formato Simple Cierre X & Y

Sub imprime_cajaturno()

    Dim buf   As String

    Dim found As Integer

    Dim buf1  As String

    On Error GoTo cmd958_err

    buf = "Ca"
    found = formateaa(buf, 2, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "T"
    found = formateaa(buf, 1, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Cajr"
    found = formateaa(buf, 4, 0, 0)
    found = formateaa("", 1, 0, 0)
     
    buf = "Cant"
    found = formateaa(buf, 4, 0, 0)
    found = formateaa("", 1, 0, 0)
       
    buf = dicmoneda
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "" '"Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)
       
    buf1 = "select * from " & usuariopos & "05"
    Set mysnap = mydbxglo.CreateSnapshot(buf1)

    Do 'Until mysnap.EOF

        If mysnap.EOF Then Exit Do
       
        '03/03/2018 Correccion espacios en blanco en cierres
        If mysnap.Fields("Caja") <> "" Then
            buf = "" & mysnap.Fields("Caja")
            found = formateaa(buf, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mysnap.Fields("turno")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mysnap.Fields("cajero")
            found = formateaa(buf, 4, 0, 0)
            found = formateaa("", 1, 0, 0)
          
            '12/06/2017 kenyo NOTA DE CREDITO
            '12/06/2017 kenyo NOTA DE CREDITO
            buf = "" & mysnap.Fields("cantidad")
            found = formateaa(buf, 4, 0, 1)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 6, 0, 0)
        
            buf = "" & mysnap.Fields("soles")

            If Val(buf) > 0 Then
                buf = Format(Val(buf), "0.00")
            Else
                buf = ""

            End If

            found = formateaa(buf, 7, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mysnap.Fields("dolares")

            If Val(buf) > 0 Then
                buf = Format(Val(buf), "0.00")
            Else
                buf = ""

            End If
        
            found = formateaa(buf, 7, 0, 1)
            found = formateaa("", 1, 2, 0)

        End If

        '03/03/2018 Correccion espacios en blanco en cierres
               
        mysnap.MoveNext
       
    Loop
    mysnap.Close
    
    Exit Sub
cmd958_err:
    MsgBox "Error en imprime cajaturno"
    mysnap.Close
    Exit Sub

End Sub

Sub detalle_divisas()

    Dim found As Integer

    Dim buf   As String

    buf = "" & mysnap.Fields("estado")
    found = formateaa(buf, 1, 0, 0)
    buf = "" & mysnap.Fields("acu")
    found = formateaa(buf, 1, 0, 0)

    buf = Mid$("" & mysnap.Fields("numero"), 1, 10)
    found = formateaa(buf, 11, 0, 0)
    buf = Format("" & mysnap.Fields("fecha"), "dd/mm/yyyy")
    found = formateaa(buf, 10, 0, 0)
    buf = Mid$("" & mysnap.Fields("hora"), 1, 5)
    found = formateaa(buf, 5, 0, 0)
    buf = "" & mysnap.Fields("moneda")
    found = formateaa(buf, 1, 0, 0)
    buf = "" & mysnap.Fields("importe")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 7, 2, 1)

End Sub

Sub detalle_documentos(mysnapx As ADODB.Recordset)

    Dim found As Integer

    Dim buf   As String

    ''''14/09/2017 kenyo Mejor Reporte Ticket
    'buf = "" & mysnapx.Fields("estado")
    'found = formateaa(buf, 1, 0, 0)
    'buf = "" & mysnapx.Fields("servicio")
    'found = formateaa(buf, 1, 0, 0)
       
    'buf = Mid$("" & mysnapx.Fields("numero"), 1, 10)
    'found = formateaa(buf, 11, 0, 0)
    'buf = Format("" & mysnapx.Fields("fecha"), "dd/mm/yyyy")
    'found = formateaa(buf, 10, 0, 0)
       
    'buf = Mid$("" & mysnapx.Fields("hora"), 1, 5)
    'found = formateaa(buf, 5, 0, 0)
       
    'buf = "" & mysnapx.Fields("moneda")
    'found = formateaa(buf, 1, 0, 0)

    'buf = "" & mysnapx.Fields("total")
    'buf = Format(Val(buf), "0.00")
    'found = formateaa(buf, 7, 2, 1)
       
    buf = "" & mysnapx.Fields("estado")
    found = formateaa(buf, 2, 0, 0)
    buf = "" & mysnapx.Fields("servicio")
    found = formateaa(buf, 2, 0, 0)
       
    buf = Mid$("" & mysnapx.Fields("numero"), 1, 8)
    found = formateaa(buf, 9, 0, 0)
       
    buf = Format("" & mysnapx.Fields("fecha"), "dd/mm/yyyy")
    found = formateaa(buf, 5, 0, 0)
       
    found = formateaa(" ", 1, 0, 0)
              
    buf = Mid$("" & mysnapx.Fields("hora"), 1, 5)
    found = formateaa(buf, 5, 0, 0)
       
    found = formateaa(" ", 1, 0, 0)
    buf = "" & mysnapx.Fields("moneda")
    found = formateaa(buf, 1, 0, 0)
       
    found = formateaa(" ", 1, 0, 0)
      
    buf = "" & mysnapx.Fields("total")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 7, 2, 1)
       
    ''''14/09/2017 kenyo Mejor Reporte Ticket

End Sub

Sub detalle_fpagov(mysnapx As ADODB.Recordset)

    Dim found As Integer

    Dim buf   As String

    buf = "" & mysnapx.Fields("estado")
    found = formateaa(buf, 1, 0, 0)

    buf = "" & mysnapx.Fields("tipo")
    found = formateaa(buf, 3, 0, 0)

    buf = "" & mysnapx.Fields("numero")
    found = formateaa(buf, 11, 0, 0)
       
    buf = Format("" & mysnapx.Fields("fecha"), "dd/mm/yyyy")
    found = formateaa(buf, 10, 0, 0)
       
    buf = "" & mysnapx.Fields("moneda")
    found = formateaa(buf, 2, 0, 0)
    buf = "" & mysnapx.Fields("recibe")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 7, 2, 1)

End Sub

''''18/09/2017 kenyo Mejor Reporte Ticket
Sub detalle_proddoc(mysnapx As ADODB.Recordset)

    Dim found As Integer

    Dim buf   As String

    buf = "" & mysnapx.Fields("estado")
    found = formateaa(buf, 2, 0, 0)
    buf = "" & mysnapx.Fields("servicio")
    found = formateaa(buf, 2, 0, 0)

    'buf = "" & mysnapx.Fields("tipo")
    'found = formateaa(buf, 3, 0, 0)
    buf = mysnapx.Fields("serie") & "-" & mysnapx.Fields("numero")
    found = formateaa(buf, 13, 0, 0)
       
    buf = Format("" & mysnapx.Fields("fecha"), "dd/mm")
    found = formateaa(buf, 6, 0, 0)
   
    buf = "" & mysnapx.Fields("moneda")
    found = formateaa(buf, 2, 0, 0)
       
    found = formateaa("", 2, 0, 0)
       
    buf = "" & mysnapx.Fields("total")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 12, 2, 1)

End Sub

Sub detalle_recibos(xsw As String, ksw As Integer)

    Dim buf   As String

    Dim sw    As Integer

    Dim sw1   As Integer

    Dim Tmp   As String

    Dim vr    As Integer

    Dim buf1  As String

    Dim found As Integer

    ReDim secax(30) As String
    ReDim secay(30) As Double
    ReDim secaz(30) As Double

    Dim ind     As Integer

    Dim I       As Integer

    Dim j       As Integer

    Dim am      As String

    Dim am1     As String

    Dim mysnapx As New ADODB.Recordset

    On Error GoTo cmd3244_err

    sum1 = 0
    sum2 = 0
    sum3 = 0
    sum4 = 0

    ''26/09/2017 kenyo Reporte Ingreso Egreso Ventana de ventas

    buf = "select * from " & dbing & " where  "
    ''26/09/2017 kenyo Reporte Ingreso Egreso Ventana de ventas

    'buf = "select sum(total) from " & dbing & " where  "
    buf = buf & "   fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    If ksw = 0 Then
        buf = buf & " and servicio='" & xsw & "'"

    End If

    buf = buf & " and estado='2'"
   
    'buf = buf & " and fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " order by servicio,tipo,str(numero),fecha"

    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    If opcion1 = "20" Then

        '-------------------------------------
        Dim buf2 As String

        cabecera "INGRESO/EGRESO X SECCION"
        buf2 = ""
        buf = "SECC "
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "INGRESO "
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "EGRESO "
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 2, 0)

        '-------------------------------------
        ind = 1

        For I = 1 To 30
            secax(I) = ""
            secay(I) = 0#
            secaz(I) = 0#
        Next I

        Do Until mysnapx.EOF

            '-------------------------------------------
            For j = 1 To 10
                sw1 = 0
       
                '       am = "cseccion" & j
                '       am1 = "seccion" & j
                am = "cseccion" & j
                am1 = "seccion" & j
       
                For I = 1 To ind

                    If secax(I) = "" & mysnapx.Fields(am) Then
                        If "" & mysnapx.Fields("acu") = "X" Then
                            secay(I) = secay(I) + Val("" & mysnapx.Fields(am1))

                        End If

                        If "" & mysnapx.Fields("acu") = "Y" Then
                            secaz(I) = secaz(I) + Val("" & mysnapx.Fields(am1))

                        End If

                        sw1 = 1

                    End If

                Next I

                If sw1 = 0 Then
                    ind = ind + 1
                    secax(ind) = "" & mysnapx.Fields(am)

                    If "" & mysnapx.Fields("acu") = "X" Then
                        secay(ind) = Val("" & mysnapx.Fields(am1))

                    End If

                    If "" & mysnapx.Fields("acu") = "Y" Then
                        secaz(ind) = Val("" & mysnapx.Fields(am1))

                    End If

                End If

            Next j

            '-------------------------------------------
            mysnapx.MoveNext
        Loop
        sum1 = 0
        sum2 = 0

        For I = 1 To ind

            If secay(I) > 0 Or secaz(I) > 0 Then
                buf = "" & secax(I)
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)

                buf = Format(Val("" & secay(I)), "0.00")
                found = formateaa(buf, 9, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(Val("" & secaz(I)), "0.00")
                found = formateaa(buf, 9, 0, 1)
                found = formateaa("", 1, 2, 0)

            End If
       
            ''26/09/2017 kenyo Reporte Ingreso Egreso Ventana de ventas
       
            ''26/09/2017 kenyo Reporte Ingreso Egreso Ventana de ventas
       
            sum1 = sum1 + secay(I)
            sum2 = sum2 + secaz(I)
       
        Next I

        '-- totales
        buf = "TOTAL"
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)

        buf = Format(sum1, "0.00")
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = Format(sum2, "0.00")
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 2, 0)
        mysnapx.Close
        Exit Sub

        '-------------------------------------
    End If

    sw = 0
    Tmp = ""

    Do Until mysnapx.EOF
     
        ''''14/09/2017 kenyo Mejor Reporte Ticket
        buf1 = ""

        If sw = 0 Then
            Tmp = buf1
            buf = Tmp
            found = formateaa(buf, 3, -1, 0)
            found = busca_nombreTicket("" & mysnapx.Fields("tipo"))
            found = formateaa("", 1, 2, 0)
            sw = 1

        End If

        If Tmp <> buf1 Then
            buf = Format(sum1, "0.00")
            found = formateaa("", 27, -1, 0)
            found = formateaa(buf, 9, 2, 0)
            Tmp = buf1
            buf = Tmp
            sum1 = 0
            found = formateaa(buf, 3, -1, 0)
            found = busca_nombreTicket("" & mysnapx.Fields("tipo"))
            found = formateaa("", 1, 2, 0)

        End If

        ''''14/09/2017 kenyo Mejor Reporte Ticket
      
        '---------------------------------------
       
        ''''14/09/2017 kenyo Mejor Reporte Ticket

        buf = "" & mysnapx.Fields("estado")
        found = formateaa(buf, 2, 0, 0)
        buf = "" & mysnapx.Fields("servicio")
        found = formateaa(buf, 2, 0, 0)
       
        buf = Mid$("" & mysnapx.Fields("numero"), 1, 8)
        found = formateaa(buf, 9, 0, 0)
        buf = Format("" & mysnapx.Fields("fecha"), "dd/mm/yyyy")
        found = formateaa(buf, 5, 0, 0)
        found = formateaa(" ", 1, 0, 0)
        buf = Mid$("" & mysnapx.Fields("hora"), 1, 5)
        found = formateaa(buf, 5, 0, 0)
       
        found = formateaa("", 1, 0, 0)
        
        buf = "" & mysnapx.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa(" ", 1, 0, 0)
      
        buf = "" & mysnapx.Fields("total")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 7, 2, 1)
       
        ''''14/09/2017 kenyo Mejor Reporte Ticket
        If Val("" & mysnapx.Fields("estado")) = 2 Then
            found = formateaa(" ", 4, 0, 0)
            found = formateaa("TOTAL", 8, 0, 0)
        
            sum1 = sum1 + Val("" & mysnapx.Fields("total"))

            'sum2 = sum2 + Val("" & mysnapx.Fields("total"))
        End If

        '---------------------------------------
        ''''14/09/2017 kenyo Mejor Reporte Ticket
       
        mysnapx.MoveNext
    Loop
    mysnapx.Close
        
    If sum1 > 0 Then
        
        buf = Format(sum1, "0.00")
        found = formateaa("", 15, 0, 0)
        found = formateaa(buf, 9, 2, 0)

    End If
         
    ''''14/09/2017 kenyo Mejor Reporte Ticket
    '         buf = Format(sum2, "0.00")
    '         found = formateaa("", 27, 0, 0)
    '         found = formateaa(buf, 9, 2, 0)
    ''''14/09/2017 kenyo Mejor Reporte Ticket

    Exit Sub
cmd3244_err:
    MsgBox "Error en detalle recibos " + error, 48, "Aviso"
    Exit Sub

End Sub

''''18/09/2017 kenyo Mejor Reporte Ticket
Sub detalle_unidades()

    Dim found As Integer

    Dim buf   As String

    buf = ""

    If check3d1 = 1 Then

        '       buf = Trim("*" & mysnap.Fields("producto"))
        '       found = formateaa(buf, 14, 0, 0)
        '       found = formateaa("", 1, 2, 0)
    End If

    If check3d2 = 0 Then
        buf = busca_productoc(Trim("" & mysnap.Fields("producto")), 0)

    End If

    If check3d2 = 1 Then
        buf = busca_productoc(Trim("" & mysnap.Fields("producto")), 1)

    End If

    If check3d3 = 1 Then
        buf = busca_productoc(Trim("" & mysnap.Fields("producto")), 2)

    End If
       
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'found = formateaa(Trim(buf), 19, 0, 0)
    found = formateaa("*" & Trim(buf), 32, 0, 0)
    found = formateaa("", 1, 2, 0)
    buf = busca_productocExtra(Trim(" " & mysnap.Fields("producto")), 0)
    found = formateaa(" " & Trim(buf), 19, 0, 0)
       
    '''24/08/2017  Kenyo descripcion larga en reportes ticket 13/09/2017
    
    buf = "" & mysnap.Fields("sentido")
    found = formateaa(buf, 1, 0, 1)
       
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
    'found = formateaa(" ", 8, 2, 1)
    'found = formateaa(" ", 20, 0, 0)
    '''24/08/2017  Kenyo descripcion larga en reportes ticket
           
    buf = "" & mysnap.Fields("cantidad")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 6, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf = "" & mysnap.Fields("totals")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 8, 2, 1)
       
    If Val("" & mysnap.Fields("totald")) > 0 Then
        found = formateaa("", 28, 0, 0)
        buf = "" & mysnap.Fields("totald")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 2, 1)

    End If

    If Val("" & mysnap.Fields("cantidada")) > 0 Then
        found = formateaa("*** ANULADO ", 20, 0, 0)
        buf = "" & mysnap.Fields("cantidada")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("totalsa")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 2, 1)

        If Val("" & mysnap.Fields("totalda")) > 0 Then
            found = formateaa("", 22, 0, 0)
            buf = "" & mysnap.Fields("totalda")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 8, 2, 1)

        End If

    End If

    If Val("" & mysnap.Fields("VALES")) > 0 Then
        found = formateaa("*** VALES ", 20, 0, 0)
        buf = "" & mysnap.Fields("VALES")
        buf = Format(Val(buf), "0")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("TOTALVALES")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 2, 1)

    End If

    If Val("" & mysnap.Fields("exonerado")) > 0 Then
        found = formateaa("*** EXONERADO ", 20, 0, 0)
        buf = "" & mysnap.Fields("exonerado")
        buf = Format(Val(buf), "0")
        found = formateaa(buf, 6, 2, 1)

    End If

End Sub

Sub detalle_vendoc()

    Dim found As Integer

    Dim buf   As String

    If Val("" & mysnap.Fields("estado")) = 2 Then
        sum1 = sum1 + 1
        buf = Format(sum1, "000")
    Else
        buf = ""

    End If

    found = formateaa(buf, 3, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf = "" & mysnap.Fields("estado")
    found = formateaa(buf, 2, 0, 0)

    buf = "" & mysnap.Fields("comanda")
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf = "" & mysnap.Fields("nrocomanda")
    found = formateaa(buf, 5, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf = "" & mysnap.Fields("moneda")
    found = formateaa(buf, 2, 0, 0)

    buf = "" & mysnap.Fields("nrototal")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 7, 0, 1)

    buf = "" & mysnap.Fields("personas")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 7, 2, 1)

End Sub

Sub DKIW2_Click()

    'LPROCE.Show 1
End Sub

Sub documentos_emitidos()

    Dim titulo As String

    Dim I      As Integer

    Dim buf    As String

    Dim found  As Integer

    titulo = " * DOCUMENTOS EMITIDOS * "
    buf = titulo
    I = (36 - Len(titulo)) / 2
    found = formateaa(".", 1, 0, 0)
    found = formateaa(" ", I, 0, 0)
    found = formateaa(titulo, Len(titulo), 2, 0)
    
    buf = "" & Format(Now, "dd/mm/yyyy") & " --- " & Format(Now, "HH:MM:SS")
    I = (36 - Len(buf)) / 2
    found = formateaa(".", 1, 0, 0)
    found = formateaa(" ", I, 0, 0)
    found = formateaa(buf, Len(buf), 2, 0)

    buf = String(35, "-")
    found = formateaa(". ", 1, 0, 0)
    found = formateaa(buf, 35, 2, 0)

    found = formateaa(".", 1, 0, 0)
    buf = "usuario    " & usuariopos
    found = formateaa(buf, Len(buf), 2, 0)

    found = formateaa(".", 1, 0, 0)
    buf = "Cuadre Dia " & fechai & " - " & fechaf
    found = formateaa(buf, Len(buf), 2, 0)

    buf = String(35, "-")
    found = formateaa(". ", 1, 0, 0)
    found = formateaa(buf, 35, 2, 0)

End Sub

Private Sub dju232_Click()
    Command1_Click

End Sub

Sub fechaf_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command1.Enabled = True
    Command1.SetFocus

End Sub

Sub fechaf_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechai.SetFocus
        Exit Sub

    End If

    If KeyCode = &H71 Then  'f2  todos
        If todos = "N" Then
            todos = "S"
            fechaf_KeyPress (13)
            Exit Sub

        End If
    
        If todos = "S" Then
            todos = "N"
            fechaf_KeyPress (13)
            Exit Sub

        End If

    End If
 
End Sub

Sub fechai_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechai) = 0 Then
        fechai = Format(Now, "dd/mm/yyyy")
        Exit Sub

    End If

    fechai = Format(fechai, "dd/mm/yyyy")
    fechaf.SetFocus

End Sub

Sub fechai_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = &H26 Then
    '   horaf.SetFocus
    '   Exit Sub
    'End If

End Sub

Private Sub flo3_Click()
    tcuadrc1.Hide
    Unload tcuadrc1

End Sub

Sub Form_Activate()

    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja
    Dim mytablec As New ADODB.Recordset

    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja

    Dim buf      As String

    If flagdiario <> "1" Then
        verifica_tradiario

    End If

    'check3d1.Value = 1
    'check3d2.Value = 1
    todos = busca_config(2)
    buf = verifica_grupo()

    Select Case buf

        Case "P"
            check3d1.Value = 1
            check3d2.Value = 0

        Case "F"
            check3d2.Value = 1
            check3d1.Value = 0

    End Select

    If vIMPRIMIR = 1 Then Me.Command1_Click

    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja

    localc.Clear
    localc.AddItem "%"
    mytablec.Open "select * from tlocal", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablec.EOF Then Exit Do
        localc.AddItem "" & mytablec.Fields("codigo") & "|" & "" & mytablec.Fields("nombre")
        mytablec.MoveNext
    Loop
    mytablec.Close
    localc.ListIndex = 0

    cajeroc.Clear
    cajeroc.AddItem "%"
    mytablec.Open "select * from vendedor", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablec.EOF Then Exit Do
        cajeroc.AddItem "" & mytablec.Fields("codigo") & "|" & mytablec.Fields("nombre")
        mytablec.MoveNext
    Loop
    mytablec.Close
    cajeroc.ListIndex = 0

    cajac.Clear
    cajac.AddItem "%"
    mytablec.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablec.EOF Then Exit Do
        If "" & mytablec.Fields("terminal") = "C" Then
            cajac.AddItem "" & mytablec.Fields("caja") & "|" & mytablec.Fields("descripcio")

        End If

        mytablec.MoveNext
    Loop
    mytablec.Close
    cajac.ListIndex = 0

    turnoc.Clear
    turnoc.AddItem "%"
    mytablec.Open "select * from turno", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablec.EOF Then Exit Do
        turnoc.AddItem "" & mytablec.Fields("turno") & "|" & mytablec.Fields("descripcio")
        mytablec.MoveNext
    Loop
    mytablec.Close
    turnoc.ListIndex = 0

    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja

End Sub

Sub Form_GotFocus()
    'data1.Connect = "FOXPRO 2.5;"
    'data2.Connect = "FOXPRO 2.5;"
    'data3.Connect = "FOXPRO 2.5;"

End Sub

Sub Form_Load()
    
    tipo_impresion = 0

End Sub

Sub Form_Unload(Cancel As Integer)

    On Error GoTo cmd456_err

    vIMPRIMIR = 0
    cerrar_glaboles 0
    cerrar_glaboles 1
    cerrar_archivo
    'Data1.Recordset.Close
    Exit Sub
cmd456_err:
    Exit Sub

    'Set cuadre40 = Nothing
End Sub

Sub forma_pago()

    Dim vr, buf, buf1, buf2 As String

    Dim buf3     As String

    Dim sdx1     As Double

    Dim sdx      As Double

    Dim asola    As String

    Dim mytablex As Table

    Dim mysnapx  As New ADODB.Recordset

    Dim signos   As Double

    On Error GoTo cmd230_err

    sum1 = 0
    sdx1 = 0
    Set mytablex = mydbxglo.OpenTable(usuariopos & "03")  'cuadre 03
    mytablex.Index = "tipo"
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    buf = "select * from " & dbfp & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    'buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"
    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    'If local1 <> "%" Then
    '   buf = buf & " and local='" & local1 & "'"
    'End If
    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    If todos = "S" Then
        buf = buf & " and ( acu='I' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='W' or acu='V' OR ACU='1') "  'E nota credito
    Else
        buf = buf & " and ( acu='I' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or  ACU='1') "  'E nota credito

    End If

    'buf = buf & " and (  acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='W' or acu='V' OR ACU='1') "  'E nota credito
    buf = buf & " and estado='2'"
    buf = buf & " order by fecha"

    'MsgBox buf
    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    If mysnapx.RecordCount > 0 Then
   
        Do

            If mysnapx.EOF Then Exit Do
            signos = 1
      
            '19/06/2017 kenyo NOTA DE CREDITO
      
            '      'NOTA CREDITO 12/06/2017 KENYO
            '      If "" & mysnapx.Fields("acu") = "E" Then     'nota credito
            '         signos = -1
            '         'MsgBox "abc"
            '      End If
            '      'formas de pago
            '19/06/2017 kenyo NOTA DE CREDITO

            sum1 = sum1 + 1
            fecha = "FORMA DE PAGO ..." & Format(sum1, "00000")
            buf3 = "" & mysnapx.Fields("servicio")

            'buf2 = "" & mysnapx.Fields("acu")
            'verificar si es ingreso egreso venta
            '---------------------------------
            'If buf2 = "V" Then  'EGRESO
            '    buf2 = "E"
            'End If
            'If buf2 <> "E" Then
            '    buf2 = "I"
            'End If
            '--------------------------------
            Select Case buf3

                Case "V"
                    buf2 = "E"

                Case "W"
                    buf2 = "I"

                    'MsgBox "mm"
                Case Else
                    buf2 = "V"

            End Select

            'MsgBox buf2
            mytablex.Seek "=", "" & mysnapx.Fields("fpago"), buf2

            If Not mytablex.NoMatch Then
                mytablex.Edit
                sdx1 = suma_fpago(buf2, mytablex, signos, mysnapx)
                mytablex.Update

                If mysnapx.Fields("moneda") = "D" And sdx1 < 0 Then
                    forma_pago1 buf2, sdx1, mytablex, mysnapx

                End If

            End If

            If mytablex.NoMatch Then
                mytablex.AddNew
                sdx1 = suma_fpago(buf2, mytablex, signos, mysnapx)
                mytablex.Fields("local") = "01"
                mytablex.Update

                If mysnapx.Fields("moneda") = "D" And sdx1 < 0 Then
                    forma_pago1 buf2, sdx1, mytablex, mysnapx

                End If

            End If

            '----
            mysnapx.MoveNext
        Loop

    End If

    mysnapx.Close
    mytablex.Close
    Exit Sub
cmd230_err:
    MsgBox "Error en Forma de Pago1 " & error$, 24, "Aviso"
    mysnapx.Close
    mytablex.Close

    Exit Sub

End Sub

'''18/09/2017 KENYO Formato Simple Cierre X & Y
Sub forma_pagoSimple()

    Dim vr, buf, buf1, buf2 As String

    Dim buf3     As String

    Dim sdx1     As Double

    Dim sdx      As Double

    Dim asola    As String

    Dim mytablex As Table

    Dim mysnapx  As New ADODB.Recordset

    Dim signos   As Double

    On Error GoTo cmd230_err

    sum1 = 0
    sdx1 = 0
    Set mytablex = mydbxglo.OpenTable(usuariopos & "03")  'cuadre 03
    mytablex.Index = "tipo"
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    buf = "select * from " & dbfp & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    'buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"
    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    'If local1 <> "%" Then
    '   buf = buf & " and local='" & local1 & "'"
    'End If
    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    If todos = "S" Then
        buf = buf & " and ( acu='I' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='W' or acu='V' OR ACU='1') "  'E nota credito
    Else
        buf = buf & " and ( acu='I' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or  ACU='1') "  'E nota credito

    End If

    'buf = buf & " and (  acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='W' or acu='V' OR ACU='1') "  'E nota credito
    buf = buf & " and estado='2'"
    buf = buf & " order by fecha"

    'MsgBox buf
    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    If mysnapx.RecordCount > 0 Then
   
        Do

            If mysnapx.EOF Then Exit Do
            signos = 1
      
            sum1 = sum1 + 1
            fecha = "FORMA DE PAGO ..." & Format(sum1, "00000")
            buf3 = "" & mysnapx.Fields("servicio")

            Select Case buf3

                Case "V"
                    buf2 = "E"

                Case "W"
                    buf2 = "I"

                    'MsgBox "mm"
                Case Else
                    buf2 = "V"

            End Select

            'MsgBox buf2
            mytablex.Seek "=", "" & mysnapx.Fields("fpago"), buf2

            If Not mytablex.NoMatch Then
                mytablex.Edit
                sdx1 = suma_fpago(buf2, mytablex, signos, mysnapx)
                mytablex.Update

                If mysnapx.Fields("moneda") = "D" And sdx1 < 0 Then
                    forma_pago1 buf2, sdx1, mytablex, mysnapx

                End If

            End If

            If mytablex.NoMatch Then
                mytablex.AddNew
                sdx1 = suma_fpago(buf2, mytablex, signos, mysnapx)
                mytablex.Fields("local") = "01"
                mytablex.Update

                If mysnapx.Fields("moneda") = "D" And sdx1 < 0 Then
                    forma_pago1 buf2, sdx1, mytablex, mysnapx

                End If

            End If

            '----
            mysnapx.MoveNext
        Loop

    End If

    mysnapx.Close
    mytablex.Close
    Exit Sub
cmd230_err:
    MsgBox "Error en Forma de Pago1 " & error$, 24, "Aviso"
    mysnapx.Close
    mytablex.Close

    Exit Sub

End Sub

'''18/09/2017 KENYO Formato Simple Cierre X & Y

Sub forma_pago1(buf2 As String, _
                sdx1 As Double, _
                mytablex As Table, _
                mysnapx As ADODB.Recordset)

    Dim sdx As Double

    If "" & mysnapx.Fields("fpago") = "2" Then
        mytablex.Seek "=", "1", buf2  'busco soles  1+servicio

        '---------------
        If mytablex.NoMatch Then
            mytablex.AddNew
            mytablex.Fields("local") = "01"
            mytablex.Fields("tipo") = "1"
            mytablex.Fields("servicio") = buf2
            sdx = Val("" & mytablex.Fields("valors")) + sdx1
            mytablex.Fields("valors") = Val(Format(sdx, "0.00"))
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            
            sdx = Val("" & mytablex.Fields("valors")) + sdx1
            mytablex.Fields("valors") = Format(sdx, "0.00")
            mytablex.Update

        End If

        '---------------------
    End If

End Sub

Function graba_cierres(buf As String) As Double

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd_34emp1

    Dim sdx As Double

    ''01/07/2017 Kenyo Correcion cierre correlativo al realizar copia de cierre de caja
    If tcuadrc1.opcioncierres = "N" Then
        Exit Function

    End If

    ''01/07/2017 Kenyo Correcion cierre correlativo al realizar copia de cierre de caja

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parameca where caja='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("cierres")) + 1
        'mytablex.Edit
        mytablex.Fields("cierres") = Format(sdx, "00000")
        mytablex.Update
        graba_cierres = sdx

    End If

    mytablex.Close
    Exit Function
cmd_34emp1:
    MsgBox " Aviso en graba Cierres  " & error$, 48, "AVISO DE NO ERROR"
    Exit Function

End Function

Sub habilita(sw As Integer)

    Dim xsw

    If sw = 0 Then
        xsw = True

    End If

    If sw = 1 Then
        xsw = False

    End If

    cajero.Enabled = xsw
    caja.Enabled = xsw
    turno.Enabled = xsw
    horai.Enabled = xsw
    horaf.Enabled = xsw
    fechai.Enabled = xsw
    fechaf.Enabled = xsw
    Command1.Enabled = xsw

End Sub

Sub horaf_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Val(horaf) = 0 Then
        horaf = "24"
        Exit Sub

    End If

    If Val(horaf) >= 0 And Val(horaf) <= 24 Then
        fechai.SetFocus
        Exit Sub

    End If

End Sub

Sub horaf_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = &H26 Then
    '   horai.SetFocus
    '   Exit Sub
    'End If

End Sub

Sub horai_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If horai = "" Then
        horai = "00"
        Exit Sub

    End If

    If Val(horai) >= 0 And Val(horai) <= 24 Then
        horaf.SetFocus
        Exit Sub

    End If

End Sub

Sub horai_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = &H26 Then
    '   turno.SetFocus
    '   Exit Sub
    'End If

End Sub

Sub Image1_Click()
    sa11_Click

End Sub

Sub imprime_divisas()

    Dim vr, buf1  As String

    Dim buf       As String

    Dim sdx       As Double

    Dim sw        As Integer

    Dim found     As Integer

    Dim soles     As Double

    Dim dolares   As Double

    Dim tsoles    As Double

    Dim tdolares  As Double

    Dim asoles    As Double

    Dim adolares  As Double

    Dim atsoles   As Double

    Dim atdolares As Double

    Dim mysnapx   As New ADODB.Recordset

    'cabeza_divisas
    tsoles = 0
    tdolares = 0
    soles = 0
    dolares = 0
    atsoles = 0
    atdolares = 0
    asoles = 0
    adolares = 0
    buf = "select * from divisa  where  "
    buf = buf & "   fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If
   
    'buf = buf & " and fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    buf = buf & " order by tipo,str(numero),fecha"

    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    Do Until mysnapx.EOF

        If sw = 0 Then
            sw = 1
            buf1 = "" & mysnapx.Fields("tipo")
            buf = "" & mysnapx.Fields("tipo")
            found = formateaa(buf, 6, 0, 0)
            found = busca_nombre(buf)
            found = formateaa("", 1, 2, 0)

        End If

        If buf1 <> "" & mysnapx.Fields("tipo") Then
            subtotal_documentos "Subt", soles, dolares
            buf1 = "" & mysnapx.Fields("tipo")
            buf = "" & mysnapx.Fields("tipo")
            found = formateaa(buf, 6, 0, 0)
            found = busca_nombre(buf)
            found = formateaa("", 1, 2, 0)
            soles = 0
            dolares = 0
            asoles = 0
            adolares = 0

        End If

        detalle_divisas

        If Val("" & mysnapx.Fields("estado")) = 2 Then
            If "" & mysnapx.Fields("moneda") = "S" Then
                soles = soles + Val("" & mysnapx.Fields("importe"))
                tsoles = tsoles + Val("" & mysnapx.Fields("importe"))

            End If

            If "" & mysnapx.Fields("moneda") = "D" Then
                dolares = dolares + Val("" & mysnapx.Fields("importe"))
                tdolares = tdolares + Val("" & mysnapx.Fields("importe"))

            End If

        End If

        If Val("" & mysnapx.Fields("estado")) = 1 Then
            If "" & mysnapx.Fields("moneda") = "S" Then
                asoles = asoles + Val("" & mysnapx.Fields("importe"))
                atsoles = atsoles + Val("" & mysnapx.Fields("importe"))

            End If

            If "" & mysnapx.Fields("moneda") = "D" Then
                adolares = adolares + Val("" & mysnapx.Fields("importe"))
                atdolares = atdolares + Val("" & mysnapx.Fields("importe"))

            End If

        End If

        mysnapx.MoveNext
    Loop

    If soles > 0 Or dolares > 0 Then
        subtotal_documentos "Subt ", soles, dolares

    End If

    subtotal_documentos "Total", tsoles, tdolares
    subtotal_documentos "Anula", atsoles, atdolares
    mysnapx.Close

End Sub

Sub imprime_doctos(sw As Integer)

    Dim soles   As Double

    Dim dolares As Double

    Dim buf     As String

    Dim found   As Integer

    Dim buf2    As String

    Dim xsw     As Integer

    On Error GoTo cmd49_err

    'cabecera "DOCUMENTOS EMITIDOS"
    buf2 = ""
    buf = "Tipo "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)
    soles = 0
    dolares = 0
    Set mysnap = mydbxglo.CreateSnapshot(usuariopos & "02")  'cuadre 02
    xsw = 1
    Do

        If mysnap.EOF Then Exit Do
        buf2 = Mid$("" & mysnap.Fields("tipo"), 2, Len("" & mysnap.Fields("tipo")))
        '    If sw = 2 Then
        '   '12/06/2017 kenyo NOTA DE CREDITO POR E PARA NC ---G
        '     If Mid$("" & mysnap.Fields("tipo"), 1, 2) = "NCF" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
        '     ' If Mid$("" & mysnap.Fields("tipo"), 1, 1) = "N" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
        '         GoTo masvalex
        '         Else: GoTo masvale
        '      End If
        '    End If
        '
        '    If sw = 0 Or sw = 1 Then
        '   '12/06/2017 kenyo NOTA DE CREDITO POR E PARA NC ---G
        '    If Mid$("" & mysnap.Fields("tipo"), 1, 2) = "NCF" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
        '    '  If Mid$("" & mysnap.Fields("tipo"), 1, 1) = "N" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
        '         GoTo masvale
        '      End If
        '    End If
        '
        '     If sw = "N" Then
        '      If Mid$("" & mysnap.Fields("tipo"), 1, 1) = "E" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
        '         GoTo masvale
        '      End If
        '    End If
       
masvalex:

        If sw = 0 Then
            xsw = 0

            If Val(buf2) <> 5 Then
                xsw = 1

            End If

        End If
       
        If sw = 1 Then
            xsw = 0

            If Val(buf2) = 5 Then
                xsw = 1

            End If

        End If
       
        If xsw = 1 Then
            If Len(buf2) > 0 Then
                buf = "" & mysnap.Fields("tipo")
                found = busca_nombre(buf2)
                found = formateaa("", 1, 0, 0)
                buf = "" & mysnap.Fields("nro")
                found = formateaa(buf, 6, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = "" & mysnap.Fields("valors")
                buf = Format(Val(buf), "0.00")
                found = formateaa(buf, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = "" & mysnap.Fields("valord")
                buf = Format(Val(buf), "0.00")
                found = formateaa(buf, 8, 0, 1)
                found = formateaa("", 1, 2, 0)
                soles = soles + Val("" & mysnap.Fields("valors"))
                dolares = dolares + Val("" & mysnap.Fields("valord"))
       
                If Val("" & mysnap.Fields("nroa")) > 0 Then
                    '---------------------------------
                    found = formateaa("ANULAD", 6, 0, 0)
                    'buf = "" & mysnap.Fields("tipo")
                    'found = busca_nombre(buf)
                    found = formateaa("", 1, 0, 0)
                    buf = "" & mysnap.Fields("nroa")
                    found = formateaa(buf, 6, 0, 1)
                    found = formateaa("", 1, 0, 0)
                    buf = "" & mysnap.Fields("valorsa")
                    buf = Format(Val(buf), "0.00")
                    found = formateaa(buf, 8, 0, 1)
                    found = formateaa("", 1, 0, 0)
                    buf = "" & mysnap.Fields("valorda")
                    buf = Format(Val(buf), "0.00")
                    found = formateaa(buf, 8, 0, 1)
                    found = formateaa("", 1, 2, 0)

                End If

                '---------------------------------
            End If

        End If

masvale:
        mysnap.MoveNext
    Loop
    mysnap.Close

    buf = "Total "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)

    If sw = 0 Then
        buf = "Ventas"

    End If

    If sw = 1 Then
        buf = "Otros "

    End If

    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(soles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(dolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd49_err:
    MsgBox "Error en imprime doctos " & error$
    mysnap.Close
    'mydb.Close

    Exit Sub

End Sub

Sub imprime_documentos()

    Dim vr, buf1  As String

    Dim buf       As String

    Dim sdx       As Double

    Dim sw        As Integer

    Dim found     As Integer

    Dim soles     As Double

    Dim dolares   As Double

    Dim tsoles    As Double

    Dim tdolares  As Double

    Dim asoles    As Double

    Dim adolares  As Double

    Dim atsoles   As Double

    Dim atdolares As Double

    Dim mysnapx   As New ADODB.Recordset

    On Error GoTo cmd3411_err

    cabeza_documento
    tsoles = 0
    tdolares = 0
    soles = 0
    dolares = 0

    atsoles = 0
    atdolares = 0
    asoles = 0
    adolares = 0

    buf = "select * from " & dbca & "  where  "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E') "  'E nota credito

    'buf = buf & " and fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    buf = buf & " order by tipo,str(numero),fecha"

    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    ''''14/09/2017 kenyo Mejor Reporte Ticket
    Do Until mysnapx.EOF

        If sw = 0 Then
            sw = 1
            found = formateaa("", 1, 3, 0)
     
            buf1 = "" & mysnapx.Fields("tipo")
            found = formateaa(buf, 16, -1, 0)
            found = busca_nombreTicket(buf1)
            found = formateaa("", 1, 2, 0)

        End If

        If buf1 <> "" & mysnapx.Fields("tipo") Then
            found = formateaa("", 1, 2, 0)
            subtotal_documentos "TOTAL", dolares, soles
            found = formateaa("", 1, 2, 0)
            buf1 = "" & mysnapx.Fields("tipo")
            'buf = "" & mysnapx.Fields("tipo")
            found = formateaa(buf, 16, -1, 0)
            found = busca_nombreTicket(buf1)
            found = formateaa("", 1, 2, 0)
            soles = 0
            dolares = 0
            asoles = 0
            adolares = 0

            ''''14/09/2017 kenyo Mejor Reporte Ticket
        End If

        detalle_documentos mysnapx

        If Val("" & mysnapx.Fields("estado")) = 2 Then
            If "" & mysnapx.Fields("moneda") = "S" Then
                soles = soles + Val("" & mysnapx.Fields("total"))
                tsoles = tsoles + Val("" & mysnapx.Fields("total"))

            End If

            If "" & mysnapx.Fields("moneda") = "D" Then
                dolares = dolares + Val("" & mysnapx.Fields("total"))
                tdolares = tdolares + Val("" & mysnapx.Fields("total"))

            End If

        End If

        If Val("" & mysnapx.Fields("estado")) = 1 Then
            If "" & mysnapx.Fields("moneda") = "S" Then
                asoles = asoles + Val("" & mysnapx.Fields("total"))
                atsoles = atsoles + Val("" & mysnapx.Fields("total"))

            End If

            If "" & mysnapx.Fields("moneda") = "D" Then
                adolares = adolares + Val("" & mysnapx.Fields("total"))
                atdolares = atdolares + Val("" & mysnapx.Fields("total"))

            End If

        End If

        mysnapx.MoveNext
    Loop

    If soles > 0 Or dolares > 0 Then
        found = formateaa("", 1, 2, 0)
        subtotal_documentos "TOTAL ", dolares, soles

    End If

    ''''14/09/2017 kenyo Mejor Reporte Ticket
    found = formateaa("", 1, 2, 0)
    subtotal_documentos "> GRAN TOTAL", tdolares, tsoles
    subtotal_documentos "> ANULADOS", atdolares, atsoles
    found = formateaa("", 1, 2, 0)
    ''''14/09/2017 kenyo Mejor Reporte Ticket
    mysnapx.Close
    Exit Sub
cmd3411_err:
    MsgBox "Error en Imprime Documentos " + error, 48, "Aviso"
    Exit Sub

End Sub

Sub imprime_fpago()

    Dim buf   As String

    Dim found As Integer

    Dim buf1  As String

    Dim sw    As Integer

    Dim vsoles, vdolares As Double

    Dim isoles, idolares As Double

    Dim esoles, edolares As Double

    Dim ssoles, sdolares As Double

    Dim xsoles, xdolares As Double

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim sdx3     As Double

    Dim sdx4     As Double

    Dim psoles   As Double

    Dim pdolares As Double

    On Error GoTo cmd9999_err

    Dim pmoneda As String

    Dim mysnapx As Snapshot
       
    '       buf = "FORMA DE PAGO/INGRESOS"
    '       found = formateaa(buf, Len(buf), 2, 0)
    
    buf = "Fpago "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)

    isoles = 0
    idolares = 0
    
    Set mysnapx = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & " where servicio='I' order by tipo") 'cuadre 03

    Do
    
        If mysnapx.EOF Then Exit Do
        sdx = 0
        fecha = "NO TOQUE EL TECLADO..."
        buf = "" & mysnapx.Fields("tipo")
        found = busca_fpago(buf, psoles, pdolares)
    
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnapx.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnapx.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

        '----------------------------------
        If psoles > 0 Or pdolares > 0 Then
            found = formateaa("*DECLARADO ", 14, 0, 0)
            buf = Format(psoles, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            buf = Format(pdolares, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)
            found = formateaa("*DIFERENCIA ", 14, 0, 0)
            sdx = psoles - Val("" & mysnapx.Fields("valors"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            sdx = pdolares - Val("" & mysnapx.Fields("valord"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)

        End If

        '----------------------------------
        isoles = isoles + Val("" & mysnapx.Fields("valors"))
        idolares = idolares + Val("" & mysnapx.Fields("valord"))
        mysnapx.MoveNext
    Loop
   
    mysnapx.Close

    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    
    '----------- ventas-----------------------
    buf = "FORMA DE PAGO/VENTAS"
    found = formateaa(buf, Len(buf), 2, 0)
    buf = "Fpago "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)

    isoles = 0
    idolares = 0
    
    Set mysnapx = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & " where   servicio='V' order by tipo") 'cuadre 03

    Do
    
        If mysnapx.EOF Then Exit Do
        sdx = 0
       
        fecha = "NO TOQUE EL TECLADO..."
        buf = "" & mysnapx.Fields("tipo")
        found = busca_fpago(buf, psoles, pdolares)
    
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnapx.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnapx.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

        '----------------------------------
        If psoles > 0 Or pdolares > 0 Then
            found = formateaa("*DECLARADO ", 14, 0, 0)
            buf = Format(psoles, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            buf = Format(pdolares, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)
            found = formateaa("*DIFERENCIA ", 14, 0, 0)
            sdx = psoles - Val("" & mysnapx.Fields("valors"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            sdx = pdolares - Val("" & mysnapx.Fields("valord"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)

        End If

        '----------------------------------
        vsoles = vsoles + Val("" & mysnapx.Fields("valors"))
        vdolares = vdolares + Val("" & mysnapx.Fields("valord"))
        mysnapx.MoveNext
    Loop
   
    mysnapx.Close

    '------fin ventas-----------------------------------
    
    buf = "Total "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fpago "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(vsoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(vdolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

    buf = "FORMA DE PAGO/EGRESOS"
    found = formateaa(buf, Len(buf), 2, 0)

    esoles = 0
    edolares = 0

    Set mysnapx = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & " where servicio='E'") 'cuadre 03
    Do

        If mysnapx.EOF Then Exit Do
        buf = "" & mysnapx.Fields("tipo")
        found = busca_fpago(buf, psoles, pdolares)
        'found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        esoles = esoles + Val("" & mysnapx.Fields("valors"))
        edolares = edolares + Val("" & mysnapx.Fields("valord"))
        mysnapx.MoveNext
    Loop
    mysnapx.Close

    buf = "Total "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fpago "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(esoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf = Format(edolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    'saldos
    sw = 0
    ssoles = 0
    sdolares = 0
    xsoles = 0
    xdolares = 0
    
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    ''''06/10/2017 kenyo Mejora Formato Cierre de Caja

    buf = "FORMA DE PAGO/SALDO FINAL"
    found = formateaa(buf, Len(buf), 2, 0)

    Set mysnapx = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & "  order by tipo ") 'cuadre 03

    Do

        If mysnapx.EOF Then Exit Do
        If sw = 0 Then
            buf1 = "" & mysnapx.Fields("tipo")
            buf = "" & mysnapx.Fields("tipo")
            found = busca_fpago(buf, psoles, pdolares)
            'found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 8, 0, 0)
            sw = 1

        End If

        If buf1 <> "" & mysnapx.Fields("tipo") Then
            buf = Format(ssoles, "0.00")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(sdolares, "0.00")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 2, 0)

            buf1 = "" & mysnapx.Fields("tipo")
            buf = "" & mysnapx.Fields("tipo")
            found = busca_fpago(buf, psoles, pdolares)
            'found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 8, 0, 0)
            ssoles = 0
            sdolares = 0

        End If

        If "" & mysnapx.Fields("servicio") <> "E" Then
            ssoles = ssoles + Val("" & mysnapx.Fields("valors"))
            sdolares = sdolares + Val("" & mysnapx.Fields("valord"))
            xsoles = xsoles + Val("" & mysnapx.Fields("valors"))
            xdolares = xdolares + Val("" & mysnapx.Fields("valord"))

        End If

        If "" & mysnapx.Fields("servicio") = "E" Then
            ssoles = ssoles - Val("" & mysnapx.Fields("valors"))
            sdolares = sdolares - Val("" & mysnapx.Fields("valord"))
            xsoles = xsoles - Val("" & mysnapx.Fields("valors"))
            xdolares = xdolares - Val("" & mysnapx.Fields("valord"))

        End If

        mysnapx.MoveNext
    Loop
    mysnapx.Close
    
    'lo puse en el peaje
    'If ssoles > 0 Or sdolares > 0 Then
    buf = Format(ssoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    'End If
    'Exit Sub
    '-----------------------------------------------------
    'OJO ES TEMPORAL JOHNNY SOLO PARA VICUS
    buf = "Subtotal "
    found = formateaa(buf, 14, 0, 0)
    buf = Format(xsoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(xdolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    '-----------------------------------------------------
    sdx = busca_igv()

    If sdx = 0 Then
        sdx = 1

    End If

    sdx1 = xsoles + xdolares * sdx
    sdx1 = Format(sdx1, "0.00")
    sdx2 = sdx1 / sdx
    sdx2 = Format(sdx2, "0.00")

    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    'SOLO TEMPORAL------------------
    buf = "Total "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = " "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    If busca_config(1) = "N" Then
        sdx2 = 0

    End If

    buf = Format(sdx2, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    '---ABACA AQUI---
    Exit Sub
cmd9999_err:
    MsgBox "xx.Error en Imprime Fpago1 " & error$, 24, "Aviso "
    Exit Sub

End Sub

'''18/09/2017 KENYO Formato Simple Cierre X & Y
Sub imprime_fpagoSimple()

    Dim buf   As String

    Dim found As Integer

    Dim buf1  As String

    Dim sw    As Integer

    Dim vsoles, vdolares As Double

    Dim isoles, idolares As Double

    Dim esoles, edolares As Double

    Dim ssoles, sdolares As Double

    Dim xsoles, xdolares As Double

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim sdx3     As Double

    Dim sdx4     As Double

    Dim psoles   As Double

    Dim pdolares As Double

    On Error GoTo cmd9999_err

    Dim pmoneda As String

    Dim mysnapx As Snapshot
       
    buf = "F.PAGO  "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "  Nro.  "
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)

    found = formateaa("", 1, 2, 0)
       
    If todos = "S" Then
        buf = "* INGRESOS"
        found = formateaa(buf, Len(buf), 2, 0)

    End If
    
    isoles = 0
    idolares = 0
    
    Set mysnapx = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & " where  servicio='I' order by tipo") 'cuadre 03
    Do

        If mysnapx.EOF Then Exit Do
        sdx = 0
        fecha = "NO TOQUE EL TECLADO..."
        buf = "" & mysnapx.Fields("tipo")
        found = busca_fpago(buf, psoles, pdolares)
    
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnapx.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnapx.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

        '----------------------------------
        If psoles > 0 Or pdolares > 0 Then
            found = formateaa("*DECLARADO ", 14, 0, 0)
            buf = Format(psoles, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            buf = Format(pdolares, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)
            found = formateaa("*DIFERENCIA ", 14, 0, 0)
            sdx = psoles - Val("" & mysnapx.Fields("valors"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            sdx = pdolares - Val("" & mysnapx.Fields("valord"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)

        End If

        '----------------------------------
        isoles = isoles + Val("" & mysnapx.Fields("valors"))
        idolares = idolares + Val("" & mysnapx.Fields("valord"))
        mysnapx.MoveNext
    Loop
   
    mysnapx.Close

    '----------- ventas-----------------------
    found = formateaa("", 1, 2, 0)
    buf = "* VENTAS"
    found = formateaa(buf, Len(buf), 2, 0)
    '       buf = "Fpago "
    '       found = formateaa(buf, 6, 0, 0)
    '       found = formateaa("", 1, 0, 0)
    '       buf = "Nro   "
    '       found = formateaa(buf, 6, 0, 0)
    '       found = formateaa("", 1, 0, 0)
    '       buf = xxxsoles
    '       found = formateaa(buf, 8, 0, 0)
    '       found = formateaa("", 1, 0, 0)
    '       buf = "Dolares"
    '       found = formateaa(buf, 8, 0, 0)
    '       found = formateaa("", 1, 2, 0)

    isoles = 0
    idolares = 0
    
    Set mysnapx = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & " where  servicio='V' order by tipo") 'cuadre 03

    Do
    
        If mysnapx.EOF Then Exit Do
        sdx = 0
        fecha = "NO TOQUE EL TECLADO..."
        buf = "" & mysnapx.Fields("tipo")
        found = busca_fpago(buf, psoles, pdolares)
    
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnapx.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnapx.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

        '----------------------------------
        If psoles > 0 Or pdolares > 0 Then
            found = formateaa("*DECLARADO ", 14, 0, 0)
            buf = Format(psoles, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            buf = Format(pdolares, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)
            found = formateaa("*DIFERENCIA ", 14, 0, 0)
            sdx = psoles - Val("" & mysnapx.Fields("valors"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            sdx = pdolares - Val("" & mysnapx.Fields("valord"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)

        End If

        '----------------------------------
        vsoles = vsoles + Val("" & mysnapx.Fields("valors"))
        vdolares = vdolares + Val("" & mysnapx.Fields("valord"))
        mysnapx.MoveNext
    Loop
   
    mysnapx.Close

    '------fin ventas-----------------------------------

    If todos = "S" Then
        found = formateaa("", 1, 2, 0)
        buf = "* EGRESOS"
        found = formateaa(buf, Len(buf), 2, 0)

    End If

    esoles = 0
    edolares = 0

    Set mysnapx = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & " where  servicio='E'") 'cuadre 03
    Do

        If mysnapx.EOF Then Exit Do
        buf = "" & mysnapx.Fields("tipo")
        found = busca_fpago(buf, psoles, pdolares)
        'found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "-" & mysnapx.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "-" & mysnapx.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        esoles = esoles + Val("" & mysnapx.Fields("valors"))
        edolares = edolares + Val("" & mysnapx.Fields("valord"))
        mysnapx.MoveNext
    Loop
    mysnapx.Close

    '       buf = "Total "
    '       found = formateaa(buf, 6, 0, 0)
    '       found = formateaa("", 1, 0, 0)
    '       buf = "Fpago "
    '       found = formateaa(buf, 6, 0, 0)
    '       found = formateaa("", 1, 0, 0)
    '       buf = Format(esoles, "0.00")
    '       found = formateaa(buf, 8, 0, 1)
    '       found = formateaa("", 1, 0, 0)
    '
    '       buf = Format(edolares, "0.00")
    '       found = formateaa(buf, 8, 0, 1)
    '       found = formateaa("", 1, 2, 0)
    'saldos
    sw = 0
    ssoles = 0
    sdolares = 0
    xsoles = 0
    xdolares = 0
    found = formateaa("", 1, 2, 0)
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    buf = ">>RESUMEN TOTAL"
    found = formateaa(buf, Len(buf), 2, 0)

    Set mysnapx = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & "  order by tipo ") 'cuadre 03

    Do

        If mysnapx.EOF Then Exit Do
        If sw = 0 Then
            buf1 = "" & mysnapx.Fields("tipo")
            buf = "" & mysnapx.Fields("tipo")
            found = busca_fpago(buf, psoles, pdolares)
            'found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 8, 0, 0)
            sw = 1

        End If

        If buf1 <> "" & mysnapx.Fields("tipo") Then
            buf = Format(ssoles, "0.00")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(sdolares, "0.00")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 2, 0)

            buf1 = "" & mysnapx.Fields("tipo")
            buf = "" & mysnapx.Fields("tipo")
            found = busca_fpago(buf, psoles, pdolares)
            'found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 8, 0, 0)
            ssoles = 0
            sdolares = 0

        End If

        If "" & mysnapx.Fields("servicio") <> "E" Then
            ssoles = ssoles + Val("" & mysnapx.Fields("valors"))
            sdolares = sdolares + Val("" & mysnapx.Fields("valord"))
            xsoles = xsoles + Val("" & mysnapx.Fields("valors"))
            xdolares = xdolares + Val("" & mysnapx.Fields("valord"))

        End If

        If "" & mysnapx.Fields("servicio") = "E" Then
            ssoles = ssoles - Val("" & mysnapx.Fields("valors"))
            sdolares = sdolares - Val("" & mysnapx.Fields("valord"))
            xsoles = xsoles - Val("" & mysnapx.Fields("valors"))
            xdolares = xdolares - Val("" & mysnapx.Fields("valord"))

        End If

        mysnapx.MoveNext
    Loop
    mysnapx.Close
    
    'lo puse en el peaje
    'If ssoles > 0 Or sdolares > 0 Then
    buf = Format(ssoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    'End If
    'Exit Sub
    '-----------------------------------------------------
    'OJO ES TEMPORAL JOHNNY SOLO PARA VICUS
    found = formateaa("", 1, 2, 0)
    buf = "*GRAN TOTAL"
    found = formateaa(buf, 14, 0, 0)
    buf = Format(xsoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(xdolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    '-----------------------------------------------------
    sdx = busca_igv()

    If sdx = 0 Then
        sdx = 1

    End If

    If xdolares = "0" Then
        sdx1 = xsoles + xdolares * sdx
        sdx1 = Format(sdx1, "0.00")
        sdx2 = 0
        sdx2 = Format(sdx2, "0.00")
    Else
   
        sdx1 = xsoles + xdolares * sdx
        sdx1 = Format(sdx1, "0.00")
        sdx2 = sdx1 / sdx
        sdx2 = Format(sdx2, "0.00")
    
    End If
    
    '
    '    buf = String(35, "-")
    '    found = formateaa(buf, 35, 2, 0)
    'SOLO TEMPORAL------------------
    '       buf = "GRAN TOTAL"
    '       found = formateaa(buf, 10, 0, 0)
    '       'found = formateaa("", 1, 0, 0)
    '       buf = ""
    '       found = formateaa(buf, 3, 0, 0)
    '       found = formateaa("", 1, 0, 0)
    '       buf = Format(sdx1, "0.00")
    '       found = formateaa(buf, 8, 0, 1)
    '       found = formateaa("", 1, 0, 0)
    '    If busca_config(1) = "N" Then
    '       sdx2 = 0
    '    End If
    '       buf = Format(sdx2, "0.00")
    '       found = formateaa(buf, 8, 0, 1)
    '       found = formateaa("", 1, 2, 0)
    '---ABACA AQUI---
    Exit Sub
cmd9999_err:
    MsgBox "xx.Error en Imprime Fpago1 " & error$, 24, "Aviso "
    Exit Sub

End Sub

'''18/09/2017 KENYO Formato Simple Cierre X & Y

Sub imprime_fpagodoc()

    Dim vr, buf1  As String

    Dim buf      As String

    Dim sdx      As Double

    Dim sw       As Integer

    Dim found    As Integer

    Dim soles    As Double

    Dim dolares  As Double

    Dim tsoles   As Double

    Dim tdolares As Double

    Dim mysnapx  As New ADODB.Recordset

    'cabeza_documento
    '----------------

    buf = "A"
    found = formateaa(buf, 1, 0, 0)

    buf = "Tip"
    found = formateaa(buf, 3, 0, 0)

    buf = "Numero"
    found = formateaa(buf, 11, 0, 0)
    found = formateaa("", 1, 0, 0)

    buf = "Fecha"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)

    buf = "M"
    found = formateaa(buf, 1, 0, 0)

    buf = "Total"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 2, 0)
    '----------------
    tsoles = 0
    tdolares = 0
    soles = 0
    dolares = 0
    buf = "select * from " & dbfp & " where  "
    buf = buf & "   fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    buf = buf & " and (acu='1' or acu='I' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='W' or acu='V') "  'E nota credito

    'buf = buf & " and fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    buf = buf & " order by fpago,tipo,str(numero)"
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    Do Until mysnapx.EOF

        If sw = 0 Then
            sw = 1
            buf = "" & mysnapx.Fields("fpago")
            found = formateaa(buf, 7, 0, 0)
            buf = "" & mysnapx.Fields("descripcio")
            found = formateaa(buf, 20, 2, 0)
            buf1 = "" & mysnapx.Fields("fpago")

        End If

        If buf1 <> "" & mysnapx.Fields("fpago") Then
            subtotal_proddoc soles, dolares
            buf = "" & mysnapx.Fields("fpago")
            found = formateaa(buf, 7, 0, 0)
            buf = "" & mysnapx.Fields("descripcio")
            found = formateaa(buf, 20, 2, 0)
            buf1 = "" & mysnapx.Fields("fpago")
            soles = 0
            dolares = 0

        End If

        detalle_fpagov mysnapx

        If Val("" & mysnapx.Fields("estado")) = 2 Then
            If "" & mysnapx.Fields("moneda") = "S" Then
                soles = soles + Val("" & mysnapx.Fields("recibe"))
                tsoles = tsoles + Val("" & mysnapx.Fields("recibe"))

            End If

            If "" & mysnapx.Fields("moneda") = "D" Then
                dolares = dolares + Val("" & mysnapx.Fields("recibe"))
                tdolares = tdolares + Val("" & mysnapx.Fields("recibe"))

            End If

        End If

        mysnapx.MoveNext
    Loop

    If soles > 0 Or dolares > 0 Then
        subtotal_proddoc soles, dolares

    End If

    subtotal_proddoc tsoles, tdolares
    mysnapx.Close

End Sub

'''18/09/2017 kenyo Mejor Reporte Ticket
Sub imprime_proddoc()

    Dim vr, buf1  As String

    Dim buf      As String

    Dim sdx      As Double

    Dim sw       As Integer

    Dim found    As Integer

    Dim soles    As Double

    Dim dolares  As Double

    Dim tsoles   As Double

    Dim tdolares As Double

    Dim mysnapx  As New ADODB.Recordset

    'cabeza_documento
    '----------------

    buf = "A"
    found = formateaa(buf, 2, 0, 0)
    buf = "S"
    found = formateaa(buf, 2, 0, 0)

    '       buf = "Tip"
    '       found = formateaa(buf, 3, 0, 0)
    '

    buf = "Numero"
    found = formateaa(buf, 11, 0, 0)
    found = formateaa("", 1, 0, 0)

    buf = " Fecha"
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)

    buf = "M"
    found = formateaa(buf, 3, 0, 0)

    buf = " Total"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 2, 0)

    '----------------
    tsoles = 0
    tdolares = 0
    soles = 0
    dolares = 0
    buf = "select * from " & dbde & " where  "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    'buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E') "  'E nota credito
    buf = buf & " order by producto,tipo,fecha"

    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    Do

        If mysnapx.EOF Then Exit Do
        If sw = 0 Then
            sw = 1
      
            buf1 = "" & mysnapx.Fields("producto")
            buf = "" & mysnapx.Fields("producto")
            'found = formateaa(buf, 7, 0, 0)
            found = formateaa("", 1, 2, 0)
            buf = "* " & mysnapx.Fields("descripcio")
            found = formateaa(buf, 34, 0, 0)
            found = formateaa("", 1, 2, 0)
      
        End If

        If buf1 <> "" & mysnapx.Fields("producto") Then
            subtotal_proddoc dolares, soles
      
            buf1 = "" & mysnapx.Fields("producto")
            buf = "" & mysnapx.Fields("producto")
            found = formateaa("", 1, 2, 0)
            'found = formateaa(buf, 7, 0, 0)
            buf = "* " & mysnapx.Fields("descripcio")
            found = formateaa(buf, 34, 0, 0)
            found = formateaa("", 1, 2, 0)
            soles = 0
            dolares = 0

        End If

        detalle_proddoc mysnapx

        If Val("" & mysnapx.Fields("estado")) = 2 Then
            If "" & mysnapx.Fields("moneda") = "S" Then
                soles = soles + Val("" & mysnapx.Fields("total"))
                tsoles = tsoles + Val("" & mysnapx.Fields("total"))

            End If

            If "" & mysnapx.Fields("moneda") = "D" Then
                dolares = dolares + Val("" & mysnapx.Fields("total"))
                tdolares = tdolares + Val("" & mysnapx.Fields("total"))

            End If

        End If

        mysnapx.MoveNext
    Loop

    If soles > 0 Or dolares > 0 Then
        subtotal_proddoc dolares, soles

    End If

    ''''18/09/2017 kenyo Mejor Reporte Ticket
    total_proddoc tdolares, tsoles
    
    mysnapx.Close

End Sub

Sub imprime_recibos()

    Dim buf     As String

    Dim buf1    As String

    Dim found   As Integer

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim sdx3    As Double

    Dim mysnapx As Snapshot

    On Error GoTo cmd87912_err

    buf = "Servc "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)
    sum3 = 0
    sum4 = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0

    buf1 = "select * from " + usuariopos
    buf1 = buf1 + "01"
    buf1 = buf1 + " where servicio<>null and  mid(servicio,1,1)='W' or mid(servicio,1,1)='V' " 'cuadre 01
    'MsgBox buf1
    Set mysnapx = mydbxglo.CreateSnapshot(buf1)

    Do
      
        If mysnapx.EOF Then Exit Do
        If Mid$("" & mysnapx.Fields("servicio"), 1, 1) = "W" Then
            buf = "Ingreso"
            sum3 = sum3 + Val("" & mysnapx.Fields("valors"))
            sum4 = sum4 + Val("" & mysnapx.Fields("valord"))
            sdx = sdx + Val("" & mysnapx.Fields("valors"))
            sdx1 = sdx1 + Val("" & mysnapx.Fields("valord"))

        End If
       
        If Mid$("" & mysnapx.Fields("servicio"), 1, 1) = "V" Then
            buf = "Egreso"
            sum3 = sum3 - Val("" & mysnapx.Fields("valors"))
            sum4 = sum4 - Val("" & mysnapx.Fields("valord"))
            sdx2 = sdx2 + Val("" & mysnapx.Fields("valors"))
            sdx3 = sdx3 + Val("" & mysnapx.Fields("valord"))

        End If

        buf = busca_tipo(Mid$("" & mysnapx.Fields("servicio"), 2, Len("" & mysnapx.Fields("servicio"))))
        'found = formateaa(Mid$("" & mysnapx.Fields("servicio"), 1, 1) & "*", 2, 0, 0)
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 2, 0, 0)
        buf = "" & mysnapx.Fields("valors")

        If Val(buf) > 0 Then
            buf = Format(Val(buf), "0.00")

        End If

        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("valord")

        If Val(buf) > 0 Then
            buf = Format(Val(buf), "0.00")

        End If

        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        mysnapx.MoveNext
    Loop
    mysnapx.Close
    
    If sdx > 0 Or sdx1 > 0 Then
        buf = "TOT INGRESOS "
        found = formateaa(buf, 15, 0, 0)

        buf = Format(sdx, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = Format(sdx1, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

    End If

    If sdx2 > 0 Or sdx3 > 0 Then
        buf = "TOT EGRESOS "
        found = formateaa(buf, 15, 0, 0)

        buf = Format(sdx2, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = Format(sdx3, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

    End If

    Exit Sub
cmd87912_err:
    MsgBox "11.Mensaje en Imprime_recibos " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub imprime_ordenes()

    Dim buf     As String

    Dim buf1    As String

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim xsoles  As Double

    Dim xdolar  As Double

    Dim found   As Integer

    Dim mysnapx As New ADODB.Recordset

    On Error GoTo cmd1287912_err
   
    buf1 = "select * from cpedidov where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    'buf1 = buf1 & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf1 = buf1 & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"
   
    If local1 <> "%" Then
        buf1 = buf1 & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf1 = buf1 & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf1 = buf1 & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf1 = buf1 & " and turno like '" & turno & "%'"

    End If

    'buf1 = buf1 & " and acu='I' "  'I pedidos
    buf1 = buf1 & " order by fecha ,str(numero)"
   
    mysnapx.Open buf1, cn, adOpenStatic, adLockOptimistic
   
    sdx = 0
    sdx1 = 0
    'suma5 = 0
    'suma6 = 0
    Do

        If mysnapx.EOF Then Exit Do
        buf = "" & mysnapx.Fields("local")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        'buf = "" & mysnapx.Fields("serie")
        'found = formateaa(buf, 3, 0, 0)
        'found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("Tipo")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("Numero")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("total")
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 2, 0)
        mysnapx.MoveNext
    Loop
    mysnapx.Close
   
    Exit Sub
cmd1287912_err:
    MsgBox "11.Mensaje en Imprime_Pedidos " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub imprime_servicio()

    Dim buf   As String

    Dim found As Integer

    Dim buf1  As String

    On Error GoTo cmd58_err

    buf = "Servc "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)
    buf1 = "select * from " & usuariopos & "01" & " where  servicio='Q' OR servicio='S' or servicio='R' OR servicio='A' or servicio='C' or servicio='D'" 'cuadre 01
    Set mysnap = mydbxglo.CreateSnapshot(buf1)

    Do 'Until mysnap.EOF

        If mysnap.EOF Then Exit Do
        If "" & mysnap.Fields("servicio") = "A" Then
            buf = "Rapid"

        End If

        If "" & mysnap.Fields("servicio") = "C" Then
            buf = "SA:" & mysnap.Fields("salon")

        End If

        If "" & mysnap.Fields("servicio") = "D" Then
            buf = "Domic"

        End If

        If "" & mysnap.Fields("servicio") <> "D" And "" & mysnap.Fields("servicio") <> "C" And "" & mysnap.Fields("servicio") <> "A" Then
            buf = servicio_tabla("" & mysnap.Fields("servicio"))

        End If

        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        mysnap.MoveNext
    Loop
    mysnap.Close
    
    Exit Sub
cmd58_err:
    MsgBox "Error en imprime servicio"
    mysnap.Close
    Exit Sub

End Sub

Sub imprime_unidades()

    Dim buf       As String

    Dim buf1      As String

    Dim sw        As Integer

    Dim found     As Integer

    Dim soles     As Double

    Dim dolares   As Double

    Dim cantidad  As Double

    Dim tsoles    As Double

    Dim tdolares  As Double

    Dim tcantidad As Double

    Dim vales     As Double

    Dim tvales    As Double

    Dim bufh      As String

    On Error GoTo cmd89333_err

    tvales = 0
    vales = 0
    soles = 0
    dolares = 0
    cantidad = 0
    tsoles = 0
    tdolares = 0
    tcantidad = 0
    buf = ""
    buf1 = ""
    sw = 0
    bufh = "select * from " & usuariopos & "04" & "  order by grupo,producto"

    If Check3d4.Value = 1 Then
        bufh = "select * from " & usuariopos & "04" & "  order by grupo,producto"

    End If

    Set mysnap = mydbxglo.CreateSnapshot(bufh) 'cuadre 04
    
    Do

        If mysnap.EOF Then Exit Do
        If sw = 0 Then
            If check3d1.Value = 1 Then
                buf1 = Trim("" & mysnap.Fields("grupo"))
           
                ''''13/09/2017 kenyo Mejor Reporte Ticket
                'found = formateaa(buf1, 6, 0, 0)
                'found = formateaa("", 1, 0, 0)
                found = formateaa("", 1, 2, 0)
                found = formateaa(">>>", 5, 0, 0)
                ''''13/09/2017 kenyo Mejor Reporte Ticket
         
                found = busca_linea(buf1)
                found = formateaa("", 1, 2, 0)

            End If

            sw = 1

        End If
       
        If buf1 <> Trim("" & mysnap.Fields("grupo")) Then
            If check3d1.Value = 1 Then
      
                ''''13/09/2017 kenyo Mejor Reporte Ticket
                'subtotal_unidades "Subt", cantidad, soles, dolares, vales
                subtotal_unidades "TOTAL", cantidad, soles, dolares, vales
                ''''13/09/2017 kenyo Mejor Reporte Ticket
         
                buf1 = Trim("" & mysnap.Fields("grupo"))
         
                ''''13/09/2017 kenyo Mejor Reporte Ticket
                'found = formateaa(buf1, 6, 0, 0)
                'found = formateaa("", 1, 0, 0)
                found = formateaa("", 1, 2, 0)
                found = formateaa(">>>", 5, 0, 0)
                ''''13/09/2017 kenyo Mejor Reporte Ticket
     
                found = busca_linea(buf1)
                found = formateaa("", 1, 2, 0)
                buf1 = Trim("" & mysnap.Fields("grupo"))

            End If

            cantidad = 0
            soles = 0
            dolares = 0
            vales = 0

        End If

        detalle_unidades
       
        cantidad = cantidad + Val("" & mysnap.Fields("cantidad"))
        soles = soles + Val("" & mysnap.Fields("totals"))
        dolares = dolares + Val("" & mysnap.Fields("totald"))
        tcantidad = tcantidad + Val("" & mysnap.Fields("cantidad"))
        tsoles = tsoles + Val("" & mysnap.Fields("totals"))
        tdolares = tdolares + Val("" & mysnap.Fields("totald"))
        vales = vales + Val("" & mysnap.Fields("totalvales"))
        tvales = tvales + Val("" & mysnap.Fields("totalvales"))
        mysnap.MoveNext
    Loop
    mysnap.Close

    If cantidad > 0 Then
        If check3d1 = 1 Then
            subtotal_unidades "TOTAL", cantidad, soles, dolares, vales

        End If

    End If
    
    ''''13/09/2017 kenyo Mejor Reporte Ticket
    'subtotal_unidades "<<Total", tcantidad, tsoles, tdolares, tvales
    found = formateaa("", 1, 2, 0)
    subtotal_unidades "GRAN TOTAL", tcantidad, tsoles, tdolares, tvales
    ''''13/09/2017 kenyo Mejor Reporte Ticket
    
    Exit Sub
cmd89333_err:
    MsgBox "Aviso en imprime unidades ,Los campos no son numericos de Familias " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub imprime_valorv()

    Dim sdx      As Double

    Dim asoles   As Double

    Dim adolares As Double

    Dim buf2     As String

    Dim soles    As Double

    Dim dolares  As Double

    Dim solesv   As Double

    Dim dolaresv As Double

    Dim igvs     As Double

    Dim igvd     As Double

    Dim buf      As String

    Dim found    As Integer

    Dim exos     As Double

    Dim exod     As Double

    Dim tax1s    As Double

    Dim tax1d    As Double

    Dim dsctos   As Double

    Dim dsctod   As Double

    Dim brutos   As Double

    Dim brutod   As Double

    Dim tresd    As Double

    Dim tress    As Double

    Dim nodsctos As Double

    Dim nodsctod As Double

    Dim FADX     As Double

    Dim cdetras  As Double

    Dim ndetras  As Double

    Dim cdetrad  As Double

    Dim ndetrad  As Double

    Dim pertots  As Double

    Dim pertotd  As Double

    Dim sertots  As Double

    Dim sertotd  As Double

    Dim redons   As Double

    Dim redond   As Double

    On Error GoTo cmd50_err

    redons = 0
    redond = 0

    sertots = 0
    sertotd = 0

    pertots = 0
    pertotd = 0
    cdetras = 0
    ndetras = 0
    cdetrad = 0
    ndetrad = 0

    nodsctos = 0
    nodsctod = 0
    tresd = 0
    tress = 0
    brutos = 0
    brutod = 0
    dsctos = 0
    dsctod = 0
    asoles = 0
    adolares = 0
    solesv = 0
    dolaresv = 0
    soles = 0
    dolares = 0
    igvs = 0
    tax1s = 0
    tax1d = 0
    igvd = 0
    sum1 = 0
    sum2 = 0

    Set mysnap = mydbxglo.CreateSnapshot(usuariopos & "02")  'cuadre 02

    Do 'Until mysnap.EOF

        If mysnap.EOF Then Exit Do
        buf2 = Mid$("" & mysnap.Fields("tipo"), 1, 1)
    
        '12/06/2017 kenyo NOTA DE CREDITO N POR NOTA CREDITO
        'If Mid$("" & mysnap.Fields("tipo"), 1, 1) = "E" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
        If Mid$("" & mysnap.Fields("tipo"), 1, 1) = "NCF" Then
            '12/06/2017 kenyo NOTA DE CREDITO
         
            GoTo masvale2

        End If

        If Val(buf2) <> 5 Then
            redons = redons + Val("" & mysnap.Fields("redons"))
            redond = redond + Val("" & mysnap.Fields("redond"))
       
            pertots = pertots + Val("" & mysnap.Fields("pertots"))
            pertotd = pertotd + Val("" & mysnap.Fields("pertotd"))
       
            sertots = sertots + Val("" & mysnap.Fields("sertots"))
            sertotd = sertotd + Val("" & mysnap.Fields("sertotd"))
       
            cdetras = cdetras + Val("" & mysnap.Fields("cdetras"))
            ndetras = ndetras + Val("" & mysnap.Fields("ndetras"))
            cdetrad = cdetrad + Val("" & mysnap.Fields("cdetrad"))
            ndetrad = ndetrad + Val("" & mysnap.Fields("ndetrad"))

            solesv = solesv + Val("" & mysnap.Fields("valorvs"))
            dolaresv = dolaresv + Val("" & mysnap.Fields("valorvd"))
       
            '12/06/2017 kenyo NOTA DE CREDITO
            '12/06/2017 kenyo NOTA DE CREDITO
             
            igvs = igvs + Val("" & mysnap.Fields("igvs"))
       
            igvd = igvd + Val("" & mysnap.Fields("igvd"))
            exod = exod + Val("" & mysnap.Fields("exod"))
            exos = exos + Val("" & mysnap.Fields("exos"))
            tax1s = tax1s + Val("" & mysnap.Fields("tax1s"))
            tax1d = tax1d + Val("" & mysnap.Fields("tax1d"))
            soles = soles + Val("" & mysnap.Fields("valors"))
            dolares = dolares + Val("" & mysnap.Fields("valord"))
            dsctos = dsctos + Val("" & mysnap.Fields("dsctos"))
            dsctod = dsctod + Val("" & mysnap.Fields("dsctod"))
            brutos = brutos + Val("" & mysnap.Fields("brutos"))
            brutod = brutod + Val("" & mysnap.Fields("brutod"))
            tress = tress + Val("" & mysnap.Fields("retes"))
            tresd = tresd + Val("" & mysnap.Fields("reted"))
            nodsctos = nodsctos + Val("" & mysnap.Fields("nodsctos"))
            nodsctod = nodsctod + Val("" & mysnap.Fields("nodsctod"))

        Else
            asoles = asoles + Val("" & mysnap.Fields("valors"))
            adolares = adolares + Val("" & mysnap.Fields("valord"))

        End If

masvale2:
        mysnap.MoveNext
    Loop
    mysnap.Close

    buf = "Valor Bruto"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(brutos, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(brutod, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Descuentos "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(dsctos, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(dsctod, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Valor Venta "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    'soles v
    FADX = solesv - exos

    buf = Format(FADX, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(dolaresv, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = dicigv
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
        
    buf = Format(igvs, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(igvd, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
       
    buf = "Percepcion"
    found = formateaa(buf, 13, 2, 0)
    'found = formateaa("", 1, 0, 0)
    'buf = Format(pertots, "0.00")
    'found = formateaa(buf, 8, 0, 1)
    'found = formateaa("", 1, 0, 0)
    'buf = Format(pertotd, "0.00")
    'found = formateaa(buf, 8, 0, 1)
    'found = formateaa("", 1, 2, 0)
    percepcion_quiebre
       
    buf = "Servicio 10%"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(sertots, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sertotd, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
       
    buf = "Imp adicional"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(tax1s, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(tax1d, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
       
    buf = "Detracc.Cobradas"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(cdetras, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(cdetrad, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "DetraccNoCobrabas"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(ndetras, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ndetrad, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Exonerado "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(exos, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(exod, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Otros Dsctos "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(tress, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(tresd, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Imp.Excep. "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(nodsctos, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(nodsctod, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Redondeo "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(redons, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(redond, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Total Ventas"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(soles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(dolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    sum1 = soles + asoles
    sum2 = dolares + adolares

    '-----------------
    If opcion1 = "5" Then  'si es cierre

        'acumulado hasta la fecha
        '--------------se quito
        'sdx = suma_las_ventas()
        'buf = "ACUMUL VTAS. "
        'found = formateaa(buf, 14, 0, 0)
        'buf = Format(sdx, "0.00")
        'found = formateaa(buf, 8, 0, 1)
        'found = formateaa("", 1, 2, 0)
    End If
       
    '-----------------
    Exit Sub
cmd50_err:
    MsgBox "Error en imprime_valorv" & error$, 24, "Aviso"
    mysnap.Close

    Exit Sub

End Sub

Sub imprime_vendoc()

    Dim vr, buf1  As String

    Dim buf      As String

    Dim sdx      As Double

    Dim sw       As Integer

    Dim found    As Integer

    Dim soles    As Double

    Dim dolares  As Double

    Dim tsoles   As Double

    Dim tdolares As Double

    Dim mysnapx  As New ADODB.Recordset

    'cabeza_documento
    '----------------
    buf = "NRO "
    found = formateaa(buf, 4, 0, 0)

    buf = "X"
    found = formateaa(buf, 2, 0, 0)

    buf = "COMANDA"
    found = formateaa(buf, 8, 0, 0)

    buf = "CANTI "
    found = formateaa(buf, 6, 0, 1)
       
    buf = "M"
    found = formateaa(buf, 2, 0, 0)
       
    buf = "TOTAL"
    found = formateaa(buf, 7, 2, 1)

    '----------------
    tsoles = 0
    tdolares = 0
    soles = 0
    dolares = 0
    sum1 = 0
    buf = "select vendedor,estado,comanda,moneda,count(comanda) as nrocomanda,sum(total) as nrototal,count(categoria) as personas from " & dbde & " where "
    buf = buf & "   fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    'buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E') "  'E nota credito

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    buf = buf & " group by vendedor,estado,comanda,moneda "

    buf = buf & " order by vendedor,comanda,moneda "
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    Do

        If mysnapx.EOF Then Exit Do
        If sw = 0 Then
            sw = 1
            buf1 = "" & mysnapx.Fields("vendedor")
            buf = "" & mysnapx.Fields("vendedor")
            found = formateaa(buf, 7, 0, 0)
            buf = "" 'busca nombre
            found = formateaa(buf, 20, 0, 0)
            found = formateaa("", 1, 2, 0)

        End If

        If buf1 <> "" & mysnapx.Fields("vendedor") Then
            subtotal_proddoc soles, dolares
            sum1 = 0
            buf1 = "" & mysnapx.Fields("vendedor")

            buf = "" & mysnapx.Fields("vendedor")
            found = formateaa(buf, 7, 0, 0)
            buf = ""
            found = formateaa(buf, 20, 0, 0)
            found = formateaa("", 1, 2, 0)
            soles = 0
            dolares = 0

        End If

        If Val("" & mysnapx.Fields("estado")) = 2 Then
            detalle_vendoc

        End If

        If Val("" & mysnapx.Fields("estado")) = 2 Then
            If "" & mysnapx.Fields("moneda") = "S" Then
                soles = soles + Val("" & mysnapx.Fields("nrototal"))
                tsoles = tsoles + Val("" & mysnapx.Fields("nrototal"))

            End If

            If "" & mysnapx.Fields("moneda") = "D" Then
                dolares = dolares + Val("" & mysnapx.Fields("nrototal"))
                tdolares = tdolares + Val("" & mysnapx.Fields("nrototal"))

            End If

        End If

        mysnapx.MoveNext
    Loop

    If soles > 0 Or dolares > 0 Then
        subtotal_proddoc soles, dolares

    End If

    subtotal_proddoc tsoles, tdolares
    mysnapx.Close

End Sub

Sub lista_chicas()

End Sub

Sub menu_unidades()

    Dim found As Integer

    Dim buf   As String

    Dim I     As Integer

    On Error GoTo cmd46_err

    'filename = usuariopos
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    
    'If opcion3 = 2 Then
    '   filename = cajero
    'End If
    'borra_archivo
    found = borra_nombre("" & FileName)
    ncanal = 1
    Open FileName For Append As #ncanal
    cabecera "UNIDADES VENDIDAS"
    buf = "PROD"
    found = formateaa(buf, 14, 0, 0)
    buf = "DESC"
    found = formateaa(buf, 6, 0, 0)
    buf = "CANT"
    found = formateaa(buf, 7, 0, 0)
    buf = "TOTAL"
    found = formateaa(buf, 9, 2, 0)
    buf = String(38, "-")
    found = formateaa(buf, 35, 2, 0)
    unidades_vendidas
    imprime_unidades

    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    Close #ncanal
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    'found = ejecuta_shell(0)
    Exit Sub
cmd46_err:
    Close #ncanal
    MsgBox error$
    'MsgBox "Mensaje,Error en Menu unidades"
    Exit Sub

End Sub

Sub procesar_cuadre(sw As Integer)

    Dim found As Integer

    tcuadrc1.Enabled = False
    cajero = UCase(cajero)
    cerrar_archivo
    dbbase = "producto.mdb"
    dbca = "factura"
    dbing = "recibo"
    dbde = "detalle"
    dbfp = "fpagov"
    dbtalla = "tacolor"
    dbserial = "serial"

    If tradiario = "D" Then
        dbbase = "diario.mdb"
        dbca = "cadiario"
        dbing = "recibo"
        dbde = "dediario"
        dbfp = "fpdiario"
        dbtalla = "tacolor"
        dbserial = "sediario"

    End If

    '-----------------------------------
    'verificamos si es cuadre del dia ono
    'If Len(fechaf) = 0 Then
    '   fechaf = Format(Now, "dd/mm/yyyy")
    '   Exit Sub
    'End If
    FileName = usuariopos

    xnpuerto = busca_usuario(FileName)  'aqui se dice que puerto debe imprimir el usuario

    If Len(xnpuerto) = 0 Then
        xnpuerto = "LPT"

    End If

    xnpuerto = busca_puerto_caja("" & caja)

    If Len(xnpuerto) = 0 Then
        xnpuerto = "LPT"

    End If

    'fechaf = Format(fechaf, "dd/mm/yyyy")
    Select Case opcion1

        Case "1":
            Screen.MousePointer = 11
            habilita 1

            'CAMBIO 1 POR 0 PARA QUE IMPRIMA CADA VEZ QUE SE HACE PARCIAL
            'KENYO 12/10/2017
            If opcionimprimir = "I" Then
                cuadre_parcial 1, sw
            Else
                cuadre_parcial 0, sw

            End If
         
            cerrar_archivo
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "2", "20":
            habilita 1
            Screen.MousePointer = 11
            proceso_documentos
            cerrar_archivo
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "3":
            habilita 1
            Screen.MousePointer = 11
            cerrar_archivo
            menu_unidades
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "4":
            habilita 1
            Screen.MousePointer = 11
            proceso_proddoc
            cerrar_archivo
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "5":
            habilita 1
            Screen.MousePointer = 11
            'found = abre_cajon("LPT1", 1)
            cuadre_parcial 1, sw
            cerrar_archivo

            If numcuadre.Visible = False Then

                'MsgBox pantalla
                If pantalla <> "PANTALLA" Then
                    'MsgBox "abc"
                    cierre_dia

                End If

            End If

            '-----------Enviar por correo-----------
         
            envio_correos "11"
         
            Screen.MousePointer = 0
            habilita 0
            End
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "6":
            habilita 1
            Screen.MousePointer = 11
            proceso_fpagodoc
            cerrar_archivo
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "7":
            habilita 1
            Screen.MousePointer = 11
            proceso_vendoc
            cerrar_archivo
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "8"   'divisas
            habilita 1
            Screen.MousePointer = 11
            proceso_divisas
            cerrar_archivo
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "9"   'vendedores productos
            habilita 1
            Screen.MousePointer = 11
            unidades_vendedor
            cerrar_archivo
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

        Case "10"   'vendedores productos
            habilita 1
            Screen.MousePointer = 11
            unidades_vendedorp
            cerrar_archivo
            Screen.MousePointer = 0
            habilita 0
            sa11_Click
            opcion3 = "0"
            Exit Sub

    End Select

End Sub

Function proceso_diario_maestro()

    Dim found As Integer

    'Exit Function
    'If busca_empresa() = "S" Then
    If busca_config(0) = "S" Then
        found = copia_cabecera()
      
    End If

    'End If
    proceso_diario_maestro = found
    Exit Function

End Function

Sub proceso_divisas()

    Dim found As Integer

    Dim I     As Integer

    On Error GoTo cmd1332_err

    'filename = usuariopos
    FileName = globaldir & "\temporal\" & gusuario & ".txt"

    If opcion3 = "2" Then

        'filename = cajero
    End If

    'borra_archivo
    found = borra_nombre("" & FileName)
    ncanal = 1
    Open FileName For Append As #ncanal
    imprime_divisas

    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    Close #ncanal
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    'found = ejecuta_shell(0)
    'editor.Show 1
    Exit Sub
cmd1332_err:
    MsgBox "1.Mensaje,Error en Procesa Documentos " & error$
    Exit Sub

End Sub

Sub proceso_documentos()

    Dim found As Integer

    Dim I     As Integer

    On Error GoTo cmd133_err

    'filename = usuariopos
    FileName = globaldir & "\temporal\" & gusuario & ".txt"

    If opcion3 = "2" Then

        'filename = cajero
    End If

    'borra_archivo
    found = borra_nombre("" & FileName)
    ncanal = 1
    Open FileName For Append As #ncanal

    If opcion1 = "20" Then
        detalle_recibos "", 1
    Else
        imprime_documentos
        detalle_recibos "V", 0
        detalle_recibos "W", 0

    End If

    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    Close #ncanal
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    'found = ejecuta_shell(0)
    'editor.Show 1
    Exit Sub
cmd133_err:
    MsgBox "Mensaje,Error en Procesa Documentos " & error$
    Exit Sub

End Sub

Sub proceso_fpagodoc()

    Dim found As Integer

    Dim I     As Integer

    'filename = usuariopos
    FileName = globaldir & "\temporal\" & gusuario & ".txt"

    If opcion3 = "2" Then

        'filename = cajero
    End If

    found = borra_nombre("" & FileName)
    'borra_archivo
    ncanal = 1
    Open FileName For Append As #ncanal
    cabecera "FORMA DE PAGO VS DOCUMENTOS"
    imprime_fpagodoc

    'printer.FontName = "courier new"
    'printer.FontSize = 8
    'editor.Show 1
    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    Close #ncanal
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    'found = ejecuta_shell(0)

End Sub

'''18/09/2017 KENYO Formato Simple Cierre X & Y
Function verifica_formatocierre() As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select formatocierre from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        verifica_formatocierre = "" & mytablex.Fields("formatocierre")

    End If

    mytablex.Close
    Exit Function

End Function

'Function verifica_formatocierrey() As String
'Dim mytablex As New ADODB.Recordset
'If mytablex.State = 1 Then mytablex.Close
'   mytablex.Open "select formatocierrey from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
'   If mytablex.RecordCount > 0 Then
'        verifica_formatocierrey = "" & mytablex.Fields("formatocierrey")
'    End If
'    mytablex.Close
'    Exit Function
'
'End Function

'''18/09/2017 KENYO Formato Simple Cierre X & Y

Sub proceso_impresion(sw As Integer, sw1 As Integer)

    Dim Puerto As String

    Dim I      As Integer

    On Error GoTo cmd34_err

    Dim buf   As String

    Dim buf2  As String

    Dim buf3  As String

    Dim found As Integer

    contpag = 0
    contlin = 0
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    'cerrar_archivo
    'FileName = usuariopos
    'If opcion3 = 2 Then
    '  filename = "" & cajero
    'End If
    'borra_archivo
    'ncanal = 1
    'Open FileName For Append As #ncanal
    'MsgBox "abc"
    cabecera " HORA " & Format(Now, "hh:mm:ss")
    
    '''18/09/2017 KENYO Formato Simple Cierre X & Y
    'cuerpo_programa sw
   
    buf2 = verifica_formatocierre()

    If buf2 = "S" Then
        cuerpo_programaSimple sw
    Else
        cuerpo_programa sw

    End If
    
    '
    '    If opcion1 = "5" Then 'cierre x
    '     buf3 = verifica_formatocierrey()
    '        If buf3 = "S" Then
    '            cuerpo_programaSimple sw
    '        Else
    '            cuerpo_programa sw
    '        End If
    '    End If
    '
     
    '''18/09/2017 KENYO Formato Simple Cierre X & Y
    
    '-----------------
    'UNIDADES VENDIDAS
    'cabecera "UNIDADES VENDIDAS"
    If sw = 0 Then
        If busca_param() = 1 Then
            'MsgBox "unidades"
         
            fecha = "UNIDADES VENDIDAS PROCESANDO ...."
         
            '13/06/2017 kenyo NOTA DE CREDITO
            '13/06/2017 kenyo NOTA DE CREDITO
            found = formateaa("*** UNIDADES VENDIDAS *** ", 36, 2, 0)
            'buf = "PROD"
            'found = formateaa(buf, 8, 0, 0)
            'buf = "DESC"
            'found = formateaa(buf, 12, 0, 0)
            'buf = "CANT"
            'found = formateaa(buf, 7, 0, 0)
            'buf = "TOTAL"
            'found = formateaa(buf, 9, 2, 0)
            buf = String(38, "-")
            found = formateaa(buf, 35, 2, 0)

            If check3d2.Value = 1 And check3d3.Value = 0 Then
                'MsgBox "abc"
                found = formateaa("*** FAMILIAS *** ", 30, 2, 0)
                unidades_vendidas
                imprime_unidades
                GoTo pasoporaqui

            End If

            If check3d2.Value = 0 And check3d3.Value = 1 Then
                found = formateaa("*** SECCIONES *** ", 30, 2, 0)
                unidades_vendidas
                imprime_unidades
                GoTo pasoporaqui

            End If

            If check3d2.Value = 1 And check3d3.Value = 1 Then
                check3d3.Value = False
                found = formateaa("*** FAMILIAS *** ", 30, 2, 0)
                unidades_vendidas
                imprime_unidades
                check3d2.Value = 0
                check3d3.Value = 1
                found = formateaa("*** SECCIONES *** ", 30, 2, 0)
                unidades_vendidas
                imprime_unidades
                GoTo pasoporaqui

            End If

            'MsgBox "abc"
            If check3d1.Value = 1 Then
                'MsgBox "abc"
                nuevo_unidades
                check3d2.Value = 0
                check3d3.Value = 0

                'found = formateaa("*** PRODUCTOS *** ", 30, 2, 0)
                'unidades_vendidas
                'imprime_unidades
            End If
         
pasoporaqui:

            'imprime_unidades

        End If

    End If

    If sw = 1 Then
          
        If busca_param() = 1 Then
            If check3d2.Value = 1 Or check3d3.Value = 1 Then
                fecha = "UNIDADES VENDIDAS PROCESANDO ...."
                found = formateaa("*** UNIDADES VENDIDAS *** ", 36, 2, 0)

            End If

            'found = formateaa("*** UNIDADES VENDIDAS *** ", 36, 2, 0)
            'buf = "PROD"
            'found = formateaa(buf, 8, 0, 0)
            'buf = "DESC"
            'found = formateaa(buf, 12, 0, 0)
            'buf = "CANT"
            'found = formateaa(buf, 7, 0, 0)
            'buf = "TOTAL"
            'found = formateaa(buf, 9, 2, 0)
            'buf = String(38, "-")
            'found = formateaa(buf, 35, 2, 0)
            'unidades_vendidas
            'imprime_unidades
            '-----------------------
            If check3d2.Value = 1 And check3d3.Value = 0 Then
                found = formateaa("*** FAMILIAS *** ", 30, 2, 0)
                unidades_vendidas
                imprime_unidades

            End If

            If check3d2.Value = 0 And check3d3.Value = 1 Then
                found = formateaa("*** SECCIONES *** ", 30, 2, 0)
                unidades_vendidas
                imprime_unidades

            End If

            If check3d2.Value = 1 And check3d3.Value = 1 Then
                check3d3.Value = 0
                found = formateaa("*** FAMILIAS *** ", 30, 2, 0)
                unidades_vendidas
                imprime_unidades
                check3d2.Value = 0
                check3d3.Value = 1
                found = formateaa("*** SECCIONES *** ", 30, 2, 0)
                unidades_vendidas
                imprime_unidades

            End If

            If check3d1.Value = 1 Then
                nuevo_unidades
                check3d2.Value = 0
                check3d3.Value = 0

                'found = formateaa("*** PRODUCTOS *** ", 30, 2, 0)
                'unidades_vendidas
                'imprime_unidades
            End If

            '-----------------------
        End If

    End If

    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    '-----------------
    Close #ncanal

    If sw = 0 Then
        cerrar_archivo

        If sw1 = 1 Then
            '--------------------------------------------
            Puerto = xnpuerto

            If Len(Puerto) = 0 Then
                Puerto = "LPT1"

            End If

            found = star_sp342(Puerto, ticketera_cajon)
            found = corte_papel(Puerto, Val("" & xnpuerto1))
            contlin = 0
            contpag = 0
            Exit Sub

            '--------------------------------------------
        End If
       
        fecha = "IMPRIMIENDO ...."
        genver.file = globaldir & "\temporal\" & gusuario & ".txt"
        genver.Show 1

        'genver.File = globaldir & "\temporal\" & gusuario & ".txt"
        'genver.Show 1
        'found = ejecuta_shell(0) ojo lo quite aqui no se para que es
        'editor.Show 1
    End If

    If sw = 1 Then
        cerrar_archivo
        Puerto = xnpuerto

        If Len(Puerto) = 0 Then
            Puerto = "LPT"

        End If

        'VERIFICAR EN QUE PUERTO SE VA A IMPRIMIR
        'MsgBox ""
        'End
        If pantalla = "PANTALLA" Then
            genver.file = globaldir & "\temporal\" & gusuario & ".txt"
            genver.Show 1
            contlin = 0
            contpag = 0
            Exit Sub

        End If

        found = verifica_cola()

        If found = 0 Then 'si es impresion directa
            'MsgBox puerto
            'MsgBox "pase found=0"
            'End
            found = star_sp342(Puerto, ticketera_cajon)
            found = corte_papel(Puerto, Val("" & xnpuerto1))

        End If

        'MsgBox "nosequepasa"
        'End
        contlin = 0
        contpag = 0

    End If

    'genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    'found = ejecuta_shell(0)
    Close #1
    cerrar_archivo

    Exit Sub
cmd34_err:
    Close #ncanal
    MsgBox "..Mensaje,..Error en proceso Impresion1 " & error$, 24, "AVISO"
    Exit Sub

End Sub

Sub proceso_proddoc()

    Dim found As Integer

    Dim I     As Integer

    FileName = globaldir & "\temporal\" & gusuario & ".txt"

    'filename = usuariopos
    If opcion3 = "2" Then

        'filename = cajero
    End If

    'borra_archivo
    found = borra_nombre("" & FileName)
    ncanal = 1
    Open FileName For Append As #ncanal
    cabecera "PRODUCTO VS DOCUMENTOS"
    imprime_proddoc

    'printer.FontName = "courier new"
    'printer.FontSize = 8
    'editor.Show 1
    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    Close #ncanal
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    'found = ejecuta_shell(0)

    'editor.Show 1
    'found = ejecuta_shell()

End Sub

Function verifica_cola()

    Dim xcolax   As String

    Dim xxpuerto As String

    Dim mytablex As New ADODB.Recordset

    Dim oldprinter

    Dim found As Integer

    Dim sFile As String

    On Error GoTo cmd812_err

    xxpuerto = ""
    xcolax = ""
    
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parameca where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xcolax = "" & mytablex.Fields("colacie")

        If opcionimprimir = "I" Then
            xxpuerto = "" & mytablex.Fields("puertocua")
        Else
            xxpuerto = "" & mytablex.Fields("puertocie")

        End If
      
    End If

    mytablex.Close

    If xcolax = "S" Then
        oldprinter = Printer.DeviceName
        selecciona_impresoras (xxpuerto)
        sFile = globaldir & "\temporal\" & gusuario & ".txt"
        found = Imprime_archivojj(sFile, 0, "9", "", "S", "")
        selecciona_impresoras (oldprinter)
        verifica_cola = 1

    End If

    Exit Function
cmd812_err:
    MsgBox "Aviso en verifica cola ", 48, "Aviso"
    Exit Function

End Function

Sub proceso_vendoc()

    Dim found As Integer

    Dim I     As Integer

    'filename = usuariopos
    FileName = globaldir & "\temporal\" & gusuario & ".txt"

    If opcion3 = "2" Then

        'filename = cajero
    End If

    'borra_archivo
    found = borra_nombre("" & FileName)
    ncanal = 1
    Open FileName For Append As #ncanal
    cabecera "VENDEDOR VS COMANDA"
    imprime_vendoc

    'printer.FontName = "courier new"
    'printer.FontSize = 8
    'editor.Show 1
    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    Close #ncanal
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    'found = ejecuta_shell(0)

    'editor.Show 1
    'found = ejecuta_shell()

End Sub

Sub sa11_Click()
    tcuadrc1.Hide
    Unload tcuadrc1

    'Set cuadre40 = Nothing
End Sub

Sub suma_cajas(mytablez As Table, signos As Double, mysnapx As ADODB.Recordset)
   
    If Val("" & mysnapx.Fields("tipo")) = 5 Then
        If todos <> "S" Then Exit Sub

    End If
      
    mytablez.Fields("caja") = "" & mysnapx.Fields("caja")
    mytablez.Fields("turno") = "" & mysnapx.Fields("turno")
    mytablez.Fields("cajero") = "" & mysnapx.Fields("usuario")

    '12/06/2017 kenyo NOTA DE CREDITO
    'mytablez.Fields("cantidad") = Val("" & mytablez.Fields("cantidad")) + SIGNOS
    mytablez.Fields("cantidad") = Val("" & mytablez.Fields("cantidad")) + 1
    '12/06/2017 kenyo NOTA DE CREDITO
   
    If "" & mysnapx.Fields("moneda") = "S" Then
        mytablez.Fields("soles") = Val("" & mytablez.Fields("soles")) + (signos) * Val("" & mysnapx.Fields("total"))

    End If

    If "" & mysnapx.Fields("moneda") = "D" Then
        mytablez.Fields("dolares") = Val("" & mytablez.Fields("dolares")) + (signos) * Val("" & mysnapx.Fields("total"))

    End If
   
End Sub

Sub servicio_realizado()

    Dim found As Integer

    Dim vr, buf, buf1 As String

    Dim buf2 As String

    On Error GoTo cmd56_err

    Dim sdx      As Double

    Dim mytablex As Table

    Dim mytabley As Table

    Dim mytablez As Table

    Dim mysnapx  As New ADODB.Recordset

    Dim signos   As Double

    sum1 = 0
    Set mytablex = mydbxglo.OpenTable(usuariopos & "01")   'cuadre 01
    mytablex.Index = "salon"
    Set mytabley = mydbxglo.OpenTable(usuariopos & "02")  'cuadre 02
    mytabley.Index = "tipo"
    Set mytablez = mydbxglo.OpenTable(usuariopos & "05")  'cuadre 02
    mytablez.Index = "caja"
   
    buf2 = "select * from " & dbca & " where "
    buf2 = buf2 & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf2 = buf2 & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf2 = buf2 & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf2 = buf2 & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf2 = buf2 & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf2 = buf2 & " and turno like '" & turno & "%'"

    End If
   
    buf2 = buf2 & " and ( acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='1') "  'E nota credito
   
    buf2 = buf2 & " and (servicio='Q' OR servicio='S' or servicio='R' OR servicio='A' or servicio='D' or servicio='C') "
    buf2 = buf2 & " order by fecha "
    'MsgBox buf2
   
    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf2, cn, adOpenStatic, adLockOptimistic
    'If mysnapx.RecordCount = 0 Then
    '   GoTo a1
    'End If
    buf = ""
   
    Do

        If mysnapx.EOF Then Exit Do
        '

        ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
        signos = 1

        If "" & mysnapx.Fields("acu") = "E" Then  'nota de credito
            signos = 1

        End If

        ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
      
        sum1 = sum1 + 1
        fecha = "CABECERAS ..." & Format(sum1, "00000")
        buf = "" & mysnapx.Fields("salon")

        If Len(buf) = 0 Then
            buf = "0"

        End If
       
        fecha = "" & mysnapx.Fields("fecha")

        'If "" & mysnapx.Fields("acu") <> "S" And "" & mysnapx.Fields("acu") <> "T" Then  'entrdas /salidas
        'servicios
        'MsgBox "paso"
        '-----------------------------------
        If "" & mysnapx.Fields("estado") = "2" Then
            mytablez.Seek "=", "" & mysnapx.Fields("caja"), "" & mysnapx.Fields("turno"), "" & mysnapx.Fields("usuario")

            If Not mytablez.NoMatch Then
             
                mytablez.Edit
                suma_cajas mytablez, signos, mysnapx
                mytablez.Update

            End If

            If mytablez.NoMatch Then
                mytablez.AddNew
                suma_cajas mytablez, signos, mysnapx
                mytablez.Update

            End If

        End If

        '-----------------------------------
        mytablex.Seek "=", buf, "" & mysnapx.Fields("servicio")

        If Not mytablex.NoMatch Then
            mytablex.Edit
            suma_contador mytablex, signos, mysnapx
            mytablex.Update

        End If

        If mytablex.NoMatch Then
            mytablex.AddNew
            suma_contador mytablex, signos, mysnapx
            mytablex.Fields("local") = "01"
            mytablex.Update

        End If

        ' MsgBox ""
        'End If
        'documentos
        '--------------
        buf1 = "" & mysnapx.Fields("acu")
        mytabley.Seek "=", buf1 & "" & mysnapx.Fields("tipo")

        If Not mytabley.NoMatch Then
            mytabley.Edit
            suma_contador1 mytabley, signos, mysnapx
            mytabley.Fields("tipo") = "" & mysnapx.Fields("acu") & "" & mysnapx.Fields("tipo")
            mytabley.Update

        End If

        'MsgBox ""
        If mytabley.NoMatch Then
            mytabley.AddNew
          
            If opcion1 = "5" Then
          
                mytabley.Fields("cierre") = "" & busca_cierre("" & caja)
             
                mytabley.Fields("cajero") = "" & cajero
                mytabley.Fields("caja") = "" & caja
                mytabley.Fields("turno") = "" & turno
                mytabley.Fields("fecha") = Format(Now, "dd/mm/yyyy")
                mytabley.Fields("hora") = Format(Now, "hh:mm:ss")
          
            End If

            suma_contador1 mytabley, signos, mysnapx
            mytabley.Fields("tipo") = "" & mysnapx.Fields("acu") & "" & mysnapx.Fields("tipo")
            mytabley.Fields("local") = "01"
            mytabley.Update

        End If

        'MsgBox ""
        '--------------
        mysnapx.MoveNext
    Loop
a1:
    mysnapx.Close
    'MsgBox ""

    sum1 = 0
    mytablex.Index = "servicio"
    '---------- ingresos /egresos----------------------------------
    buf = "select * from " & dbing & " where  "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    'buf = buf & " and fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " order by fecha"

    'MsgBox buf
    'Set mysnap = mydbxglo.CreateSnapshot(buf)

    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic
   
    Do

        If mysnapx.EOF Then Exit Do
        signos = 1
        sum1 = sum1 + 1
        fecha = "INGRESOS/EGRESOS ..." & Format(sum1, "00000")
        buf1 = "" & mysnapx.Fields("servicio")

        If buf1 <> "W" And buf1 <> "V" Then GoTo a32
        mytablex.Seek "=", buf1 & "" & mysnapx.Fields("tipo")

        If Not mytablex.NoMatch Then
            mytablex.Edit
            suma_contador mytablex, signos, mysnapx
            mytablex.Fields("servicio") = "" & mysnapx.Fields("servicio") & "" & mysnapx.Fields("tipo")
            mytablex.Update

        End If

        If mytablex.NoMatch Then
            mytablex.AddNew
            suma_contador mytablex, signos, mysnapx
            mytablex.Fields("servicio") = "" & mysnapx.Fields("servicio") & "" & mysnapx.Fields("tipo")
            mytablex.Fields("local") = "01"
            mytablex.Update

        End If

        mysnapx.MoveNext
    Loop
a32:
    mysnapx.Close
    '--------------------------------------------------------------
    mytablex.Close
    mytabley.Close
    Exit Sub
cmd56_err:
    MsgBox "12***Mensaje,Error en servicio realizado " & buf & " " & error$, 24, "AVISO"
    mysnapx.Close
    mytablex.Close
    mytabley.Close
    Exit Sub

End Sub

'''14/09/2017 kenyo Mejor Reporte Ticket
Sub subtotal_documentos(buf1 As String, soles As Double, dolares As Double)

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    buf = " "
    found = formateaa(buf, 1, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = " "
    found = formateaa(buf, 1, 0, 0)
    found = formateaa("", 1, 0, 0)

    buf = buf1
    found = formateaa(buf, 12, 0, 0)
    found = formateaa("", 1, 0, 0)
       
    buf = Format(soles, "0.00")

    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
       
    buf = Format(dolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub
 
''''18/09/2017 kenyo Mejor Reporte Ticket
Sub subtotal_proddoc(soles As Double, dolares As Double)

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    found = formateaa("", 1, 2, 0)
    found = formateaa("", 1, 0, 0)
    buf = " TOTAL"
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(soles, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)

    'buf = "Dolares"
    'found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(dolares, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

''''18/09/2017 kenyo Mejor Reporte Ticket
Sub total_proddoc(soles As Double, dolares As Double)

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    found = formateaa("", 1, 2, 0)
    found = formateaa(">", 1, 0, 0)
    buf = "GRAN TOTAL"
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(soles, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)

    'buf = "Dolares"
    'found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(dolares, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Sub subtotal_unidades(buf1 As String, _
                      cantidad As Double, _
                      soles As Double, _
                      dolares As Double, _
                      vales As Double)

    Dim found As Integer

    Dim I     As Integer

    Dim buf   As String

    Dim sdx   As Double

    buf = buf1
       
    ''''13/09/2017 kenyo Mejor Reporte Ticket
    'found = formateaa(buf, 8, 0, 0)
    'found = formateaa("", 1, 0, 0)
    'found = formateaa("", 11, 0, 0)
    
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 10, 0, 0)
    ''''13/09/2017 kenyo Mejor Reporte Ticket

    buf = Format(cantidad, "0.00")
    found = formateaa(buf, 6, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf = Format(soles, "0.00")
    found = formateaa(buf, 8, 2, 1)

    If dolares > 0 Then
        found = formateaa("", 28, 0, 0)
        buf = Format(dolares, "0.00")
        found = formateaa(buf, 8, 2, 1)

    End If
       
    If vales > 0 Then
        found = formateaa(buf1 & " Vale", 27, 0, 0)
        buf = Format(vales, "0.00")
        found = formateaa(buf, 8, 2, 1)

    End If
       
    If UCase(buf1) = "<<TOTAL" And vales > 0 Then
        found = formateaa("TOTAL G.", 27, 0, 0)
        sdx = vales + soles
        buf = Format(sdx, "0.00")
        found = formateaa(buf, 8, 2, 1)

    End If

End Sub

Sub suma_contador(mytablex As Table, signos As Double, mysnapx As ADODB.Recordset)

    Dim sdx As Double

    Dim buf As String

    On Error GoTo cmd57_err

    If Val("" & mysnapx.Fields("tipo")) = 5 Then
        If todos <> "S" Then Exit Sub

    End If

    buf = "" & mysnapx.Fields("salon")

    If Len(buf) = 0 Then
        buf = "0"

    End If

    mytablex.Fields("servicio") = "" & mysnapx.Fields("servicio")
    mytablex.Fields("salon") = buf

    If Val("" & mysnapx.Fields("estado")) = 2 Then
        sdx = Val("" & mytablex.Fields("nro")) + 1
        mytablex.Fields("nro") = sdx

        If "" & mysnapx.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("valors")) + signos * Val("" & mysnapx.Fields("total"))
            mytablex.Fields("valors") = sdx

        End If

        If "" & mysnapx.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("valord")) + signos * Val("" & mysnapx.Fields("total"))
            mytablex.Fields("valord") = sdx

        End If

    End If

    If Val("" & mysnapx.Fields("estado")) = 1 Then
        sdx = Val("" & mytablex.Fields("nroa")) + 1
        mytablex.Fields("nroa") = sdx

        If "" & mysnapx.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("valorsa")) + signos * Val("" & mysnapx.Fields("total"))
            mytablex.Fields("valorsa") = sdx

        End If

        If "" & mysnapx.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("valorda")) + signos * Val("" & mysnapx.Fields("total"))
            mytablex.Fields("valorda") = sdx

        End If

    End If

    Exit Sub
cmd57_err:
    MsgBox "Error en suma contador " & error$, 24, "AVISO"
    Exit Sub

End Sub

Sub suma_contador1(mytablex As Table, signos As Double, mysnapx As ADODB.Recordset)

    Dim sdx As Double

    On Error GoTo cmd54311_err

    If Val("" & mysnapx.Fields("tipo")) = 5 Then
        If todos <> "S" Then Exit Sub

    End If

    mytablex.Fields("tipo") = "" & mysnapx.Fields("tipo")

    If Val("" & mysnapx.Fields("estado")) = 2 Then
        sdx = Val("" & mytablex.Fields("nro")) + 1
        mytablex.Fields("nro") = sdx

        If "" & mysnapx.Fields("moneda") = "S" Then
               
            sdx = Val("" & mytablex.Fields("brutos")) + signos * Val("" & mysnapx.Fields("neto"))
            mytablex.Fields("brutos") = sdx
            sdx = Val("" & mytablex.Fields("dsctos")) + signos * Val("" & mysnapx.Fields("descuento"))
            mytablex.Fields("dsctos") = sdx
            sdx = Val("" & mytablex.Fields("valorvs")) + signos * Val("" & mysnapx.Fields("subtotal"))
            mytablex.Fields("valorvs") = sdx
           
            '12/06/2017 kenyo NOTA DE CREDITO
            sdx = Val("" & mytablex.Fields("igvs")) + signos * Val("" & mysnapx.Fields("impuesto"))
            mytablex.Fields("igvs") = sdx
            '12/06/2017 kenyo NOTA DE CREDITO
           
            sdx = Val("" & mytablex.Fields("redons")) + signos * Val("" & mysnapx.Fields("redondeo"))
            mytablex.Fields("redons") = sdx
           
            sdx = Val("" & mytablex.Fields("sertots")) + signos * Val("" & mysnapx.Fields("servicioco"))
            mytablex.Fields("sertots") = sdx
            sdx = Val("" & mytablex.Fields("pertots")) + signos * Val("" & mysnapx.Fields("percepcion"))
            mytablex.Fields("pertots") = sdx
            sdx = Val("" & mytablex.Fields("exos")) + signos * Val("" & mysnapx.Fields("gravado"))
            mytablex.Fields("exos") = sdx
            sdx = Val("" & mytablex.Fields("tax1s")) + signos * Val("" & mysnapx.Fields("tisc"))
            mytablex.Fields("tax1s") = sdx
            sdx = Val("" & mytablex.Fields("valors")) + signos * Val("" & mysnapx.Fields("total"))
            mytablex.Fields("valors") = sdx
            'sdx = 0
            'sdx = Val("" & mytablex.Fields("retes")) + signos * Val("" & mysnapx.Fields("tretencion"))
            sdx = 0
            mytablex.Fields("retes") = sdx
            sdx = Val("" & mytablex.Fields("nodsctos")) + signos * Val("" & mysnapx.Fields("tivap"))
            mytablex.Fields("nodsctos") = sdx
           
            If "" & mysnapx.Fields("dflag") = "" Then
                sdx = Val("" & mytablex.Fields("cdetras")) + signos * Val("" & mysnapx.Fields("tdetra"))
                mytablex.Fields("cdetraS") = sdx

            End If

            If "" & mysnapx.Fields("dflag") = "1" Then
                sdx = Val("" & mytablex.Fields("ndetras")) + signos * Val("" & mysnapx.Fields("tdetra"))
                mytablex.Fields("ndetraS") = sdx

            End If

        End If

        If "" & mysnapx.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("valord")) + signos * Val("" & mysnapx.Fields("total"))
            mytablex.Fields("valord") = sdx
            sdx = Val("" & mytablex.Fields("valorvd")) + signos * Val("" & mysnapx.Fields("subtotal"))
            mytablex.Fields("valorvd") = sdx
            sdx = Val("" & mytablex.Fields("igvd")) + signos * Val("" & mysnapx.Fields("impuesto"))
            mytablex.Fields("igvd") = sdx
            sdx = Val("" & mytablex.Fields("exod")) + signos * Val("" & mysnapx.Fields("gravado"))
            mytablex.Fields("exod") = sdx
            sdx = Val("" & mytablex.Fields("tax1d")) + signos * Val("" & mysnapx.Fields("tisc"))
            'sdx = 0
            mytablex.Fields("tax1d") = sdx
            sdx = Val("" & mytablex.Fields("dsctod")) + signos * Val("" & mysnapx.Fields("descuento"))
            'sdx = 0
            mytablex.Fields("dsctod") = sdx
            'sdx = Val("" & mytablex.Fields("reted")) + signos * Val("" & mysnapx.Fields("tretencion"))
           
            sdx = Val("" & mytablex.Fields("sertotd")) + signos * Val("" & mysnapx.Fields("servicioco"))
            mytablex.Fields("sertotd") = sdx
            sdx = Val("" & mytablex.Fields("pertotd")) + signos * Val("" & mysnapx.Fields("percepcion"))
            mytablex.Fields("pertotd") = sdx
          
            sdx = Val("" & mytablex.Fields("redond")) + signos * Val("" & mysnapx.Fields("redondeo"))
            mytablex.Fields("redond") = sdx
           
            sdx = 0
            mytablex.Fields("reted") = sdx
            sdx = Val("" & mytablex.Fields("nodsctod")) + signos * Val("" & mysnapx.Fields("tivap"))
            sdx = 0
            mytablex.Fields("nodsctod") = sdx
            sdx = Val("" & mytablex.Fields("brutod")) + signos * Val("" & mysnapx.Fields("neto"))
            mytablex.Fields("brutod") = sdx

            If "" & mysnapx.Fields("dflag") = "" Then
                sdx = Val("" & mytablex.Fields("cdetrad")) + signos * Val("" & mysnapx.Fields("tdetra"))
                mytablex.Fields("cdetrad") = sdx

            End If

            If "" & mysnapx.Fields("dflag") = "1" Then
                sdx = Val("" & mytablex.Fields("ndetrad")) + signos * Val("" & mysnapx.Fields("tdetra"))
                mytablex.Fields("ndetrad") = sdx

            End If

        End If

    End If

    If Val("" & mysnapx.Fields("estado")) = 1 Then
        sdx = Val("" & mytablex.Fields("nroa")) + 1
        mytablex.Fields("nroa") = sdx

        If "" & mysnapx.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("valorsa")) + signos * Val("" & mysnapx.Fields("total"))
            mytablex.Fields("valorsa") = sdx

        End If

        If "" & mysnapx.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("valorda")) + signos * Val("" & mysnapx.Fields("total"))
            mytablex.Fields("valorda") = sdx

        End If

    End If

    Exit Sub
cmd54311_err:
    MsgBox "Error en suma_contador 1" + error, 48, "Aviso"
    Exit Sub

End Sub

Function suma_fpago(buf As String, _
                    mytablex As Table, _
                    signos As Double, _
                    mysnapx As ADODB.Recordset) As Double

    Dim sdx  As Double

    Dim buf1 As String

    On Error GoTo cmd4556_err

    suma_fpago = 0

    If Val("" & mysnapx.Fields("tipo")) = 5 Then
        If todos <> "S" Then Exit Function

    End If

    If Val("" & mysnapx.Fields("estado")) = 2 Then
        mytablex.Fields("tipo") = "" & mysnapx.Fields("fpago")
        mytablex.Fields("servicio") = buf
        sdx = Val("" & mytablex.Fields("nro")) + 1
        mytablex.Fields("nro") = sdx

        If "" & mysnapx.Fields("moneda") = "S" Then
            If Val("" & mysnapx.Fields("saldos")) <= 0 Then
                mytablex.Fields("valors") = Val("" & mytablex.Fields("valors")) + signos * Val(Format(Val("" & mysnapx.Fields("recibe")) + Val("" & mysnapx.Fields("saldos")), "0.00"))
            Else
                mytablex.Fields("valors") = Val("" & mytablex.Fields("valors")) + signos * Val(Format(Val("" & mysnapx.Fields("recibe"))))

            End If

        End If

        If "" & mysnapx.Fields("moneda") = "D" Then
            mytablex.Fields("valord") = Val("" & mytablex.Fields("valord")) + signos * Val(Format(Val("" & mysnapx.Fields("recibe")), "0.00"))
            buf1 = Format(Val("" & mysnapx.Fields("saldos")), "0.00")
            'mytablex.Fields("valord") = Val("" & mytablex.Fields("valord")) + signos * Val("" & mysnapx.Fields("recibed"))
            suma_fpago = Val(buf1)

        End If

    End If

    Exit Function
cmd4556_err:
    MsgBox "Error en Suma Fpago " & error$, 24, "Aviso"
    Exit Function
      
End Function

Function suma_las_ventas() As Double

    Dim buf     As String

    Dim cfechai As String

    Dim cfechaf As String

    Dim mysnap1 As New ADODB.Recordset

    On Error GoTo cmd899_err

    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    cfechai = "01/"
    cfechai = cfechai & Format(Month(fechai), "00") & "/"
    cfechai = cfechai & Format(Year(fechai), "0000")
    cfechaf = Format(CVDate(fechaf) - 1, "dd/mm/yyyy")

    If Not IsDate(cfechai) Then Exit Function
    If Not IsDate(cfechaf) Then Exit Function
   
    buf = "select sum(total) as TOT from " & dbca & " where  val(estado)=2  "
    buf = buf & "  and fecha>='" & Format(cfechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(cfechaf, "YYYYMMDD") & "' "
   
    'buf = buf & " and fecha>=" & "DateValue('" & cfechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & cfechaf & "'" & ")"
    buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E') "  'E nota credito
   
    mysnap1.Open buf, cn, adOpenStatic, adLockOptimistic

    If mysnap1.RecordCount > 0 Then
        suma_las_ventas = Val("" & mysnap1.Fields("TOT"))

    End If

    mysnap1.Close
   
    'Set mysnap1 = Nothing
    'Set mydb9 = Nothing
    Exit Function
cmd899_err:
    MsgBox "Error en Suma las Ventas " & error$, 24, "Aviso"
    mysnap1.Close
   
    Exit Function

End Function

Sub suma_productos(Grupo As String, mytable1x As ADODB.Recordset, mytable2x As Table)

    Dim sdx    As Double

    Dim signos As Double

    On Error GoTo cmd3490_err

    signos = 1
    'If "" & mytable1.Fields("acu") = "D" Then  'si es nota de credito ojo en el nuevo es otro
    '   MsgBox "hola"
    '   signos = -1
    'End If
    mytable2x.Fields("sentido") = "" & mytable1x.Fields("sentido")

    If check3d2.Value = 0 Then
        mytable2x.Fields("producto") = "" & mytable1x.Fields("producto")
        mytable2x.Fields("descripcio") = Mid$("" & mytable1x.Fields("descripcio"), 1, 20)

    End If

    If check3d2.Value = 1 Then
        mytable2x.Fields("producto") = Grupo
        mytable2x.Fields("descripcio") = Grupo

    End If

    If check3d3.Value = 1 Then
        mytable2x.Fields("producto") = Grupo
        mytable2x.Fields("descripcio") = Grupo

    End If

    'MsgBox grupo
    mytable2x.Fields("grupo") = Grupo
    mytable2x.Fields("unidad") = "" & mytable1x.Fields("unidad")

    If Val("" & mytable1x.Fields("estado")) = 2 Then
        If Val("" & mytable1x.Fields("tipo")) <> 7 And Val("" & mytable1x.Fields("tipo")) <> 6 Then
            sdx = Val("" & mytable2x.Fields("cantidad")) + signos * Val("" & mytable1x.Fields("cantidad")) * Val("" & mytable1x.Fields("factor"))
            mytable2x.Fields("cantidad") = sdx

        End If

        If Val("" & mytable1x.Fields("tipo")) = 7 Then
            sdx = Val("" & mytable2x.Fields("exonerado")) + signos * Val("" & mytable1x.Fields("cantidad")) * Val("" & mytable1x.Fields("factor"))
            mytable2x.Fields("exonerado") = sdx

        End If

        If Val("" & mytable1x.Fields("tipo")) = 6 Then
            sdx = Val("" & mytable2x.Fields("vales")) + signos * Val("" & mytable1x.Fields("cantidad")) * Val("" & mytable1x.Fields("factor"))
            mytable2x.Fields("vales") = sdx

        End If

        If "" & mytable1x.Fields("moneda") = "S" Then
            If Val("" & mytable1x.Fields("tipo")) = 6 Then
                sdx = Val("" & mytable2x.Fields("totalVALES")) + signos * Val("" & mytable1x.Fields("total"))
                mytable2x.Fields("totalVALES") = sdx

            End If

            If Val("" & mytable1x.Fields("tipo")) <> 6 Then
                sdx = Val("" & mytable2x.Fields("totals")) + signos * Val("" & mytable1x.Fields("total"))
                mytable2x.Fields("totals") = sdx

            End If

        End If

        If "" & mytable1x.Fields("moneda") = "D" Then
            If Val("" & mytable1x.Fields("tipo")) <> 6 Then
                sdx = Val("" & mytable2x.Fields("totald")) + signos * Val("" & mytable1x.Fields("total"))
                mytable2x.Fields("totald") = sdx

            End If

        End If

    End If

    If Val("" & mytable1x.Fields("estado")) = 1 Then
        sdx = Val("" & mytable2x.Fields("cantidada")) + signos * Val("" & mytable1x.Fields("cantidad")) * Val("" & mytable1x.Fields("factor"))
        mytable2x.Fields("cantidada") = sdx

        If "" & mytable1x.Fields("moneda") = "S" Then
            sdx = Val("" & mytable2x.Fields("totalsa")) + signos * Val("" & mytable1x.Fields("total"))
            mytable2x.Fields("totalsa") = sdx

        End If

        If "" & mytable1x.Fields("moneda") = "D" Then
            sdx = Val("" & mytable2x.Fields("totalda")) + signos * Val("" & mytable1x.Fields("total"))
            mytable2x.Fields("totalda") = sdx

        End If

    End If

    Exit Sub
cmd3490_err:
    MsgBox "Error en suma_productos", 24, "Aviso"
    Exit Sub

End Sub

Private Sub Label2_Click()

    If todos = "N" Then
        todos = "S"
        fechaf_KeyPress (13)
        Exit Sub

    End If
    
    If todos = "S" Then
        todos = "N"
        fechaf_KeyPress (13)
        Exit Sub

    End If

End Sub

Private Sub Label8_Click()

    If flag_proveedor = "S" Then
        flag_proveedor = ""
        Exit Sub

    End If

    If flag_proveedor = "" Then
        flag_proveedor = "S"
        Exit Sub

    End If

End Sub

Private Sub localc_Click()

    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja..Crear combo cajac en tcuadrc1
    'local1 = extra_loquesea(localc)
    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja..Crear combo cajac en tcuadrc1
End Sub

Sub turno_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    horai.SetFocus

End Sub

Sub turno_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = &H26 Then
    '   caja.SetFocus
    '   Exit Sub
    'End If

End Sub

Sub unidades_vendidas()

    Dim mytable1x As New ADODB.Recordset

    Dim mytable3x As New ADODB.Recordset

    Dim vr, buf, buf1 As String

    Dim Grupo     As String

    Dim xcajero   As String

    Dim xturno    As String

    Dim xcaja     As String

    Dim mytable2x As Table

    Dim buf2      As String

    Dim signos    As Double

    On Error GoTo cmd488_err

    Dim sdx As Double

    sum1 = 0
    borrar_cuadres
   
    'MsgBox "abc"
    Set mytable2x = mydbxglo.OpenTable(usuariopos & "04")  'cuadre 04
    buf2 = "select * from " & dbde & " where "
    buf2 = buf2 & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf2 = buf2 & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf2 = buf2 & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf2 = buf2 & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf2 = buf2 & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf2 = buf2 & " and turno like '" & turno & "%'"

    End If

    If flag_proveedor = "S" Then
        buf = buf & " and len(proveedorp)>0 "

    End If

    buf2 = buf2 & " and ( acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='1') "  'E nota credito
    buf2 = buf2 & " order by fecha "
   
    'MsgBox buf2
    mytable1x.Open buf2, cn, adOpenStatic, adLockOptimistic

    If mytable1x.RecordCount = 0 Then
        Exit Sub

    End If
   
    Do

        If mytable1x.EOF Then Exit Do
        fecha = "UNIDADES VENDIDAS ...." & "" & mytable1x.Fields("numero")
        vr = DoEvents()
        sum1 = sum1 + 1

        If Val("" & mytable1x.Fields("tipo")) = 5 Then
            If todos <> "S" Then
                GoTo a2o

            End If

        End If
         
        buf1 = "" & mytable1x.Fields("servicio")

        'OJO VER ESTO NO DEBE PASAR
        If buf1 <> "A" And buf1 <> "C" And buf1 <> "D" Then GoTo a2o
        Grupo = "NT"

        If Check3d4.Value = 1 Then
            Grupo = ""
            GoTo amkiko

        End If

        If mytable3x.State = 1 Then mytable3x.Close
        mytable3x.Open "select * from producto where producto='" & Trim("" & mytable1x.Fields("producto")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytable3x.RecordCount > 0 Then
            Grupo = "" & mytable3x.Fields("familia")

            If check3d3.Value = 1 Then
                'MsgBox ""
                Grupo = "" & mytable3x.Fields("seccion")

                'MsgBox "xxx"
            End If

        End If

        mytable3x.Close
amkiko:
               
        'servicios
        mytable2x.Index = "producto"
        mytable2x.Seek "=", "" & mytable1x.Fields("producto"), "" & mytable1x.Fields("sentido")

        If check3d2 = 1 Or check3d3 = 1 Then
            mytable2x.Seek "=", Grupo, "" & mytable1x.Fields("sentido")

        End If

        If mytable2x.NoMatch Then
            mytable2x.AddNew

            If opcion1 = "5" Then
                'mytabley.fields("cierre") = ""
                mytable2x.Fields("cierre") = busca_cierre(xcaja)
                mytable2x.Fields("cajero") = "" & cajero
                mytable2x.Fields("caja") = "" & caja
                mytable2x.Fields("turno") = "" & turno
                mytable2x.Fields("fecha") = Format(Now, "dd/mm/yyyy")
                mytable2x.Fields("hora") = Format(Now, "hh:mm:ss")

            End If

            suma_productos Grupo, mytable1x, mytable2x
            mytable2x.Fields("local") = "01"
            mytable2x.Update

        End If

        If Not mytable2x.NoMatch Then
            mytable2x.Edit
            suma_productos Grupo, mytable1x, mytable2x
            mytable2x.Update

        End If

a2o:
        mytable1x.MoveNext
    Loop
    mytable1x.Close
    mytable2x.Close
    Exit Sub
cmd488_err:
    MsgBox "0.Mensaje,Error en unidades vendidas " & error$, 48, "Aviso"
  
    mytable1x.Close
    mytable2x.Close
    Exit Sub

End Sub

Sub ver_cajeros()

    Dim buf      As String

    Dim buf1     As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd2caj_err

    If cajero = "%" Or caja = "%" Or turno = "%" And opcion1 = "5" Then
        buf1 = "select * from apertura where  cajero like '" & cajero & "'" & " and caja like '" & caja & "'" & " and turno like '" & turno & "'"
   
        buf1 = buf1 & "  and fechai>='" & Format(fechai, "YYYYMMDD") & "'"
        buf1 = buf1 & " and fechaf<='" & Format(fechaf, "YYYYMMDD") & "' "
   
        'buf1 = buf1 & " and fechai>=" & "DateValue('" & fechai & "'" & ")"
        'buf1 = buf1 & " and fechai<=" & "DateValue('" & fechaf & "'" & ")"
   
        mytablex.Open buf1, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Sub

        End If
   
        'Set mytablex = mydbxglo.CreateSnapshot(buf1)
        Do

            If mytablex.EOF Then Exit Do
      
            buf = "" & mytablex.Fields("cajero")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
       
            buf = "" & mytablex.Fields("caja")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)

            buf = "" & mytablex.Fields("turno")
            found = formateaa(buf, 5, 2, 0)
            mytablex.MoveNext
        Loop
        mytablex.Close

    End If

    Exit Sub
cmd2caj_err:
    MsgBox "Aviso en ver cajeros " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub verifica_tradiario()

    Dim mytablex As New ADODB.Recordset

    tradiario = ""

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        tradiario = "" & mytablex.Fields("tradiario")

    End If

    mytablex.Close
    Exit Sub

End Sub

Sub visualiza_cajeros()

    Dim buf   As String

    Dim buf1  As String

    Dim buf2  As String

    Dim buf3  As String

    Dim found As Integer

    On Error GoTo cmd1_err:

    'If opcion1 <> "5" Then Exit Sub
  
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("1", buf2, buf3)
    ' Testing Proyecto Facturacion Electronica 01/03/2018
    'buf = "Ticket Factura:  " & buf2
    buf = busca_tipo2(1) & ":  " & buf2
    ' Testing Proyecto Facturacion Electronica 01/03/2018
    found = formateaa(buf, 25, 0, 0)
    buf1 = busca_inicio("1", buf2, buf3)

    If buf2 <> "" Then
        buf = "al  " & buf3
    Else
        buf = "    " & buf3

    End If
      
    found = formateaa(buf, 10, 0, 0)
    'found = formateaa("", 5, 0, 0)
    found = formateaa("", 1, 2, 0)
  
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("2", buf2, buf3)
      
    ' Testing Proyecto Facturacion Electronica 01/03/2018
    'buf = "Ticket Factura:  " & buf2
    buf = busca_tipo2(2) & ":  " & buf2
    ' Testing Proyecto Facturacion Electronica 01/03/2018
      
    found = formateaa(buf, 25, 0, 0)
    buf1 = busca_inicio("2", buf2, buf3)

    If buf2 <> "" Then
        buf = "al  " & buf3
    Else
        buf = "    " & buf3

    End If

    found = formateaa(buf, 10, 0, 0)
    'found = formateaa("", 5, 0, 0)
    found = formateaa("", 1, 2, 0)
      
    If todos = "S" Then
        buf2 = ""
        buf3 = ""
        buf1 = busca_inicio("5", buf2, buf3)
        buf = "Nota de Venta :  " & buf2
        found = formateaa(buf, 25, 0, 0)
        buf1 = busca_inicio("5", buf2, buf3)

        If buf2 <> "" Then
            buf = "al  " & buf3
        Else
            buf = "    " & buf3

        End If

        found = formateaa(buf, 10, 0, 0)
        'found = formateaa("", 5, 0, 0)
        found = formateaa("", 1, 2, 0)

    End If
    
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("3", buf2, buf3)
    buf = "Boleta Manual :  " & buf2
    found = formateaa(buf, 25, 0, 0)
    buf1 = busca_inicio("3", buf2, buf3)

    If buf2 <> "" Then
        buf = "al  " & buf3
    Else
        buf = "    " & buf3

    End If

    found = formateaa(buf, 10, 0, 0)
    'found = formateaa("", 5, 0, 0)
    found = formateaa("", 1, 2, 0)
      
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("4", buf2, buf3)
    buf = "Factura Manual:  " & buf2
    found = formateaa(buf, 25, 0, 0)
    buf1 = busca_inicio("4", buf2, buf3)

    If buf2 <> "" Then
        buf = "al  " & buf3
    Else
        buf = "    " & buf3

    End If

    found = formateaa(buf, 10, 0, 0)
    'found = formateaa("", 5, 0, 0)
    found = formateaa("", 1, 2, 0)
    
    'NOTA DE CREDITO
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("71", buf2, buf3)
    buf = "NotaCredBoleta:    " & buf2

    found = formateaa(buf, 25, 0, 0)
    buf1 = busca_inicio("71", buf2, buf3)

    If buf2 <> "" Then
        buf = "al  " & buf3
    Else
        buf = "    " & buf3

    End If

    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)
    'found = formateaa("", 5, 0, 0)
      
    ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("72", buf2, buf3)
    buf = "NotaCredFactur:    " & buf2
    found = formateaa(buf, 25, 0, 0)
    buf1 = busca_inicio("72", buf2, buf3)

    If buf2 <> "" Then
        buf = "al  " & buf3
    Else
        buf = "    " & buf3

    End If

    found = formateaa(buf, 10, 0, 0)
    'found = formateaa("", 5, 0, 0)
    ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
      
    found = formateaa("", 1, 2, 0)
    
    '
    '
    '
    '
    '    buf2 = ""
    '    buf3 = ""
    '    buf1 = busca_inicio("1", buf2, buf3)
    '    buf = "- TBoleta-Inicial        :" & buf2
    '    found = formateaa(buf, 34, 0, 0)
    '    found = formateaa("", 5, 0, 0)
    '    found = formateaa("", 1, 2, 0)
    '    buf = "  TBoleta-Final          :" & buf3
    '    found = formateaa(buf, 34, 0, 0)
    '
    '
    '    buf2 = ""
    '    buf3 = ""
    '    buf1 = busca_inicio("2", buf2, buf3)
    '    'espacio enter
    '    found = formateaa("", 1, 15, 0)
    '    found = formateaa("", 1, 2, 0)
    '    'espacio enter
    '    buf = "- TFactura-Inicial       :" & buf2
    '    found = formateaa(buf, 34, 0, 0)
    '    found = formateaa("", 5, 0, 0)
    '    found = formateaa("", 1, 2, 0)
    '    buf = "  TFactura-Final         :" & buf3
    '    found = formateaa(buf, 34, 0, 0)
    '
    '    buf2 = ""
    '    buf3 = ""
    '    buf1 = busca_inicio("5", buf2, buf3)
    '    'espacio enter
    '    found = formateaa("", 1, 15, 0)
    '    found = formateaa("", 1, 2, 0)
    '    'espacio enter
    '    buf = "- TNota-Inicial          :" & buf2
    '    found = formateaa(buf, 34, 0, 0)
    '    found = formateaa("", 5, 0, 0)
    '    found = formateaa("", 1, 2, 0)
    '    buf = "  TNota-Final            :" & buf3
    '    found = formateaa(buf, 34, 0, 0)
    '
    '    buf2 = ""
    '    buf3 = ""
    '    buf1 = busca_inicio("3", buf2, buf3)
    '    'espacio enter
    '    found = formateaa("", 1, 15, 0)
    '    found = formateaa("", 1, 2, 0)
    '    'espacio enter
    '    buf = "- BolManual-Inicial      :" & buf2
    '    found = formateaa(buf, 34, 0, 0)
    '    found = formateaa("", 5, 0, 0)
    '    found = formateaa("", 1, 2, 0)
    '    buf = "  BolManual-Final        :" & buf3
    '    found = formateaa(buf, 34, 0, 0)
    '
    '    buf2 = ""
    '    buf3 = ""
    '    buf1 = busca_inicio("4", buf2, buf3)
    '    'espacio enter
    '    found = formateaa("", 1, 15, 0)
    '    found = formateaa("", 1, 2, 0)
    '    'espacio enter
    '    buf = "- FacManual-Inicial      :" & buf2
    '    found = formateaa(buf, 34, 0, 0)
    '    found = formateaa("", 5, 0, 0)
    '    found = formateaa("", 1, 2, 0)
    '    buf = "  FacManual-Final        :" & buf3
    '    found = formateaa(buf, 34, 0, 0)
    '    buf2 = ""
    '    buf3 = ""
    '
    '    found = formateaa("", 1, 2, 0)
    '    found = formateaa(" ", 40, 2, 0)
  
    'comentario 22/05/2017 kenyo
    
    '    buf1 = busca_inicio("9", buf2, buf3)
    '     buf1 = busca_inicio("2", buf2, buf3)
    '    found = formateaa("", 1, 15, 0)
    '       buf = "NC-I:" & buf2
    '       found = formateaa(buf, 16, 0, 0)
    '       found = formateaa("", 1, 0, 0)
    '       buf = "NC-F:" & buf3
    '       found = formateaa(buf, 16, 0, 0)
    '       found = formateaa("", 1, 2, 0)
       
    Exit Sub
cmd1_err:
    Exit Sub

End Sub

'''18/09/2017 KENYO Formato Simple Cierre X & Y
Sub visualiza_cajerosSimple()

    Dim buf   As String

    Dim buf1  As String

    Dim buf2  As String

    Dim buf3  As String

    Dim found As Integer

    On Error GoTo cmd1_err:

    'If opcion1 <> "5" Then Exit Sub
     
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("1", buf2, buf3)
    buf = "Ticket Inicio :  " & buf2
    found = formateaa(buf, 25, 0, 0)
    buf1 = busca_inicio("1", buf2, "")

    found = formateaa("", 1, 2, 0)
    buf2 = ""
    buf3 = ""

    buf = "Ticket Final:  "
    found = formateaa(buf, 17, 0, 0)
    buf1 = busca_inicio("1", buf2, buf3)

    If buf2 <> "" Then
        buf = buf3

    End If

    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 5, 0, 0)
    found = formateaa("", 1, 2, 0)
              
    Exit Sub
cmd1_err:
    Exit Sub

End Sub

'''18/09/2017 KENYO Formato Simple Cierre X & Y

Function creando_cuadres(buf2 As String)

    Dim found      As Integer

    Dim globaldat1 As String

    Dim buf        As String

    On Error GoTo cmd56rre_err

    buf = buf2
    globaldat1 = globaldat & "\"
    copiando globaldat1 & "cuadre01.dbf", globaldat1 & buf & "01.dbf"
    copiando globaldat1 & "cuadre01.cdx", globaldat1 & buf & "01.cdx"
    copiando globaldat1 & "cuadre02.dbf", globaldat1 & buf & "02.dbf"
    copiando globaldat1 & "cuadre02.cdx", globaldat1 & buf & "02.cdx"
    copiando globaldat1 & "cuadre03.dbf", globaldat1 & buf & "03.dbf"
    copiando globaldat1 & "cuadre03.cdx", globaldat1 & buf & "03.cdx"
    copiando globaldat1 & "cuadre04.dbf", globaldat1 & buf & "04.dbf"
    copiando globaldat1 & "cuadre04.cdx", globaldat1 & buf & "04.cdx"
    copiando globaldat1 & "cuadre05.dbf", globaldat1 & buf & "05.dbf"
    copiando globaldat1 & "cuadre05.cdx", globaldat1 & buf & "05.cdx"
    creando_cuadres = 1
    Exit Function
cmd56rre_err:

    MsgBox "Por favor Llame a servicio tecnico", 24, "Aviso"
    Exit Function

End Function

Function busca_fpago_ordenes(mysnapx As Snapshot) As Double

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    sdx = 0
    mytablex.Open "select * from fpagov where local='" & "" & mysnapx.Fields("local") & "' and tipo='" & "" & mysnapx.Fields("tipo") & "' and serie='" & "" & mysnapx.Fields("serie") & "' and numero='" & "" & mysnapx.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            If "" & mytablex.Fields("local") = "" & mysnapx.Fields("local") And "" & mytablex.Fields("tipo") = "" & mysnapx.Fields("tipo") And "" & mytablex.Fields("serie") = "" & mysnapx.Fields("serie") And "" & mytablex.Fields("numero") = "" & mysnapx.Fields("numero") Then
                sdx = sdx + Val("" & mytablex.Fields("recibe"))

            End If

            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    busca_fpago_ordenes = sdx
    Exit Function

End Function

Function valida_caja()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parameca where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_caja = 1

    End If

    mytablex.Close
    Exit Function

End Function

Function verifica_grupo() As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        verifica_grupo = "" & mytablex.Fields("grupoproducto")

    End If

    mytablex.Close
    Exit Function

End Function

Sub imprime_orden_trabajo()

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim js       As Double

    Dim jd       As Double

    Dim sdx      As Double

    Dim jindx    As Double

    Dim xsolesx  As Double

    Dim xdolarx  As Double

    On Error GoTo cmd891213

    jindx = 0
    js = 0
    jd = 0
    buf = "select * from cpedidov where "
    buf = buf & "   fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    'buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    'buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"
   
    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "%'"

    End If

    buf = buf & " and acu='I'"
    'buf = buf & " and yausado<>'1'"
    buf = buf & " order by str(numero) "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount = 0 Then
    'End If
   
    Do

        If mytablex.EOF Then Exit Do
        jindx = jindx + 1
        'If "" & mytablex.Fields("moneda") = "S" Then
        xsolesx = Val("" & mytablex.Fields("acuenta"))
        js = js + Val("" & mytablex.Fields("total"))
        jd = jd + Val("" & mytablex.Fields("acuenta"))
        'End If
        'If "" & mytablex.Fields("moneda") = "D" Then
        '   jd = jd + Val("" & mytablex.Fields("acuenta"))
        '   xdolarx = Val("" & mytablex.Fields("acuenta"))
        'End If
       
        found = formateaa("" & mytablex.Fields("numero"), 11, 0, 0)
        found = formateaa("", 2, 0, 0)
       
        sdx = Val("" & mytablex.Fields("total"))
        found = formateaa("" & sdx, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
       
        sdx = Val("" & mytablex.Fields("acuenta"))
        found = formateaa("" & sdx, 8, 0, 1)
        found = formateaa("", 1, 2, 1)
       
        'buf = ""
        'found = formateaa(buf, 6, 0, 1)
        'found = formateaa("", 1, 0, 0)
       
        'buf = "" & xsolesx
        'buf = Format(Val(buf), "0.00")
        'found = formateaa(buf, 8, 0, 1)
        'found = formateaa("", 1, 0, 0)
        'buf = "" & xdolarx
        'buf = Format(Val(buf), "0.00")
        'found = formateaa(buf, 8, 0, 1)
        'found = formateaa("", 1, 2, 0)
        mytablex.MoveNext
    Loop
    mytablex.Close
      
    found = formateaa("TotalOrden", 14, 0, 0)
    buf = Format(js, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(jd, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
       
    'sum1 = sum1 + jd
    'sum2 = suma2 + 0
    Exit Sub
cmd891213:
    Exit Sub
   
End Sub

Function servicio_tabla(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from servicio where servicio='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        servicio_tabla = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

''''13/09/2017 kenyo Mejor Reporte Ticket
Sub unidades_vendedor()

    Dim found    As Integer

    Dim buf2     As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim Tmp      As String

    Dim buf      As String

    Dim I        As Integer

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim ssdx     As Double

    Dim ssdx1    As Double

    sw = 0
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
    ncanal = 1
    Open FileName For Append As #ncanal
    cabecera "VENDEDOR PRODUCTO"
    buf = "PRODUC"
    found = formateaa(buf, 14, 0, 0)
    buf = "DESC"
    found = formateaa(buf, 6, 0, 0)
    buf = "CANT"
    found = formateaa(buf, 7, 0, 0)
    buf = "TOTAL"
    found = formateaa(buf, 9, 2, 0)
    buf = String(38, "-")
    found = formateaa(buf, 35, 2, 0)
       
    sdx = 0
    sdx1 = 0
    ssdx = 0
    ssdx1 = 0
    buf2 = "select vendedor,producto,sum(cantidad*factor) as xcan,sum(total) as xtot from detalle where "
    buf2 = buf2 & "   fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf2 = buf2 & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf2 = buf2 & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf2 = buf2 & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf2 = buf2 & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf2 = buf2 & " and turno like '" & turno & "%'"

    End If

    buf2 = buf2 & " and estado='2'"
    buf2 = buf2 & " and ( acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='1') "  'E nota credito
    buf2 = buf2 & " group by vendedor,producto  order by vendedor,producto "
   
    'MsgBox buf2
    mytablex.Open buf2, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sw = 0
        Do

            If mytablex.EOF Then Exit Do
            If sw = 0 Then
                Tmp = "" & mytablex.Fields("vendedor")
                sdx = 0
                sdx1 = 0
                sw = 1
                ''''13/09/2017 kenyo Mejor Reporte Ticket
                'buf = "" & mytablex.Fields("vendedor")
                'found = formateaa(buf, 5, 0, 0)
                'buf = busca_vendedor("" & mytablex.Fields("vendedor"))
                'found = formateaa(buf, 10, 2, 0)
             
                ''''13/09/2017 kenyo Mejor Reporte Ticket
                found = formateaa("", 1, 2, 0)
                found = formateaa(">>>", 4, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("vendedor"))
                found = formateaa(buf, 20, 2, 0)
               
            End If

            If Tmp <> "" & mytablex.Fields("vendedor") Then
                buf = Format(sdx, "0.00")
                found = formateaa("TOTAL", 18, 0, 0)
                found = formateaa(buf, 7, 0, 1)
                found = formateaa("", 3, 0, 0)
         
                buf = Format(sdx1, "0.00")
         
                '13/09/2017 kenyo Mejor Reporte Ticket
                'found = formateaa(buf, 5, 2, 1)
                '13/09/2017 kenyo Mejor Reporte Ticket
         
                Tmp = "" & mytablex.Fields("vendedor")
         
                found = formateaa(buf, 5, 0, 0)
         
                ''''13/09/2017 kenyo Mejor Reporte Ticket
                found = formateaa("", 1, 2, 0)
                found = formateaa("", 1, 2, 0)
                found = formateaa(">>>", 4, 0, 0)
                ''''13/09/2017 kenyo Mejor Reporte Ticket
         
                buf = busca_vendedor("" & mytablex.Fields("vendedor"))
                found = formateaa(buf, 10, 2, 0)
         
                sdx = 0
                sdx1 = 0

            End If

            found = formateaa("*", 1, 0, 0)
            buf = "" & mytablex.Fields("producto")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = busca_nombrexx("" & mytablex.Fields("producto"))
            '13/09/2017 kenyo Mejor Reporte Ticket
            'found = formateaa(buf, 15, 0, 0)
            'found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 27, 0, 0)
            found = formateaa("", 1, 2, 0)
       
            buf = busca_nombrexxExtra("" & mytablex.Fields("producto"))
            found = formateaa(buf, 17, 0, 0)
            'found = formateaa("", 1, 2, 0)
       
            'found = formateaa("", 20, 0, 0)
            '13/09/2017 kenyo Mejor Reporte Ticket
      
            found = formateaa("", 1, 0, 0)
      
            buf = "" & mytablex.Fields("xcan")
      
            buf = Format(buf, "0.00")
        
            found = formateaa(buf, 7, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("xtot")
            buf = Format(buf, "0.00")
            found = formateaa(buf, 7, 0, 1)
            found = formateaa("", 1, 2, 0)
            sdx = sdx + Val("" & mytablex.Fields("xcan"))
            sdx1 = sdx1 + Val("" & mytablex.Fields("xtot"))
      
            ssdx = ssdx + Val("" & mytablex.Fields("xcan"))
            ssdx1 = ssdx1 + Val("" & mytablex.Fields("xtot"))
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    buf = Format(sdx, "0.00")
    found = formateaa("TOTAL", 18, 0, 0)
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 3, 0, 0)
    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 7, 2, 1)
    buf = Format(ssdx, "0.00")
         
    found = formateaa("", 7, 2, 1)
    found = formateaa("GRAN TOTAL", 18, 0, 0)
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 3, 0, 0)
    buf = Format(ssdx1, "0.00")
    found = formateaa(buf, 7, 2, 1)
        
    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    Close #ncanal
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

''''13/09/2017 kenyo Mejor Reporte Ticket

Sub unidades_vendedorp()

    Dim found    As Integer

    Dim buf2     As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim Tmp      As String

    Dim buf      As String

    Dim I        As Integer

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim ssdx     As Double

    Dim ssdx1    As Double

    sw = 0
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
    ncanal = 1
    Open FileName For Append As #ncanal
    cabecera "VENDEDOR PRODUCTO"
    buf = "PRODUC"
    found = formateaa(buf, 14, 0, 0)
    buf = "DESC"
    found = formateaa(buf, 6, 0, 0)
    buf = "CANT"
    found = formateaa(buf, 7, 0, 0)
    buf = "TOTAL"
    found = formateaa(buf, 9, 2, 0)
    buf = String(38, "-")
    found = formateaa(buf, 35, 2, 0)

    sdx = 0
    sdx1 = 0
    ssdx = 0
    ssdx1 = 0

    buf2 = "select producto.productoequ as yproducto,producto.producto as zproducto,sum(detalle.cantidad*detalle.factor) as xcan,sum(detalle.total) as xtot from producto inner join detalle on producto.producto=detalle.producto  "
    buf2 = buf2 & " and  detalle.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf2 = buf2 & " and detalle.fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf2 = buf2 & " and detalle.local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf2 = buf2 & " and detalle.usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf2 = buf2 & " and detalle.caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf2 = buf2 & " and detalle.turno like '" & turno & "%'"

    End If

    buf2 = buf2 & " and producto.productoequ>0 "
    buf2 = buf2 & " and detalle.estado='2'"
    buf2 = buf2 & " and ( detalle.acu='A' or detalle.acu='B' or detalle.acu='C' or detalle.acu='D' or detalle.acu='G' or detalle.acu='E' or detalle.acu='1') "  'E nota credito
    buf2 = buf2 & " group by producto.productoequ,producto.producto order by producto.productoequ"
    'MsgBox buf2
    mytablex.Open buf2, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sw = 0
        Tmp = ""
   
        Do

            If mytablex.EOF Then Exit Do
   
            If sw = 0 Then
                buf = "" & mytablex.Fields("yproducto")
                found = formateaa(buf, 8, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombrexx("" & mytablex.Fields("yproducto"))
                found = formateaa(buf, 15, 0, 0)
                found = formateaa("", 1, 2, 0)
                sw = 1
                Tmp = "" & mytablex.Fields("yproducto")

            End If

            If Tmp <> "" & mytablex.Fields("yproducto") Then
                buf = Format(sdx, "0.00")
                found = formateaa("Totales ...", 10, 0, 0)
                found = formateaa(buf, 5, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(sdx1, "0.00")
                found = formateaa(buf, 5, 2, 1)
                sdx = 0
                sdx1 = 0
      
                buf = "" & mytablex.Fields("yproducto")
                found = formateaa(buf, 8, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombrexx("" & mytablex.Fields("yproducto"))
                found = formateaa(buf, 15, 0, 0)
                found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("yproducto")

            End If
      
            buf = "" & mytablex.Fields("zproducto")
            found = formateaa(buf, 8, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = busca_nombrexx("" & mytablex.Fields("zproducto"))
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
      
            buf = "" & mytablex.Fields("xcan")
            found = formateaa(buf, 5, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("xtot")
            found = formateaa(buf, 5, 0, 1)
            found = formateaa("", 1, 2, 0)
            sdx = sdx + Val("" & mytablex.Fields("xcan"))
            sdx1 = sdx1 + Val("" & mytablex.Fields("xtot"))
      
            ssdx = ssdx + Val("" & mytablex.Fields("xcan"))
            ssdx1 = ssdx1 + Val("" & mytablex.Fields("xtot"))
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    
    buf = Format(sdx, "0.00")
    found = formateaa("Totales ...", 10, 0, 0)
    found = formateaa(buf, 5, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 5, 2, 1)
    
    buf = Format(ssdx, "0.00")
    found = formateaa("Gran Total ...", 10, 0, 0)
    found = formateaa(buf, 5, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssdx1, "0.00")
    found = formateaa(buf, 5, 2, 1)
        
    For I = 1 To 8
        found = formateaa("", 1, 2, 0)
    Next I

    Close #ncanal
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Function busca_nombrexx(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_nombrexx = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

''''13/09/2017 kenyo Mejor Reporte Ticket
Function busca_nombrexxExtra(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select SUBSTRING(DESCRIPCIO,28,50) as descripcio from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_nombrexxExtra = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

''''13/09/2017 kenyo Mejor Reporte Ticket

Function busca_vendedor(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_vendedor = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close

End Function

'centralizacion de cuadres
Function copia_cabecera()

    Dim found     As Integer

    Dim I         As Integer

    Dim vr        As Integer

    Dim num_rec   As Long

    Dim banula    As Integer

    Dim pasos     As Double

    Dim sw        As Integer

    Dim sum5      As Double

    Dim sum6      As Double

    Dim internos  As String

    Dim otros     As String

    Dim otros1    As String

    Dim internos1 As String

    Dim mytabley  As New ADODB.Recordset

    Dim mytablex  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim mytablem  As New ADODB.Recordset

    Dim mytablev  As New ADODB.Recordset

    Dim mytableT  As New ADODB.Recordset

    Dim anulados  As String

    On Error GoTo cmd15_err

    sum1 = 0
    banula = 0
    sum2 = 0
    sum3 = 0
    sum4 = 0
    sum5 = 0
    sum6 = 0
    pasos = 0
    otros = ""
    internos = ""
    otros1 = ""
    internos1 = ""
    sw = 0
    nregistro = ""
    'Set mytablex = mydbxglo.OpenTable("factura")
    'mytablex.Index = "tfactura"
    'Set mytablez = mydbxglo.OpenTable("detalle")
    'Set mytablem = mydbxglo.OpenTable("fpagov")
    'Set mydby = OpenDatabase(RUTA, False, False, "foxpro 2.5;")
    'Set mytabley = mydby.OpenTable("cadiario")
    'Set mytables = mydby.OpenTable("dediario")
    'Set mytablen = mydby.OpenTable("fpdiario")
    
    'mytablez.Open "select * from detalle ", cn, adOpenStatic, adLockOptimistic
    
    mytablev.Open "select * from detalle where tipo='s' ", cn, adOpenStatic, adLockOptimistic
    mytableT.Open "select * from fpagov where tipo='s' ", cn, adOpenStatic, adLockOptimistic
   
    mytabley.Open "select * from cadiario where fecha='" & Format(fechai, "YYYYMMDD") & "' and caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Or mytabley.BOF Then Exit Do
            sum1 = sum1 + 1
            nregistro = Format(sum1, "00000")
            vr = DoEvents()
            Set mytablex = Nothing
            sw = 0

            If mytablex.State = 1 Then mytablex.Close
            mytablex.Open "select * from factura where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount > 0 Then

                '-------------------------------
                'GoTo pepito
                If "" & mytabley.Fields("estado") = "1" Then
                    If "" & mytabley.Fields("TIPO") = "5" Then
                        sum5 = sum5 + 1
                    Else
                        sum6 = sum6 + 1

                    End If

                    sum4 = sum4 + 1
                    banula = banula + 1
                    anulados = Format(banula, "00000")
                    otros = Format(sum6, "00000")
                    internos = Format(sum5, "00000")
                    vr = DoEvents()
                    'mytablex.Edit
                    mytablex.Fields("estado") = "1"
                    mytablex.Update
                    'ahora en detalle
                    cn.Execute ("update detalle set estado='1' where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'")
                    'ahora en fpagov
                    sw = 2
                    cn.Execute ("update fpagov set estado='1' where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'")

                End If

                'MsgBox "xxx"
pepito:
            Else
                '-------------COPIANDO CABECERA-------------------------
                
                If "" & mytabley.Fields("TIPO") = "5" Then
                    sum2 = sum2 + 1
                Else
                    sum3 = sum3 + 1

                End If

                otros1 = Format(sum3, "00000")
                internos1 = Format(sum2, "00000")
                mytablex.AddNew

                For I = 0 To mytabley.Fields.count - 1
                    mytablex.Fields(I) = mytabley.Fields(I)
                Next I

                mytablex.Update
                'found = copiar_video(mytabley)
                found = copia_detalle(mytabley, mytablev)
                found = copia_fpago(mytabley, mytableT)
                'found = copia_grafico()
                found = 1

                '--------------------------------------
            End If

            mytablex.Close
            Set mytablex = Nothing
            mytabley.MoveNext
        Loop

    End If

    '----------------------
    mytabley.Close
    Set mytabley = Nothing
    
    found = copia_ingreso(mytableT)
    mytablev.Close
    mytableT.Close
    copia_cabecera = 1
    Exit Function
cmd15_err:
    MsgBox "Aviso en copia cabecera " + Str(sw) + " " + error$, 24, "AVISO DE NO ERROR"
    Exit Function

End Function

Function copia_detalle(mytabley As ADODB.Recordset, mytablec As ADODB.Recordset)

    Dim I        As Integer

    Dim found    As String

    Dim mytablea As New ADODB.Recordset

    On Error GoTo cmd16_err

    mytablea.Open "select * from dediario where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount > 0 Then
        Do

            If mytablea.EOF Or mytablea.BOF Then Exit Do
            '-------------COPIANDO detalle-------------------------
            mytablec.AddNew

            For I = 0 To mytablea.Fields.count - 1
                mytablec.Fields(I) = mytablea.Fields(I)
            Next I

            mytablec.Update
            found = 1
            mytablea.MoveNext
        Loop

    End If

    mytablea.Close
    Exit Function
cmd16_err:
    MsgBox "Aviso en copia detalle " + error$, 24, "AVISO DE NO ERROR"
    Exit Function

End Function

Function copia_fpago(mytabley As ADODB.Recordset, mytablec As ADODB.Recordset)

    Dim I        As Integer

    Dim found    As Integer

    Dim mytablea As New ADODB.Recordset

    On Error GoTo cmd17_err

    mytablea.Open "select * from fpdiario where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount > 0 Then
        Do

            If mytablea.EOF Or mytablea.BOF Then Exit Do
            '-------------COPIANDO detalle-------------------------
            mytablec.AddNew

            For I = 0 To mytablea.Fields.count - 1
                mytablec.Fields(I) = mytablea.Fields(I)
            Next I

            mytablec.Update
            found = 1
            mytablea.MoveNext
        Loop

    End If

    mytablea.Close
    
    Exit Function
cmd17_err:
    MsgBox "Aviso en copia fpagov " + error$, 24, "AVISO DE NO ERROR"
    Exit Function

End Function

Function copia_ingreso(mytableT As ADODB.Recordset)

    Dim mytable5 As New ADODB.Recordset
 
    Dim found    As Integer

    Dim I        As Integer

    Dim vr       As Integer

    On Error GoTo cmd216_err

    mytable5.Open "select * from recibo where fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

    If mytable5.RecordCount > 0 Then
        Do

            If mytable5.EOF Or mytable5.BOF Then Exit Do
            found = copia_fpago(mytable5, mytableT)
            mytable5.MoveNext
        Loop

    End If

    mytable5.Close
    Exit Function
cmd216_err:
    MsgBox "Aviso en copia ingreso", 24, "AVISO DE NO ERROR"
    Exit Function

End Function

Sub borra_cadiario()
    cn.Execute ("delete from cadiario   where fecha='" & Format(fechai, "YYYYMMDD") & "' and caja='" & "" & caja & "'")
    cn.Execute ("delete from dediario   where fecha='" & Format(fechai, "YYYYMMDD") & "' and caja='" & "" & caja & "'")
    cn.Execute ("delete from fpdiario   where fecha='" & Format(fechai, "YYYYMMDD") & "' and caja='" & "" & caja & "'")

End Sub

Function parametro_mesas() As String

    Dim mytablexx As New ADODB.Recordset

    If mytablexx.State = 1 Then mytablexx.Close
    mytablexx.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablexx.RecordCount > 0 Then
        parametro_mesas = "" & mytablexx.Fields("mesabierta")

    End If

    mytablexx.Close

End Function

Function mesa_abiertas() As Double

    Dim mytablexx As New ADODB.Recordset

    If mytablexx.State = 1 Then mytablexx.Close
    mytablexx.Open "select * from dcomanda ", cn, adOpenStatic, adLockOptimistic

    If mytablexx.RecordCount > 0 Then
        mesa_abiertas = mytablexx.RecordCount

    End If

    mytablexx.Close

End Function

Sub percepcion_quiebre()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim buf2     As String

    Dim I        As Integer

    Dim sdx      As Double

    Dim sdx1     As Double

    sdx = 0
    sdx1 = 0
    I = 0
    buf2 = "select clasesunat ,count(percepcion) as xcan,sum(percepcion) as xper from " & dbca & " where  "
    buf2 = buf2 & "   fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf2 = buf2 & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf2 = buf2 & " and len(clasesunat)>0 "

    If local1 <> "%" Then
        buf2 = buf2 & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf2 = buf2 & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf2 = buf2 & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf2 = buf2 & " and turno like '" & turno & "%'"

    End If

    buf2 = buf2 & " and estado='2'"
    buf2 = buf2 & " and ( acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='1') "  'E nota credito
    buf2 = buf2 & " group by clasesunat order by clasesunat"
    mytablex.Open buf2, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        I = 1
        'buf = "" & mytablex.Fields("clasesunat")
        'found = formateaa(buf, 2, 0, 0)
        'found = formateaa("", 1, 0, 0)
        buf = "" & busca_clase("" & mytablex.Fields("clasesunat"))
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
      
        buf = "" '& mytablex.Fields("xcan")
        found = formateaa(buf, 5, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(Val("" & mytablex.Fields("xper")), "0.00")
        found = formateaa(buf, 5, 0, 1)
        found = formateaa("", 1, 2, 0)
        sdx = sdx + Val("" & mytablex.Fields("xcan"))
        sdx1 = sdx1 + Val("" & mytablex.Fields("xper"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    Exit Sub

    If I = 1 Then
        buf = "" & sdx
        found = formateaa(buf, 5, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & sdx1
        found = formateaa(buf, 5, 0, 1)
        found = formateaa("", 1, 2, 0)

    End If

End Sub

Sub nuevo_unidades()

    Dim buf2     As String

    Dim buf      As String

    Dim found    As Integer

    Dim Tmp      As String

    Dim sw       As Integer

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim sdx3     As Double

    Dim sdx4     As Double

    Dim sdx5     As Double

    Dim signos   As Double

    Dim mytablex As New ADODB.Recordset

    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0
    sdx5 = 0
    'MsgBox dbde

    buf2 = "select producto,estado,sum(cantidad*factor) as xcant,sum(total) as xtot from " & dbde & " where "
    buf2 = buf2 & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf2 = buf2 & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf2 = buf2 & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf2 = buf2 & " and usuario like '" & cajero & "%'"

    End If

    If caja <> "%" Then
        buf2 = buf2 & " and caja like '" & caja & "%'"

    End If

    If turno <> "%" Then
        buf2 = buf2 & " and turno like '" & turno & "%'"

    End If

    If flag_proveedor = "S" Then
        buf = buf & " and len(proveedorp)>0 "

    End If

    buf2 = buf2 & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='1') "  'E nota credito
    'buf2 = buf2 & " group by producto,estado order by str(producto)"
    buf2 = buf2 & " group by producto,estado order by producto"

    mytablex.Open buf2, cn, adOpenStatic, adLockOptimistic
    sw = 0
   
    Do

        If mytablex.EOF Then Exit Do
        If sw = 0 Then
            Tmp = "" & mytablex.Fields("producto")
            sw = 1
            buf = "" & mytablex.Fields("producto")
   
            ''''18/09/2017 kenyo Mejor Reporte Ticket
            ' found = formateaa(buf, 5, 0, 0)
            found = formateaa("", 1, 2, 0)
            found = formateaa("* ", 2, 0, 0)
            buf = Mid$(name_productos("" & mytablex.Fields("producto")), 1, 26)
            found = formateaa(buf, 26, 0, 0)

            If Len(name_productos(mytablex.Fields("producto"))) > 26 Then
                found = formateaa("", 1, 2, 0)
                buf = Mid$(name_productosExtra("" & mytablex.Fields("producto")), 1, 26)
                found = formateaa(buf, 26, 0, 0)

            End If
 
            found = formateaa("", 1, 2, 0)

        End If
 
        If Tmp <> "" & mytablex.Fields("producto") Then
            found = formateaa("", 1, 2, 0)
            Tmp = "" & mytablex.Fields("producto")
            buf = "" & mytablex.Fields("producto")
            ' found = formateaa(buf, 5, 0, 0)
            found = formateaa("* ", 2, 0, 0)
            buf = Mid$(name_productos("" & mytablex.Fields("producto")), 1, 26)
            found = formateaa(buf, 26, 0, 0)

            If Len(name_productos(mytablex.Fields("producto"))) > 26 Then
                found = formateaa("", 1, 2, 0)
                buf = Mid$(name_productosExtra("" & mytablex.Fields("producto")), 1, 26)
                found = formateaa(buf, 26, 0, 0)
      
            End If

            found = formateaa("", 1, 2, 0)

        End If
   
        ''''18/09/2017 kenyo Mejor Reporte Ticket
  
        If Trim("" & mytablex.Fields("estado")) = "2" Then
       
            found = formateaa("+Vendidos:", 10, 0, 0)
            sdx = sdx + Val("" & mytablex.Fields("xcant"))
            sdx1 = sdx1 + Val(Format(Val("" & mytablex.Fields("xtot")), "0.00"))
   
        End If
   
        If Trim("" & mytablex.Fields("estado")) = "1" Then
            found = formateaa("-Anuladas:", 10, 0, 0)
            sdx2 = sdx2 + Val("" & mytablex.Fields("xcant"))
            sdx3 = sdx3 + Val(Format(Val("" & mytablex.Fields("xtot")), "0.00"))

        End If

        If Trim("" & mytablex.Fields("estado")) <> "1" And Trim("" & mytablex.Fields("estado")) <> "2" Then
            found = formateaa("*Otros:", 10, 0, 0)
            sdx4 = sdx4 + Val("" & mytablex.Fields("xcant"))
            sdx5 = sdx5 + Val(Format(Val("" & mytablex.Fields("xtot")), "0.00"))

        End If

        buf = "" & mytablex.Fields("xcant")
        found = formateaa(buf, 7, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(Val("" & mytablex.Fields("xtot")), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        mytablex.MoveNext
    Loop
    mytablex.Close
    found = formateaa("", 1, 2, 0)
    found = formateaa("<<<<TOTALES<<<", 20, 2, 0)
    found = formateaa("", 1, 2, 0)
    found = formateaa("+TotVentas:", 10, 0, 0)
    buf = "" & sdx
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "" & sdx1
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    found = formateaa("-TotAnulad:", 10, 0, 0)
    buf = "" & sdx2
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "" & sdx3
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    found = formateaa("*TotOtros:", 10, 0, 0)
    buf = "" & sdx4
    found = formateaa(buf, 7, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = "" & sdx5
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Function name_productos(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        name_productos = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

''''18/09/2017 kenyo Mejor Reporte Ticket
Function name_productosExtra(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select SUBSTRING(DESCRIPCIO,27,50) as descripcio from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        name_productosExtra = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

''''18/09/2017 kenyo Mejor Reporte Ticket

Sub envio_correos(perfil As String)

    Dim txtserver      As String

    Dim txtusername    As String

    Dim txtpassword    As String

    Dim txtport        As String

    Dim txtto          As String

    Dim chkssl         As String

    Dim txtfromname    As String

    Dim txtfromemail   As String

    Dim txtattach      As String

    Dim txtattach2     As String

    Dim txtsubject     As String

    Dim txtmsg         As String

    Dim retval         As String

    Dim txthtml        As String

    Dim txtselecciona  As String

    'Dim txtselecciona As String
    Dim mytablex       As New ADODB.Recordset

    Dim buf            As String

    'Dim txtselecciona As String

    Dim opcorreocierre As String

    Dim opsaldocierre  As String

    On Error GoTo cmd0905677_err

    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
    opcorreocierre = Trim("" & mytablex.Fields("correocierre"))
    opsaldocierre = Trim("" & mytablex.Fields("saldocierre"))

    If mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("correocierre") <> "S" And "" & mytablex.Fields("saldocierre") <> "S" Then
            mytablex.Close
            Exit Sub

        End If

    End If

    mytablex.Close

    buf = extra_loquesea1(perfil)

    If Trim(buf) = 0 Then Exit Sub
    mytablex.Open "select * from correos where cosms='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        txtserver = Trim("" & mytablex.Fields("txtserver"))
        txtusername = Trim("" & mytablex.Fields("txtusername"))
        txtpassword = Trim("" & mytablex.Fields("txtpassword"))
        txtfromname = Trim("" & mytablex.Fields("txtfromname"))
        txtfromemail = Trim("" & mytablex.Fields("txtfromemail"))
        txtport = Trim("" & mytablex.Fields("txtport"))
        txtselecciona = Trim("" & mytablex.Fields("txtselecciona"))

        chkssl = Trim("" & mytablex.Fields("chkssl"))

        txtto = Trim("" & mytablex.Fields("txtfromemail"))
        'txtattach =  'Trim("" & mytablex.Fields("txtattach"))

        If opcorreocierre = "S" Then
            txtattach = FileName

        End If

        If opsaldocierre = "S" Then
            reporte_saldo
            txtattach2 = App.path & "\001D\06\temporal\saldo.txt"

        End If

        ''07/07/2017 KENYO cierre de caja con asunto
        'txtsubject = Trim("" & mytablex.Fields("txtsubject"))
        txtsubject = "Reporte de Ventas: " + fechai + " - " + fechaf
        ''07/07/2017 KENYO cierre de caja con asunto

        txtmsg = Trim("" & mytablex.Fields("txtmsg"))
        txtmsg = txtmsg & Chr$(10) & Chr$(13) & ""
        txtmsg = txtmsg & Format(Now, "dd/mm/yyyy") + " " + Format(Now, "hh:mm:ss")

        If Len(Trim("" & mytablex.Fields("txtfromemail"))) > 0 Then
   
            txtto = Trim("" & mytablex.Fields("txtfromemail"))
   
            If opcorreocierre = "S" And opsaldocierre = "S" Then
                retval = SendMailAdjunto(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml, Trim$(txtattach2))

            End If
   
            If opsaldocierre = "N" Or opsaldocierre = "" Or opsaldocierre = " " Then
                retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml)

            End If
   
            If opcorreocierre = "N" Or opcorreocierre = "" Or opcorreocierre = " " Then
                retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach2), True, txtselecciona, txthtml)

            End If
   
        End If

        MsgBox "Correo Enviado ", 48, "Aviso"

    End If

    mytablex.Close

    Exit Sub
cmd0905677_err:
    MsgBox "No se Pudo enviar Correo... " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function borra_nombre(buf As String)

    On Error GoTo cmd457_err

    Kill buf
    borra_nombre = 1
    Exit Function
cmd457_err:
    Exit Function

End Function

Sub cerrar_archivo()

    On Error GoTo cmd561_err

    Close
    Exit Sub
cmd561_err:
    MsgBox "Aviso en cerrar_archivo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub reporte_saldo()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim valor    As String

    valor = "SALDO"

    found = sql_producto(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & valor & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_saldo
    cuerpo_programa_saldo mytablex
    Command1.Visible = False
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    VISTA = "OK"
    found = valida_wordpad(FileName)

End Sub

Function sql_producto(mytablex As ADODB.Recordset)

    Dim buf As String

    buf = "select * from producto order by familia,Subfamilia,descripcio"

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    sql_producto = 1

End Function

Sub cabecera_saldo()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    'found = formateaa("ALMACEN : " & bodega, 25, 2, 0)
    found = formateaa("FECHA INICIO :" & fechai, 40, 2, 0)
    found = formateaa("FECHA FINAL  :" & fechaf, 40, 2, 0)
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Producto ", 11, 0, 0)
    found = formateaa("Descripcio ", 41, 0, 0)
    found = formateaa("Unid", 7, 0, 0)
    found = formateaa("Fx", 5, 0, 0)
   
    found = formateaa("Cantid ", 11, 0, 1)
    found = formateaa("", 1, 2, 0)
  
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

Function busca_familia(buf As String) As String

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from familia where familia='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_familia = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

Sub nlineas()

    '    contlin = contlin + 1
    '    If contlin > Val(nrolineas) Then
    '
    '       cabecera_saldo
    '      End If
  
End Sub

Sub cuerpo_programa_saldo(mytablex As ADODB.Recordset)

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim sw        As Integer

    Dim xnroitem  As Double

    Dim temp      As String

    Dim buf       As String

    Dim sw1       As Integer

    Dim temp1     As String

    Dim buf1      As String

    Dim bufx      As String

    Dim saldoini  As Double

    Dim saldoindx As Double

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim sdx2      As Double

    Dim found     As Integer

    Dim xprecio   As Double

    Dim vr

    Dim buf2 As String

    saldoindx = 0
    sdx1 = 0
    sdx2 = 0
    suma1 = 0
    suma2 = 0
    sw1 = 0
    saldoini = 0
    xnroitem = 0
    Command1.Visible = True
    Do

        If mytablex.EOF Then Exit Do

        vr = DoEvents()

        If Command1.Visible = False Then Exit Do

        xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00"))

        'verificamos que tipo de costeo
        If mytablez.State = 1 Then mytablez.Close

        xprecio = Val(Format(Val("" & mytablex.Fields("costop")), "0.00"))

        mytablez.Open "Select * from precios where local='" & extra_loquesea(local1) & "' and  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            xprecio = "" & mytablez.Fields("pventa1")

        End If

        mytablez.Close

        xprecio = xprecio

        '------------- verificamos la condicion
        'If Combo1 <> "TODOS" Then
        saldoini = 0
        buf2 = "Select * from almacen where local='01' and  producto='" & "" & mytablex.Fields("producto") & "' and bodega='01'"

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open buf2, cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("saldo"))
            '   If Combo1 = "SALDO>0" Then
            '      If saldoini > 0 Then GoTo sigueme
            '   End If
            '   If Combo1 = "SALDO>=0" Then
            '      If saldoini >= 0 Then GoTo sigueme
            '   End If
            '   If Combo1 = "SALDO<0" Then
            '      If saldoini < 0 Then GoTo sigueme
            '   End If
            '   If Combo1 = "SALDO<=0" Then
            '      If saldoini <= 0 Then GoTo sigueme
            '   End If
            '   If Combo1 = "SALDO=0" Then
            '      If saldoini = 0 Then GoTo sigueme
            '   End If
   
            ' If Combo1 = "TODOS" Then
            'MsgBox "Hola"
            'End
            GoTo sigueme
            'End If
            GoTo sigueme1
            Else: GoTo sigueme1

        End If

        'End If
sigueme:
        'mytablez.Close

        buf = "" & mytablex.Fields("familia")

        If sw = 0 Then
            sw = 1
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("*", 1, 2, 0)
            temp = "" & mytablex.Fields("familia")
            nlineas

        End If

        If "" & mytablex.Fields("familia") <> temp Then
            temp = "" & mytablex.Fields("familia")
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            buf = busca_familia("" & mytablex.Fields("familia"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If sw1 = 0 Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")
            sw1 = 1

        End If

        If temp1 <> "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia") Then
            found = formateaa("" & mytablex.Fields("familia"), 6, 0, 0)
            found = formateaa("-", 1, 0, 0)
            found = formateaa("" & mytablex.Fields("subfamilia"), 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            temp1 = "" & mytablex.Fields("familia") & "" & mytablex.Fields("subfamilia")

        End If

        found = formateaa("" & mytablex.Fields("producto"), 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("descripcio"), 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("unidad"), 6, 0, 0)
        found = formateaa("x", 1, 0, 0)
        found = formateaa("" & mytablex.Fields("factor"), 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        'SALDO ALMACEN
        'saldoini = 0
        'mytablez.Seek "=", local1, "" & mytablex.Fields("producto"), extra_loquesea(bodega)
        'If Not mytablez.NoMatch Then
        '   saldoini = Val("" & mytablez.Fields("saldo"))
        'End If
        saldoindx = saldoini
        bufx = "" & saldoini

        If Val(bufx) = 0 Then
            bufx = ""

        End If

        bufx = "" & saldoindx 'calcula_saldo(saldoini, Val("" & mytablex.Fields("factor")))
        found = formateaa(bufx, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        'If gcanti <> "S" Then
        '   buf = Format(xprecio, "0.00")
        '   found = formateaa(buf, 10, 0, 1)
        '   found = formateaa("", 1, 0, 0)
        '   sdx = xprecio * saldoindx
        '   buf = Format(sdx, "0.00")
        '   suma1 = suma1 + sdx
        '   suma2 = suma2 + saldoini
        '   found = formateaa(buf, 10, 0, 1)
        'End If
        found = formateaa("", 1, 2, 0)
        nlineas
        xnroitem = xnroitem + 1
        '---------------------------------------
sigueme1:
seguy5:
        'mytablez.Close

        mytablex.MoveNext
    Loop
    buf = "" & xnroitem
    found = formateaa("Nro Productos " + buf, 64, 2, 0)
    nlineas
    bufx = Format(suma2, "0.00")
    found = formateaa("", 64, 0, 0)
    found = formateaa(bufx, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    'If gcanti <> "S" Then
    '   bufx = Format(suma1, "0.00")
    '   found = formateaa(bufx, 10, 0, 1)
    '   found = formateaa("", 1, 2, 0)
    '   Else
    '   found = formateaa("", 1, 2, 0)
    'End If

End Sub

Private Sub turnoc_Click()
    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja..Crear combo cajac en tcuadrc1
    'turno = extra_loquesea(turnoc)
    '''01/07/2017 Kenyo Agregamos caja en cierre de copia de caja..Crear combo cajac en tcuadrc1

End Sub
