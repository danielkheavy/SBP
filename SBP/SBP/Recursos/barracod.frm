VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form barracod 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Codigo Barras"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6465
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Etiquetas"
      Height          =   6135
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox aaxx2 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   40
         Text            =   "65"
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox xxfila 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   35
         Text            =   "0110"
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox xxcolumna 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   34
         Text            =   "0015"
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox aaxx1 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   32
         Text            =   "2"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   31
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   5640
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   4440
         TabIndex        =   29
         Top             =   5640
         Width           =   855
      End
      Begin VB.TextBox descripcio 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         MaxLength       =   60
         TabIndex        =   28
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox marca 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   13
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox columnas 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "3"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox cantidad 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "1"
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox copias 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "1"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox barras 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "1"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox columna 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "0036"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox fila 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "0120"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox altura 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "030"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox separa 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "2"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox ancho 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "2"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox tipo 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "E"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox rotacion 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "2"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "AAXX"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   4320
         Width           =   3015
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Barras"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Letras"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3960
         Width           =   5175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fila (Cordenada Y)"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   4560
         Width           =   3015
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Columna (Cordenada X)"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   4800
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "AAXX"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marca"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Medida ticket"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Columnas x Etiquetas"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero de Etiquetas"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero Copias Iguales"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Codigo Barras"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Columna (Cordenada X)"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fila (Cordenada Y)"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Altura de las Barras"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Separacion de las Barras"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ancho de las Barras"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo Cod Barras. (A-Z a-z)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rotacion (1.Normal 2.90 3.180m 4.270 )"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label estado 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NINGUNO"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   5760
         Width           =   825
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "barracod.frx":0000
      Height          =   6975
      Left            =   120
      OleObjectBlob   =   "barracod.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   6150
   End
   Begin VB.Menu nuw23 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu mofi23 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu bo723 
      Caption         =   "&Borra"
   End
   Begin VB.Menu fdk2323 
      Caption         =   "&Print"
   End
   Begin VB.Menu dlo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "barracod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aaxx                         As Integer

'Private Const NombreImpresora_sp As String = "Generic / Text Only"
Private Const NombreImpresora_us As String = "Genérico / sólo texto"

Private Const NombreImpresora_sp As String = "Argox X-1000v PPLA"

'Private Const NombreImpresora_sp As String = "Generic / Text Only"

Private Function colocar(Texto As String, X As Integer, Y As Integer)
    'x es la fila
    'y es columna
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print Texto
  
End Function

Private Sub Command2_Click()

    Dim found As Integer

    On Error GoTo cmd5_err

    found = verifica()

    If found = 0 Then Exit Sub

    If estado = "NUEVO" Then
        Data1.Recordset.AddNew

    End If

    If estado = "MODIFICA" Then
        Data1.Recordset.Edit

    End If

    Data1.Recordset.Fields("aaxx") = aaxx1
    Data1.Recordset.Fields("rotacion") = rotacion
    Data1.Recordset.Fields("tipo") = tipo
    Data1.Recordset.Fields("ancho") = ancho
    Data1.Recordset.Fields("separa") = separa
    Data1.Recordset.Fields("altura") = altura
    Data1.Recordset.Fields("fila") = fila
    Data1.Recordset.Fields("columna") = columna
    Data1.Recordset.Fields("descripcio") = descripcio
    'Data1.Recordset.Fields("barras") = barras
    Data1.Recordset.Fields("copias") = copias
    Data1.Recordset.Fields("marca") = marca
    Data1.Recordset.Fields("cantidad") = cantidad
    Data1.Recordset.Fields("columnas") = columnas

    Data1.Recordset.Update
    Command3_Click
    Exit Sub
cmd5_err:
    MsgBox "Error en grabacion ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command3_Click()
    Frame1.Visible = False

End Sub

Private Sub dlo232_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    barracod.Hide
    Unload barracod

End Sub

Sub SQL()

    On Error GoTo cmd37_err

    Dim buf As String

    buf = "select * from etiqueta "
    Data1.Connect = "foxpro 2.5;"
    Data1.DatabaseName = globaldir
    Data1.RecordSource = buf
    Data1.refresh
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 1000
    dbGrid1.columns(2).Width = 1000
    dbGrid1.columns(3).Width = 1000
    dbGrid1.columns(4).Width = 1000
    dbGrid1.columns(5).Width = 1000
    dbGrid1.columns(6).Width = 1000
    dbGrid1.columns(7).Width = 1000
   
    Exit Sub
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub inicializa()
    aaxx1 = "50"
    rotacion = ""
    tipo = ""
    ancho = ""
    separa = ""
    altura = ""
    fila = ""
    columna = ""
    Barras = ""
    copias = ""
    columnas = ""
    descripcio = ""
    marca = ""
    cantidad = ""
    rotacion = "1"
    tipo = "E"
    ancho = "2"
    separa = "2"
    altura = "030"
    fila = "0010"
    columna = "0010"

End Sub

Private Sub fdk2323_Click()

    On Error GoTo cmd342_err

    If Frame1.Visible = True Then Exit Sub
    Command1.Enabled = True
    Command2.Enabled = False
    aaxx1 = "" & Data1.Recordset.Fields("aaxx")
    rotacion = "" & Data1.Recordset.Fields("rotacion")
    tipo = "" & Data1.Recordset.Fields("tipo")
    ancho = "" & Data1.Recordset.Fields("ancho")
    separa = "" & Data1.Recordset.Fields("separa")
    altura = "" & Data1.Recordset.Fields("altura")
    fila = "" & Data1.Recordset.Fields("fila")
    columna = "" & Data1.Recordset.Fields("columna")
    descripcio = "" & Data1.Recordset.Fields("descripcio")
    'barras = "" & Data1.Recordset.Fields("barras")
    copias = "" & Data1.Recordset.Fields("copias")
    marca = "" & Data1.Recordset.Fields("marca")
    columnas = "" & Data1.Recordset.Fields("columnas")
    cantidad = "" & Data1.Recordset.Fields("cantidad")
    Frame1.Visible = True
    estado = "PRINT"
    Exit Sub
cmd342_err:
    MsgBox "Seleccione una descripcion", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Load()
    'Combo1.Clear
    'Combo1.AddItem "3.3x1.9"
    'Combo1.AddItem "3.8x2.3"
    'Combo1.ListIndex = 0
    'If Combo1 = "3.3x1.9" Then
    '   valor33x19
    'End If
    'If Combo1 = "3.8x2.3" Then
    '   valor38x23
    'End If
    SQL
    ve_permisos

End Sub

Sub ve_permisos()

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("vendedor")
    mytablex.Index = "codigo"
    mytablex.Seek "=", gusuario

    If Not mytablex.NoMatch Then
        If Mid$("" & mytablex.Fields("permiso"), 1, 1) = "R" Then
            nuw23.Enabled = False
            mofi23.Enabled = False
            bo723.Enabled = False

        End If

    End If

    mytablex.Close

End Sub

Sub valor38x23()
    rotacion = "2"
    tipo = "E"
    ancho = "2"
    separa = "2"
    altura = "030"
    fila = "0120"
    columna = "0036"
    aaxx = 100

End Sub

Sub valor33x19()
    rotacion = "1"
    tipo = "E"
    ancho = "2"
    separa = "2"
    altura = "030"
    fila = "0010"
    columna = "0010"
    aaxx = 140

End Sub

Private Sub mofi23_Click()

    On Error GoTo cmd34_err

    If Frame1.Visible = True Then Exit Sub
    Command2.Enabled = True
    Command1.Enabled = True
    aaxx1 = "" & Data1.Recordset.Fields("aaxx")
    rotacion = "" & Data1.Recordset.Fields("rotacion")
    tipo = "" & Data1.Recordset.Fields("tipo")
    ancho = "" & Data1.Recordset.Fields("ancho")
    separa = "" & Data1.Recordset.Fields("separa")
    altura = "" & Data1.Recordset.Fields("altura")
    fila = "" & Data1.Recordset.Fields("fila")
    columna = "" & Data1.Recordset.Fields("columna")
    descripcio = "" & Data1.Recordset.Fields("descripcio")
    'barras = "" & Data1.Recordset.Fields("barras")
    copias = "" & Data1.Recordset.Fields("copias")
    marca = "" & Data1.Recordset.Fields("marca")
    columnas = "" & Data1.Recordset.Fields("columnas")
    cantidad = "" & Data1.Recordset.Fields("cantidad")
    Frame1.Visible = True
    estado = "MODIFICA"
    Exit Sub
cmd34_err:
    MsgBox "Seleccione una descripcion", 48, "Aviso"
    Exit Sub

End Sub

Private Sub nuw23_Click()

    If Frame1.Visible = True Then Exit Sub
    Command2.Enabled = True
    Command1.Enabled = False
    Frame1.Visible = True
    inicializa
    estado = "NUEVO"

End Sub

Function verifica()

    If Not IsNumeric(rotacion) Then
        rotacion.SetFocus
        Exit Function

    End If

    If Len(tipo) = 0 Then
        tipo.SetFocus
        Exit Function

    End If

    If Not IsNumeric(ancho) Then
        ancho.SetFocus
        Exit Function

    End If

    If Not IsNumeric(separa) Then
        separa.SetFocus
        Exit Function

    End If

    If Not IsNumeric(altura) Then
        altura.SetFocus
        Exit Function

    End If

    If Not IsNumeric(fila) Then
        fila.SetFocus
        Exit Function

    End If

    If Not IsNumeric(columna) Then
        columna.SetFocus
        Exit Function

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    If Not IsNumeric(copias) Then
        copias.SetFocus
        Exit Function

    End If

    If Not IsNumeric(cantidad) Then
        cantidad.SetFocus
        Exit Function

    End If

    If Not IsNumeric(columnas) Then
        columnas.SetFocus
        Exit Function

    End If

    If Len(marca) = 0 Then
        marca.SetFocus
        Exit Function

    End If

    If Len(Barras) = 0 Then
        Barras.SetFocus
        Exit Function

    End If

    verifica = 1

End Function
