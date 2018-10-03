VERSION 5.00
Begin VB.Form Cargastk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste Rapida Stock"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10635
      TabIndex        =   8
      Top             =   0
      Width           =   10695
      Begin VB.CommandButton cmdAddEntry 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Picture         =   "Cargastk.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
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
         Picture         =   "Cargastk.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Grabar registro"
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
         Picture         =   "Cargastk.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label bandera 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8760
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ztipo 
         Height          =   375
         Left            =   9480
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label zserie 
         Height          =   375
         Left            =   10200
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label znumero 
         Height          =   375
         Left            =   11040
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.TextBox saldo 
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
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   0
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox saldoa 
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
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1920
      Width           =   2655
   End
   Begin VB.ComboBox bodega 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.ComboBox local1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Unds"
      Height          =   495
      Left            =   4770
      TabIndex        =   21
      Top             =   2475
      Width           =   735
   End
   Begin VB.Label aksw 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label producto 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   19
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label descripcio 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   18
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bodega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Menu dlo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Cargastk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bodega_Click()
    saldo_actual

End Sub

Private Sub cmdExit_Click()
    dlo232_Click

End Sub

Private Sub cmdSave_Click()

    Dim found As Integer

    If Not IsNumeric(saldo) Then
        saldo = ""
        saldo.SetFocus
        Exit Sub

    End If

    If Val(saldoa) = Val(saldo) Then
        Exit Sub

    End If

    If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then Exit Sub
    found = grabar()

    If found = 0 Then
        MsgBox "No se pudo grabar", 48, "Aviso"

    End If

    dlo232_Click

End Sub

Function grabar()

    Dim found    As Integer

    Dim mytablex As Table

    Dim acu      As String

    Dim sw       As Integer

    Dim saldoa   As Double

    Dim xingreso As Double

    Dim xegreso  As Double

    sw = 0
    saldoa = 0
    xingreso = 0
    xegreso = 0
    Set mytablex = mydbxglo.OpenTable("almacen")
    mytablex.Index = "almacen"
    mytablex.Seek "=", extra_loquesea(local1), producto, extra_loquesea(bodega)

    If Not mytablex.NoMatch Then
        mytablex.Edit
       
        If Val("" & mytablex.Fields("saldo")) > Val(saldo) Then
            xegreso = Val("" & mytablex.Fields("saldo")) - Val(saldo)
            acu = "T"  'salida

        End If

        If Val("" & mytablex.Fields("saldo")) < Val(saldo) Then
            xingreso = -Val("" & mytablex.Fields("saldo")) + Val(saldo)
            acu = "S"

        End If

        mytablex.Fields("saldo") = Val("" & saldo)
        mytablex.Update
        sw = 1

    End If

    If mytablex.NoMatch Then
        mytablex.AddNew
        mytablex.Fields("producto") = "" & producto
        mytablex.Fields("local") = extra_loquesea(local1)
        mytablex.Fields("bodega") = extra_loquesea(bodega)
        mytablex.Fields("saldo") = Val("" & saldo)
        xingreso = Val(saldo)
        mytablex.Update
        acu = "S"
        sw = 1

    End If

    mytablex.Close
    found = graba_kardex(acu, xingreso, xegreso)

    If found = 1 Then
        sw = 1

    End If

    grabar = 1

End Function

Function graba_kardex(acu As String, xingreso As Double, xegreso As Double)

    On Error GoTo cmd781_err

    Dim mytablez As Table

    Set mytablez = mydbxglo.OpenTable("detalle")
    mytablez.AddNew
    mytablez.Fields("estado") = "2"
    mytablez.Fields("acu") = acu

    If acu = "S" Then
        mytablez.Fields("tipo") = "23"
        mytablez.Fields("cantidad") = xingreso

    End If

    If acu = "T" Then
        mytablez.Fields("cantidad") = xegreso
        mytablez.Fields("tipo") = "S"

    End If

    mytablez.Fields("serie") = ""
    mytablez.Fields("numero") = "AJUSTE"
    mytablez.Fields("tipoclie") = "I"
    mytablez.Fields("codigo") = "" & gusuario
    mytablez.Fields("acu1") = ""
    mytablez.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablez.Fields("moneda") = "S"
    mytablez.Fields("producto") = "" & producto
    mytablez.Fields("descripcio") = "" & descripcio
    mytablez.Fields("unidad") = "UND"
    mytablez.Fields("factor") = 1

    mytablez.Fields("precio") = 0
    mytablez.Fields("igv") = 19
    mytablez.Fields("neto") = 0
    mytablez.Fields("descuento") = 0
    mytablez.Fields("subtotal") = 0
    mytablez.Fields("impuesto") = 0
    mytablez.Fields("total") = 0
    mytablez.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablez.Fields("hora") = Format(Now, "hh:mm:ss")
    mytablez.Fields("vendedor") = ""
    mytablez.Fields("bodega") = extra_loquesea(bodega)
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = 0
    mytablez.Fields("flage") = ""
    mytablez.Fields("linea") = ""
    mytablez.Fields("t1") = 0
    mytablez.Fields("t2") = 0
    mytablez.Fields("t3") = 0
    mytablez.Fields("t4") = 0
    mytablez.Fields("t5") = 0
    mytablez.Fields("t6") = 0
    mytablez.Fields("t7") = 0
    mytablez.Fields("t8") = 0
    mytablez.Fields("t9") = 0
    mytablez.Fields("t10") = 0
    mytablez.Fields("t11") = 0
    mytablez.Fields("t12") = 0
    mytablez.Fields("t13") = 0
    mytablez.Fields("t14") = 0
    mytablez.Fields("t15") = 0
    mytablez.Fields("t16") = 0
    mytablez.Fields("l1") = ""
    mytablez.Fields("l2") = ""
    mytablez.Fields("l3") = ""
    mytablez.Fields("l4") = ""
    mytablez.Fields("local") = ""
    mytablez.Fields("proveedorp") = ""
    mytablez.Fields("observa1") = ""
    mytablez.Fields("observa2") = ""
    mytablez.Fields("observa3") = ""
    mytablez.Fields("observa4") = ""
    mytablez.Fields("zona") = ""
    mytablez.Fields("isc") = 0
    mytablez.Fields("tax") = 0
    mytablez.Fields("vtaneta") = 0
    mytablez.Fields("tcosto") = 0
    mytablez.Fields("ganancia") = 0
    mytablez.Fields("comision") = 0
    mytablez.Fields("usuario") = gusuario
    mytablez.Fields("caja") = ""
    mytablez.Fields("turno") = ""
    mytablez.Fields("servicio") = ""
    mytablez.Fields("comanda") = ""
    mytablez.Fields("mesa") = ""
    mytablez.Fields("salon") = ""
    mytablez.Fields("mesero") = ""
    mytablez.Fields("local") = extra_loquesea(local1)
    mytablez.Update
    mytablez.Close
    graba_kardex = 1
    Exit Function
cmd781_err:
    MsgBox "Error " + error$, 48, "Aviso"
    Exit Function

End Function

Private Sub dlo232_Click()
    Cargastk.Hide
    Unload Cargastk

End Sub

Private Sub Form_Activate()

    If aksw = "" Then
        carga_inicial
        saldo_actual

    End If

    aksw = "1"

End Sub

Sub carga_inicial()

    Dim mytablex As Table

    local1.Clear
    Set mytablex = mydbxglo.OpenTable("tlocal")
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    bodega.Clear
    Set mytablex = mydbxglo.OpenTable("bodega")
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    bodega.ListIndex = 0

End Sub

Sub saldo_actual()

    Dim mytablex As Table

    saldoa = ""
    Set mytablex = mydbxglo.OpenTable("almacen")
    mytablex.Index = "almacen"
    mytablex.Seek "=", extra_loquesea(local1), producto, extra_loquesea(bodega)

    If Not mytablex.NoMatch Then
        saldoa = "" & mytablex.Fields("saldo")

    End If

    mytablex.Close

End Sub

Private Sub local1_Click()
    saldo_actual

End Sub

