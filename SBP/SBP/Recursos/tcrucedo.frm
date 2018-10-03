VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcrucedo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos para cruzar"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   30
      TabIndex        =   30
      Top             =   15
      Visible         =   0   'False
      Width           =   12615
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   12091
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox local1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   28
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAddEntry 
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
      Left            =   3480
      Picture         =   "tcrucedo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Nuevo registro"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox codigo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      MaxLength       =   11
      TabIndex        =   25
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox tipoclie 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   3480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tcrucedo.frx":1212
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Borrar registro"
      Top             =   2520
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
      Left            =   3480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tcrucedo.frx":2424
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Grabar registro"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox tipo 
      Height          =   375
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox serie1 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox numero1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   13
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox serie2 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox numero2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox serie3 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox numero3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox serie4 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox numero4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox serie5 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox numero5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox serie6 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox numero6 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox serie7 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox numero7 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipoclie"
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label xarchivo1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label xarchivo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label acu 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Menu logt34342 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcrucedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        logt34342_Click
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub cmdAddEntry_Click()

    tfactura.tipo1 = ""
    tfactura.serie1 = ""
    tfactura.serie2 = ""
    tfactura.serie3 = ""
    tfactura.serie4 = ""
    tfactura.serie5 = ""
    tfactura.serie6 = ""
    tfactura.serie7 = ""
    tfactura.numero1 = ""
    tfactura.numero2 = ""
    tfactura.numero3 = ""
    tfactura.numero4 = ""
    tfactura.numero5 = ""
    tfactura.numero6 = ""
    tfactura.numero7 = ""

    tipo = ""
    serie1 = ""
    serie2 = ""
    serie3 = ""
    serie4 = ""
    serie5 = ""
    serie6 = ""
    serie7 = ""
    numero1 = ""
    numero2 = ""
    numero3 = ""
    numero4 = ""
    numero5 = ""
    numero6 = ""
    numero7 = ""

End Sub

Private Sub cmdDelete_Click()
    logt34342_Click

End Sub

Private Sub cmdSave_Click()

    Dim I As Integer

    tfactura.tipo1 = tipo
    tfactura.serie1 = serie1
    tfactura.serie2 = serie2
    tfactura.serie3 = serie3
    tfactura.serie4 = serie4
    tfactura.serie5 = serie5
    tfactura.serie6 = serie6
    tfactura.serie7 = serie7
    tfactura.numero1 = numero1
    tfactura.numero2 = numero2
    tfactura.numero3 = numero3
    tfactura.numero4 = numero4
    tfactura.numero5 = numero5
    tfactura.numero6 = numero6
    tfactura.numero7 = numero7

    '--------------------- los datos deben ser cargados -------
    I = 0

    If Len(serie1) > 0 Then
        I = I + 1

    End If

    If Len(serie2) > 0 Then
        I = I + 1

    End If

    If Len(serie3) > 0 Then
        I = I + 1

    End If

    If Len(serie4) > 0 Then
        I = I + 1

    End If

    If Len(serie5) > 0 Then
        I = I + 1

    End If

    If I = 1 Then  'si es uno solo debe cargar todo
        pone_campos_db tipo, serie1, numero1

    End If

    logt34342_Click

End Sub

Sub pone_campos_db(buf1 As String, buf2 As String, buf3 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from " & xarchivo & " where local='" & "" & local1 & "' and tipo='" & "" & buf1 & "' and serie='" & "" & buf2 & "' and numero='" & "" & buf3 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        tfactura.partida = "" & mytablex.Fields("partida")
        tfactura.destino = "" & mytablex.Fields("destino")
        tfactura.moneda = "" & mytablex.Fields("moneda")
        'inicio 10/02/2018 pll
        'tfactura.vendedor = "" & mytablex.Fields("vendedor")
        'fin 10/02/2018 pll
   
        tfactura.fpago = "" & mytablex.Fields("fpago")
        tfactura.transporte = "" & mytablex.Fields("transporte")
        tfactura.dias = "" & mytablex.Fields("dias")
        tfactura.bodega = "" & mytablex.Fields("bodega")
        'inicio 10/02/2018 pll
        'tfactura.bodegaf = "" & mytablex.Fields("bodegaf")
        'inicio 10/02/2018 pll
   
        tfactura.observa = "" & mytablex.Fields("observa")

    End If

    mytablex.Close

End Sub

Private Sub Command1_Click()

    Dim buf       As String

    Dim buf3      As String

    Dim rconsulta As New ADODB.Recordset

    If opcion1 = "3" Or opcion1 = "4" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Then
        buf3 = " tipo='" & tipo & "'"

        If acu <> "Q" Then
            buf3 = buf3 & " and codigo='" & codigo & "'"

        End If

    End If

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            'buf = "select Descripcio,Tipo from Tipo where grupo<>'" & acu & "'"
            buf = "select Descripcio,Tipo from Tipo "
        Else
            'buf = "select Descripcio,Tipo from tipo where grupo<>'" & acu & "' and " & Combo1 & " like '" & buffer & "%'"
            buf = "select Descripcio,Tipo from tipo where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "3" Or opcion1 = "4" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Then
        If Len(buffer) = 0 Then
            buf = "select Tipo,Serie,Numero,Fecha,Codigo,Total,Estado,yausado from  " & xarchivo & " where local='" & local1 & "' and " & buf3
        Else
            buf = "select Tipo,Serie,Numero,Fecha,Codigo,Total,Estado,yausado from " & xarchivo & " where local='" & local1 & "' and " & buf3 & " and " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If

    Set dbGrid1.DataSource = rconsulta
               
    If opcion1 = "3" Or opcion1 = "4" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Then
        dbGrid1.columns(0).Width = 700
        dbGrid1.columns(1).Width = 700
        dbGrid1.columns(2).Width = 1500
        dbGrid1.columns(3).Width = 2000
        dbGrid1.columns(4).Width = 1000
        dbGrid1.columns(5).Width = 1000
        dbGrid1.columns(6).Width = 700

    End If

    If opcion1 = "1" Or opcion1 = "2" Or opcion1 = "5" Then
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

    End If

    dbGrid1.SetFocus

End Sub

Private Sub Command2_Click()
    Frame2.Visible = False
    tipo.SetFocus

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    Dim buf   As String

    Dim xtemp As Variant

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            tipo = dbGrid1.columns(1)
            Frame1.Visible = False
            Frame1.Enabled = False
            tipo.SetFocus
            tipo_KeyPress 13

        End If

        If opcion1 = "3" Then
            If dbGrid1.columns(7) = "1" Then
                MsgBox "Documento ya utilizado ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            serie1 = dbGrid1.columns(1)
            numero1 = dbGrid1.columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False
            serie2.SetFocus

        End If

        If opcion1 = "4" Then
            If dbGrid1.columns(7) = "1" Then
                MsgBox "Documento ya utilizado ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            serie2 = dbGrid1.columns(1)
            numero2 = dbGrid1.columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False
            serie3.SetFocus

        End If

        If opcion1 = "6" Then
            If dbGrid1.columns(7) = "1" Then
                MsgBox "Documento ya utilizado ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            serie3 = dbGrid1.columns(1)
            numero3 = dbGrid1.columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False
            serie4.SetFocus

        End If

        If opcion1 = "7" Then
            If dbGrid1.columns(7) = "1" Then
                MsgBox "Documento ya utilizado ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            serie4 = dbGrid1.columns(1)
            numero4 = dbGrid1.columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False
            serie5.SetFocus

        End If

        If opcion1 = "8" Then
            If dbGrid1.columns(7) = "1" Then
                MsgBox "Documento ya utilizado ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            serie5 = dbGrid1.columns(1)
            numero5 = dbGrid1.columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False
            serie6.SetFocus

        End If

        If opcion1 = "9" Then
            If dbGrid1.columns(7) = "1" Then
                MsgBox "Documento ya utilizado ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            serie6 = dbGrid1.columns(1)
            numero6 = dbGrid1.columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False
            serie7.SetFocus

        End If

        If opcion1 = "10" Then
            If dbGrid1.columns(7) = "1" Then
                MsgBox "Documento ya utilizado ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            serie7 = dbGrid1.columns(1)
            numero7 = dbGrid1.columns(2)
            Frame1.Visible = False
            Frame1.Enabled = False

        End If

    End If

End Sub

Private Sub logt34342_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Frame1.Enabled = False

        If opcion1 = "1" Then
            tipo.SetFocus

        End If

        Exit Sub

    End If

    tcrucedo.Hide
    Unload tcrucedo

End Sub

Private Sub serie1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_documentos

    End If

End Sub

Private Sub serie2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_documentos1

    End If

End Sub

Private Sub serie3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_documentos2

    End If

End Sub

Private Sub serie4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_documentos3

    End If

End Sub

Private Sub serie5_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_documentos4

    End If

End Sub

Private Sub serie6_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_documentos5

    End If

End Sub

Private Sub serie7_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_documentos6

    End If

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 27 And KeyAscii <> 13 Then Exit Sub
    If Len(tipo) = 0 Then
        logt34342_Click
        Exit Sub

    End If

    found = busca_tipo(0)

    If found = 0 Then
        tipo = ""
        tipo.SetFocus
        Exit Sub

    End If

    serie1.SetFocus

End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_tipo

    End If

End Sub

Sub consulta_tipo()

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Tipo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Sub consulta_documentos()

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "3"
    Command1_Click

End Sub

Sub consulta_documentos1()

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "4"
    Command1_Click

End Sub

Sub consulta_documentos2()

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "6"
    Command1_Click

End Sub

Sub consulta_documentos3()

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "7"
    Command1_Click

End Sub

Sub consulta_documentos4()

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "8"
    Command1_Click

End Sub

Sub consulta_documentos5()

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "9"
    Command1_Click

End Sub

Sub consulta_documentos6()

    Combo1.Clear
    Combo1.AddItem "Tipo"
    Combo1.AddItem "Serie"
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "10"
    Command1_Click

End Sub

Function busca_tipo(sw As Integer)

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", tipo

    If Not mytablex.NoMatch Then
        busca_tipo = 1

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "B", "C", "D", "G", "E", "F"  'VENTAS
                xarchivo = "FACTURA"
                xarchivo1 = "DETALLE"

            Case "J", "K", "L", "M", "P", "N", "O"  'COMPRAS
                xarchivo = "FACTURA"
                xarchivo1 = "DETALLE"

            Case "H"  'COTIZACION VENTAS
                xarchivo = "CCOTIZAV"
                xarchivo1 = "DCOTIZAV"

            Case "I"  'PEDIDO VENTAS
                xarchivo = "CPEDIDOV"
                xarchivo1 = "DPEDIDOV"

            Case "Q"  'REQUISICION COMPRAS
                xarchivo = "CREQUISA"
                xarchivo1 = "DREQUISA"

                'xarchivo = "CCOTIZAC"
                'xarchivo1 = "DCOTIZAC"
            Case "R"  'ORDEN COMPRA
                xarchivo = "CORDENC"
                xarchivo1 = "DORDENC"

            Case "T", "S" 'GUIA SALIDA"
                xarchivo = "FACTURA"
                xarchivo1 = "DETALLE"

        End Select

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

