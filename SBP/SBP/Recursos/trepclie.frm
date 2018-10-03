VERSION 5.00
Begin VB.Form trepclie 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Clientes"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Transaccion"
      Height          =   375
      Left            =   4680
      TabIndex        =   36
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox fechaft 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   34
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox fechait 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   33
      Top             =   4920
      Width           =   2295
   End
   Begin VB.ComboBox meses 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   4080
      Width           =   3615
   End
   Begin VB.TextBox dia2 
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
      Left            =   3240
      MaxLength       =   11
      TabIndex        =   29
      Text            =   "%"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox dia1 
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
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   27
      Text            =   "%"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox fechai 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   24
      Text            =   "%"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox fechaf 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   23
      Text            =   "%"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8475
      TabIndex        =   20
      Top             =   0
      Width           =   8535
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
         Picture         =   "trepclie.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Imprimir"
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "trepclie.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox telefono 
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
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   18
      Text            =   "%"
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox sexo 
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
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "%"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox distrito 
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
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   14
      Text            =   "%"
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox direccion 
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
      Left            =   2280
      MaxLength       =   60
      TabIndex        =   12
      Text            =   "%"
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox codigo 
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
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   5
      Text            =   "%"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox codigo1 
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
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   4
      Text            =   "%"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox nombre 
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
      Left            =   2280
      MaxLength       =   60
      TabIndex        =   3
      Text            =   "%"
      Top             =   1440
      Width           =   3615
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
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   2
      Top             =   5880
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
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   1
      Text            =   "45"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   6600
      Width           =   3855
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
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
      Left            =   120
      TabIndex        =   35
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
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
      Left            =   120
      TabIndex        =   32
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cumplen años mes "
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
      Left            =   120
      TabIndex        =   30
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dias"
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
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaNacim. Inicio"
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
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaNacim. Final"
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
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefono"
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
      Left            =   120
      LinkTimeout     =   60
      TabIndex        =   19
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sexo"
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
      Left            =   120
      LinkTimeout     =   60
      TabIndex        =   17
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distrito"
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
      Left            =   120
      LinkTimeout     =   60
      TabIndex        =   15
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Direccion"
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
      Left            =   120
      LinkTimeout     =   60
      TabIndex        =   13
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
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
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CodigoAlterno"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
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
      Left            =   120
      LinkTimeout     =   60
      TabIndex        =   9
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      Left            =   0
      TabIndex        =   8
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label4 
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
      Left            =   0
      TabIndex        =   7
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agrupacion"
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
      Left            =   0
      TabIndex        =   6
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Menu dju2323 
      Caption         =   "&Buscar"
   End
   Begin VB.Menu dlo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trepclie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    dlo232_Click

End Sub

Private Sub cmdPrint_Click()
    dju2323_Click

End Sub

Private Sub dju2323_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    If Check1.Value = 1 Then
        If Not IsDate(fechait) Then Exit Sub
        If Not IsDate(fechaft) Then Exit Sub

    End If

    found = sql_cuentaxc(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_cuentaxc
    cuerpo_programa_cuentaxc mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub cabecera_cuentaxc()

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
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(152, "_")
    found = formateaa(buf, 152, 2, 0)
    found = formateaa("Codigo", 12, 0, 0)
    found = formateaa("Nombre", 41, 0, 0)
    found = formateaa("Direccion", 51, 0, 0)
    found = formateaa("Distrito", 11, 0, 0)
    found = formateaa("Telefono", 9, 0, 0)
    found = formateaa("FechaNac", 11, 2, 0)
    buf = String(152, "_")
    found = formateaa(buf, 152, 2, 0)

End Sub

Sub cuerpo_programa_cuentaxc(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim tmp1  As String

    Dim sw    As Integer

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    sdx = 0
    sw = 0
    suma1 = 0
    ssuma1 = 0
    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("tipoclie")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("zona")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Distrito" Then
            tmp1 = "" & mytablex.Fields("Distrito")

        End If

        If sw = 0 Then
            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("tipoclie")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("tipoclie")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("Vendedor")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Vendedor")

            End If

            If Combo1 = "Distrito" Then
                buf = "" & mytablex.Fields("distrito")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("distrito")

            End If

            sw = 1

        End If

        If Tmp <> tmp1 Then
            found = formateaa("SubGrupo ", 37, 0, 1)
            buf = Format(suma1, "00000")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
   
            If Combo1 = "tipoclie" Then
                buf = "" & mytablex.Fields("tipoclie")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("tipoclie")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("Vendedor")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Vendedor")

            End If

            If Combo1 = "Distrito" Then
                buf = "" & mytablex.Fields("distrito")
                found = formateaa(buf, 11, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("distrito")

            End If

            suma1 = 0
            ssuma1 = 0

        End If

        buf = "" & mytablex.Fields("Codigo")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Nombre")
        found = formateaa(buf, 40, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Direccion")
        found = formateaa(buf, 50, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Distrito")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Telefono")
        found = formateaa(buf, 8, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Fechanac")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas

        If Check1.Value = 1 Then  'si es con transaccion
            imprime_transaccion "" & mytablex.Fields("codigo")

        End If

        suma1 = suma1 + 1
        ssuma1 = ssuma1 + 1
        mytablex.MoveNext
    Loop
    found = formateaa("Sub Grupo ", 37, 0, 1)
    buf = Format(suma1, "00000")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("Grupo ", 37, 0, 1)
    buf = Format(ssuma1, "00000")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        cabecera_cuentaxc

    End If

End Sub

Function busca_nombre(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from clientes where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_nombre = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Function sql_cuentaxc(mytablex As ADODB.Recordset)

    Dim buf  As String

    Dim xmes As String

    On Error GoTo cmd3_err

    buf = "select * from clientes where "
    buf = buf & " codigo like '" & codigo & "'"

    If fechai <> "%" And fechaf <> "%" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and  fecha<='" & Format(fechaf, "YYYYMMDD") & "'"

    End If

    Select Case meses

        Case "Enero"
            xmes = "01"

        Case "Febrero"
            xmes = "02"

        Case "Marzo"
            xmes = "03"

        Case "Abril"
            xmes = "04"

        Case "Mayo"
            xmes = "05"

        Case "Junio"
            xmes = "06"

        Case "Julio"
            xmes = "07"

        Case "Agosto"
            xmes = "08"

        Case "Setiembre"
            xmes = "09"

        Case "Octubre"
            xmes = "10"

        Case "Noviembre"
            xmes = "11"

        Case "Diciembre"
            xmes = "12"

    End Select

    If meses <> "%" Then
        buf = buf & " and month(fechanac)=" & Val(xmes)

    End If

    If dia1 <> "%" And dia2 <> "%" Then
        buf = buf & " and day(fechanac)>=" & Val(dia1)
        buf = buf & " and day(fechanac)<=" & Val(dia2)

    End If

    If Combo1 = "Tipoclie" Then
        buf = buf & " order by TipocliE,Nombre"

    End If

    If Combo1 = "Zona" Then
        buf = buf & " order by Zona,Nombre"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " order by Vendedor,Nombre"

    End If

    If Combo1 = "Distrito" Then
        buf = buf & " order by Distrito,Nombre"

    End If

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_cuentaxc = 1
    Exit Function
cmd3_err:
    Exit Function

End Function

Private Sub dlo232_Click()
    trepclie.Hide
    Unload trepclie

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Tipoclie"
    Combo1.AddItem "Zona"
    Combo1.AddItem "Vendedor"
    Combo1.AddItem "Distrito"
    Combo1.ListIndex = 0
    meses.Clear
    meses.AddItem "%"
    meses.AddItem "Enero"
    meses.AddItem "Febrero"
    meses.AddItem "Marzo"
    meses.AddItem "Abril"
    meses.AddItem "Mayo"
    meses.AddItem "Junio"
    meses.AddItem "Julio"
    meses.AddItem "Agosto"
    meses.AddItem "Setiembre"
    meses.AddItem "Octubre"
    meses.AddItem "Noviembre"
    meses.AddItem "Diciembre"
    meses.ListIndex = 0
    fechait = Format(Now, "dd/mm/yyyy")
    fechaft = Format(Now, "dd/mm/yyyy")

End Sub

Sub imprime_transaccion(buf1 As String)

    Dim buf      As String

    Dim buf2     As String

    Dim mytablex As New ADODB.Recordset

    buf2 = "select * from factura where "
    buf2 = buf2 & "  fecha>='" & Format(fechait, "YYYYMMDD") & "'"
    buf2 = buf2 & " and  fecha<='" & Format(fechaft, "YYYYMMDD") & "'"
    buf2 = buf2 & " and codigo='" & buf1 & "'"
    buf2 = buf2 & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"
    buf2 = buf2 & " and estado='2' "
    buf2 = buf2 & " order by fecha"
    'MsgBox buf2
    mytablex.Open buf2, cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        found = formateaa("", 5, 0, 0)
        found = formateaa("----", 5, 0, 0)
        buf = "" & mytablex.Fields("local")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("tipo")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("serie")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("total")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        imprime_fpagov mytablex
        found = formateaa("", 1, 2, 0)
        nlineas
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub imprime_fpagov(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    mytablex.Open "select * from fpagov where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        found = formateaa("", 10, 0, 0)
        found = formateaa("++++", 5, 0, 0)
        buf = "" & mytablex.Fields("fpago")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = Format(Val("" & mytablex.Fields("recibe")), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        If "" & mytablex.Fields("acufp") = "C" Then
            busca_creditox mytablex

        End If

        found = formateaa("", 1, 2, 0)
        nlineas
   
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub busca_creditox(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim sdx      As Double

    Dim sdx1     As Double

    sdx = 0
    sdx1 = 0
    mytablex.Open "select * from cuentac where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("abono"))
        sdx1 = Val("" & mytablex.Fields("saldo"))

    End If

    mytablex.Close
    found = formateaa("Abono ", 6, 0, 0)
    buf = Format(sdx, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Saldo ", 6, 0, 0)
    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)

End Sub
