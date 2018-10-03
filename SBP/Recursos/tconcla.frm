VERSION 5.00
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Begin VB.Form tconcla 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clave de Acceso"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox clave 
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
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   3150
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   1380
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "0"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   1
      Left            =   1005
      TabIndex        =   5
      Top             =   1380
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "1"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   2
      Left            =   1800
      TabIndex        =   6
      Top             =   1380
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "2"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   3
      Left            =   2595
      TabIndex        =   7
      Top             =   1380
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "3"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   4
      Left            =   225
      TabIndex        =   8
      Top             =   2085
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "4"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   5
      Left            =   1005
      TabIndex        =   9
      Top             =   2085
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "5"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   6
      Left            =   1800
      TabIndex        =   10
      Top             =   2085
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "6"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   7
      Left            =   2595
      TabIndex        =   11
      Top             =   2085
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "7"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   8
      Left            =   240
      TabIndex        =   12
      Top             =   2790
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "8"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   9
      Left            =   1005
      TabIndex        =   13
      Top             =   2790
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "9"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton Comando 
      Height          =   750
      Index           =   10
      Left            =   1800
      TabIndex        =   14
      Top             =   2790
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1323
      PicturePosition =   0
      Caption         =   "CR"
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton cmdIngresar 
      Height          =   1035
      Left            =   3630
      TabIndex        =   15
      Top             =   1305
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1826
      PicturePosition =   4
      Caption         =   "Ok"
      BackColor       =   4210752
      ResalteColor    =   49152
      PictureDown     =   "tconcla.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton cmdCancelar 
      Height          =   1035
      Left            =   3660
      TabIndex        =   16
      ToolTipText     =   "Cancelar"
      Top             =   2415
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1826
      PicturePosition =   4
      Caption         =   "X"
      BackColor       =   4210752
      ResalteColor    =   12632256
      PictureDown     =   "tconcla.frx":0E8A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin VB.Label lugar 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2010
      TabIndex        =   3
      Top             =   3675
      Width           =   975
   End
   Begin VB.Label X 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   570
      TabIndex        =   2
      Top             =   3675
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave de Acceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   255
      Width           =   3150
   End
   Begin VB.Menu fdl34343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tconcla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clave_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        flag_clave1 = 0
        fdl34343_Click
        Exit Sub

    End If

    If Len(clave) = 0 Then
        clave.SetFocus
        Exit Sub

    End If

    found = busca_clave("" & clave)

    If found = 0 Then
        MsgBox "No existe Clave", 48, "Aviso"
        clave = ""
        flag_clave1 = 0
        Exit Sub

    End If

    flag_clave1 = 1

    If tconcla.X = "CP" Then
  
        tkeyboar.FLAG = "PRECIO"
        tkeyboar.Show 1
        fdl34343_Click
    Else
        fdl34343_Click

    End If

End Sub

Private Sub cmdCancelar_Click()
    flag_clave1 = 0
    fdl34343_Click

End Sub

Private Sub cmdIngresar_Click()
    clave_KeyPress 13

End Sub

Private Sub Comando_Click(Index As Integer)

    If Index = 10 Then
        clave.Text = ""
        Exit Sub

    End If

    clave = clave + Comando(Index).Caption

End Sub

Private Sub fdl34343_Click()
    tconcla.Hide
    Unload tconcla

End Sub

Function busca_clave(buf As String)

    Dim mytablex As New ADODB.Recordset

    buf = UCase$(buf)
    busca_clave = 0
    globalmesero = ""
    mytablex.Open "select * from vendedor where clave='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        globalmesero = Trim("" & mytablex.Fields("codigo"))

        If X = "CAMBIACODIGO" Then
            If "" & mytablex.Fields("cprecios") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "CIERRE" Then
            If "" & mytablex.Fields("cierre") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "RELOJ" Then
            If "" & mytablex.Fields("clavere") = "2" Then
                busca_clave = 1

            End If

        End If
   
        If X = "CUADRE" Then
            If "" & mytablex.Fields("cuadre") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "MINIREPORTE" Then
            If "" & mytablex.Fields("minireporte") = "S" Then
                busca_clave = 1

            End If

        End If
   
        If X = "CIERRE" Then
            If "" & mytablex.Fields("cierre") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "BORRA COMANDA" Then
            If "" & mytablex.Fields("borra_comanda") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "DESCUENTO" Then
            If "" & mytablex.Fields("descuento") = "S" Then
                busca_clave = 1

            End If

        End If
   
        If X = "ANULA" Then
            If "" & mytablex.Fields("anula") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "CLEAR" Then
            If "" & mytablex.Fields("anula") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "CAMBIOS" Then
            If "" & mytablex.Fields("cprecios") = "S" Then
                busca_clave = 1

            End If

        End If
   
        If X = "COPIA" Then
            If "" & mytablex.Fields("copia") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "CONGELA" Then
            If "" & mytablex.Fields("congela") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "APERTURA" Then
            If "" & mytablex.Fields("apertura") = "S" Then
                busca_clave = 1

            End If

        End If

        If X = "COMANDA" Then
            If "" & mytablex.Fields("despacho") = "S" Then
                busca_clave = 1
                tcomanda.mesero = mytablex.Fields("codigo")
                tcomanda.nmesero = mytablex.Fields("nombre")

            End If

        End If

        If X = "IMPORTACION" Then
            If "" & mytablex.Fields("vecostoimp") = "S" Then
                busca_clave = 1

            End If

        End If

        ''''12/08/2017 kenyo Bloqueo de Cambio de Precio Teclado Virtual

        '13/08/2018 Integración FE - Pizzeria
        ''' 16/07/2018 Clave en forma de pago
        If X = "FORMAPAGO" Then
            If "" & mytablex.Fields("caja") = "1" Or "" & mytablex.Fields("caja") = "2" Or "" & mytablex.Fields("caja") = "3" Then
                busca_clave = 1

            End If

        End If

        ''' 16/07/2018 Clave en forma de pago
        '13/08/2018 Integración FE - Pizzeria

        If X = "S" Or X = "S" Or X = "N" Or X = "CP" Then  'si es supervisor o normal
            If "" & mytablex.Fields("caja") = "1" Or "" & mytablex.Fields("caja") = "2" Or "" & mytablex.Fields("caja") = "3" Then
                busca_clave = 1

            End If

        End If

    End If

    'MsgBox busca_clave
    '------------------------------------- ------------
    mytablex.Close

End Function

Private Sub Form_Load()

    Dim I As Integer

    'For i = 0 To 10
    'comando(i).Sound = App.path & "\Sonido\click.wav": comando(i).PlaySound = InClick
    'comando(i).Sound = "C:\Windows\click.wav": comando(i).PlaySound = InClick
    'Next
End Sub
