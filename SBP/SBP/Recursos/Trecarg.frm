VERSION 5.00
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form Trecarg 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6585
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox valor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2340
      MaxLength       =   10
      TabIndex        =   23
      Top             =   2310
      Width           =   1980
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   6735
      TabIndex        =   0
      Top             =   3930
      Width           =   2970
   End
   Begin ChamaleonButton.ChameleonBtn cmd1 
      Height          =   855
      Left            =   270
      TabIndex        =   2
      Top             =   600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "% -"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "Trecarg.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmd3 
      Height          =   855
      Left            =   3315
      TabIndex        =   4
      Top             =   600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "% +"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "Trecarg.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmd2 
      Height          =   855
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Monto -"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "Trecarg.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmd4 
      Height          =   855
      Left            =   4815
      TabIndex        =   6
      Top             =   600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Borra Dscto"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "Trecarg.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   15
      Left            =   555
      TabIndex        =   8
      Top             =   2940
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   0
      Left            =   1500
      TabIndex        =   9
      Top             =   2925
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   1
      Left            =   2445
      TabIndex        =   10
      Top             =   2940
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   2
      Left            =   3390
      TabIndex        =   11
      Top             =   2940
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   3
      Left            =   555
      TabIndex        =   12
      Top             =   3795
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   4
      Left            =   1500
      TabIndex        =   13
      Top             =   3795
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   5
      Left            =   2445
      TabIndex        =   14
      Top             =   3795
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   6
      Left            =   3390
      TabIndex        =   15
      Top             =   3795
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   7
      Left            =   570
      TabIndex        =   16
      Top             =   4665
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   8
      Left            =   1500
      TabIndex        =   17
      Top             =   4665
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   9
      Left            =   2445
      TabIndex        =   18
      Top             =   4665
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
      PicturePosition =   0
      Caption         =   "."
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
   Begin GorditoButton.Boton znumbert 
      Height          =   825
      Index           =   11
      Left            =   3390
      TabIndex        =   19
      Top             =   4680
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
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
   Begin GorditoButton.Boton cmdok 
      Height          =   1035
      Left            =   4590
      TabIndex        =   20
      Top             =   2955
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1826
      PicturePosition =   4
      Caption         =   "Ok"
      BackColor       =   4210752
      ResalteColor    =   49152
      PictureDown     =   "Trecarg.frx":0070
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
   Begin GorditoButton.Boton cmdClose 
      Height          =   1035
      Left            =   4620
      TabIndex        =   21
      ToolTipText     =   "Cancelar"
      Top             =   4065
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1826
      PicturePosition =   4
      Caption         =   "X"
      BackColor       =   4210752
      ResalteColor    =   12632256
      PictureDown     =   "Trecarg.frx":0EFA
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
   Begin VB.Line Line1 
      X1              =   285
      X2              =   6150
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione tipo de dscto.:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   22
      Top             =   165
      Width           =   2865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Dscto:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2310
      TabIndex        =   7
      Top             =   1950
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Total"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   615
      TabIndex        =   3
      Top             =   1950
      Width           =   1365
   End
   Begin VB.Label total 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2310
      Width           =   1410
   End
   Begin VB.Menu flo89923 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Trecarg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dj782321_Click()
    flo89923_Click

End Sub

Private Sub cmd_Click()

End Sub

Private Sub cmd3_Click()
    List1.ListIndex = 3
 
    cmd3.BackColor = &H80FFFF
    cmd3.BackOver = &H80FFFF
 
    cmd1.BackColor = &H404040
    cmd1.BackOver = &H404040
 
    cmd2.BackColor = &H404040
    cmd2.BackOver = &H404040
 
    cmd4.BackColor = &H404040
    cmd4.BackOver = &H404040
 
    valor.SetFocus
    tipodescuento = 1

End Sub

Private Sub cmd2_Click()
    List1.ListIndex = 1
    cmd2.BackColor = &H80FFFF
    cmd2.BackOver = &H80FFFF
 
    cmd1.BackColor = &H404040
    cmd1.BackOver = &H404040
 
    cmd3.BackColor = &H404040
    cmd3.BackOver = &H404040
 
    cmd4.BackColor = &H404040
    cmd4.BackOver = &H404040
    valor.SetFocus
    tipodescuento = 2

End Sub

Private Sub cmd1_Click()
    List1.ListIndex = 0
 
    cmd1.BackColor = &H80FFFF
    cmd1.BackOver = &H80FFFF
 
    cmd2.BackColor = &H404040
    cmd2.BackOver = &H404040
 
    cmd3.BackColor = &H404040
    cmd3.BackOver = &H404040
 
    cmd4.BackColor = &H404040
    cmd4.BackOver = &H404040
  
    valor.SetFocus
    tipodescuento = 0

End Sub

Private Sub cmd4_Click()
    List1.ListIndex = 2
    cmd4.BackColor = &H80FFFF
    cmd4.BackOver = &H80FFFF
 
    cmd1.BackColor = &H404040
    cmd1.BackOver = &H404040
 
    cmd2.BackColor = &H404040
    cmd2.BackOver = &H404040
 
    cmd3.BackColor = &H404040
    cmd3.BackOver = &H404040
 
    valor.SetFocus
    tipodescuento = 3

End Sub

Private Sub cmdClose_Click()
    Trecarg.Hide
    Unload Trecarg

End Sub

Private Sub cmdOK_Click()

    If List1.ListIndex = -1 Then
        valor.SetFocus
        MsgBox "Seleccione una opción..."
        Exit Sub
    Else

        If List1.ListIndex = 1 Then
            If Val(valor.Text) > Val(total.Caption) Then
                MsgBox "Rango Descuentos Erroneo", 24, "Aviso"
                valor.SetFocus
                Exit Sub

            End If

        End If

        If Not IsNumeric(valor.Text) Then
            valor = ""
            valor.SetFocus
            Exit Sub

        End If

        tipodescuento = "" & List1.ListIndex
           
        valordescuento = valor
       
        flo89923_Click
       
    End If
       
End Sub

Private Sub flo89923_Click()
    Trecarg.Hide
    Unload Trecarg

End Sub

Private Sub Form_Load()
    List1.Clear
    List1.AddItem "DESCUENTO EN % -TOTAL"
    List1.AddItem "DESCUENTO EN MONTO - TOTAL"
    List1.AddItem "BORRAR TODOS LOS DESCUENTOS TOTALES"
    List1.AddItem "RECARGO EN % -TOTAL "

    'List1.ListIndex = 0
End Sub

Private Sub List1_DblClick()
    List1_KeyPress 13

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        flo89923_Click
        Exit Sub

    End If

    'valor.SetFocus
End Sub

Private Sub valor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdOK_Click

    End If

    ''If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    ''If KeyAscii = 27 Then
    ''   flo89923_Click
    ''   Exit Sub
    ''End If
    ''       If List1.ListIndex = 1 Then
    ''          If Val(valor) > Val(total) Then
    ''             MsgBox "Rango Descuentos Erroneo", 24, "Aviso"
    ''             valor.SetFocus
    ''             Exit Sub
    ''          End If
    ''       End If
    ''       If Not IsNumeric(valor) Then
    ''          valor = ""
    ''          valor.SetFocus
    ''          Exit Sub
    ''       End If
    ''       tipodescuento = "" & List1.ListIndex
    ''       valordescuento = valor
    ''       flo89923_Click
End Sub

Private Sub znumbert_Click(Index As Integer)

    If Index = 11 Then
        valor = ""
        Exit Sub

    End If

    valor = valor & znumbert(Index).Caption

End Sub
