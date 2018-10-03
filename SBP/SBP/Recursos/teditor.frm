VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form teditor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   14865
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo  No poner Extensiones solo Nombre"
      Height          =   4335
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command2 
         Caption         =   "Cerrar"
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox formadest 
         Height          =   615
         Left            =   240
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2040
         Width           =   4815
      End
      Begin VB.TextBox formabase 
         Height          =   615
         Left            =   240
         MaxLength       =   15
         TabIndex        =   3
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nuevo Archivo Formato"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Buscar Archivo Formato Base"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtNotas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   14775
   End
   Begin VB.Menu NU884 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu dl89232 
      Caption         =   "&Abrir"
   End
   Begin VB.Menu dsj7823 
      Caption         =   "&Guardar"
   End
   Begin VB.Menu dl8923 
      Caption         =   "&salir"
   End
End
Attribute VB_Name = "teditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    Frame1.Visible = False

End Sub

Private Sub dl8923_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    teditor.Hide
    Unload teditor

End Sub

Private Sub dl89232_Click()

    If Frame1.Visible = True Then Exit Sub
    CommonDialog1.DialogTitle = "Seleccione un archivo Formato"
    CommonDialog1.InitDir = globaldir & "\formatos"
    CommonDialog1.Filter = "Archivos Formato|*.*"
    CommonDialog1.ShowOpen

    'Si seleccionamos un archivo mostramos la ruta
    If CommonDialog1.FileName <> "" Then
        txtNotas = ""
        teditor.Caption = CommonDialog1.FileName
        carga_archivo teditor.Caption
    Else

        'Si no mostramos un texto de advertencia de que no se seleccionó _   ninguno, ya que FileName devuelve una cadena vacía
        'Label1 = "No se seleccionó ningún archivo"
    End If

End Sub

Private Sub dsj7823_Click()

    Dim found As Integer

    If Frame1.Visible = True Then
        '----proceso copiar al nuevo nombre---
        '----
        '----
        found = genera_copia()

        If found = 0 Then
            MsgBox "No se puede copiar ", 48, "Aviso"
            Exit Sub

        End If

        MsgBox "Proceso Realizado ", 48, "Aviso"
        Frame1.Visible = False
        Exit Sub

    End If

    If Len(teditor.Caption) > 0 Then
        If MsgBox("Desea Guardar", 1, "Aviso") <> 1 Then Exit Sub
        guarda_archivo teditor.Caption
        dl8923_Click

    End If

End Sub

Function genera_copia()

    On Error GoTo cmd90_l1

    If Dir$(globaldir & "\formatos\" & Trim(formadest)) <> "" Then
        Kill (globaldir & "\formatos\" & Trim(formadest))

    End If

    'MsgBox globaldir & "\formatos\" & Trim(formabase)
    FileCopy globaldir & "\formatos\" & Trim(formabase), globaldir & "\formatos\" & Trim(formadest)
    genera_copia = 1
    Exit Function
cmd90_l1:
    MsgBox "Aviso en Genera Copia " + error$, 48, "Aviso"
    Exit Function

End Function

Private Sub Form_Load()

    'Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Sub carga_archivo(nombre As String)

    Dim Tmp As String

    Open nombre For Input As #1

    Do While Not EOF(1)
        Line Input #1, Tmp
        txtNotas.Text = txtNotas.Text & Tmp & Chr(13) & Chr(10)
    Loop
    Close #1

End Sub

Sub guarda_archivo(nombre As String)
    Open nombre For Output As #1
    Print #1, txtNotas.Text
    Close #1

End Sub

Private Sub Label1_Click()
    'MsgBox globaldir
    CommonDialog1.DialogTitle = "Seleccione un archivo Formato"
    CommonDialog1.InitDir = globaldir & "\formatos"
    CommonDialog1.Filter = "Archivos Formato|*.*"
    CommonDialog1.ShowOpen

    'Si seleccionamos un archivo mostramos la ruta
    If CommonDialog1.FileName <> "" Then
        formabase = CommonDialog1.FileName
        formabase = CommonDialog1.FileTitle
    Else

        'Si no mostramos un texto de advertencia de que no se seleccionó _   ninguno, ya que FileName devuelve una cadena vacía
        'Label1 = "No se seleccionó ningún archivo"
    End If

End Sub

Private Sub NU884_Click()
    Frame1.Visible = True

End Sub
