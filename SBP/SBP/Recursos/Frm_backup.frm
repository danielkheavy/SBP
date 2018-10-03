VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup de tablas"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_backup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox fileBckp 
      Height          =   1845
      Left            =   4200
      TabIndex        =   13
      Top             =   1725
      Width           =   3240
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   390
      TabIndex        =   12
      Top             =   1695
      Width           =   3525
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   600
      Left            =   7830
      TabIndex        =   11
      Top             =   780
      Width           =   1275
   End
   Begin VB.Frame FraBackupBase 
      Caption         =   "Backup Base de datos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1545
      Left            =   345
      TabIndex        =   4
      Top             =   75
      Width           =   7155
      Begin MSComCtl2.DTPicker DTFfinal 
         Height          =   315
         Left            =   4785
         TabIndex        =   10
         Top             =   975
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483635
         CalendarForeColor=   -2147483634
         Format          =   234422273
         CurrentDate     =   42851
      End
      Begin MSComCtl2.DTPicker DTFinicio 
         Height          =   285
         Left            =   1650
         TabIndex        =   9
         Top             =   990
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483635
         CalendarForeColor=   -2147483634
         Format          =   234422273
         CurrentDate     =   42851
      End
      Begin VB.OptionButton OptDesempaquetar 
         Caption         =   "Desempaquetar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3435
         TabIndex        =   6
         Top             =   480
         Width           =   2610
      End
      Begin VB.OptionButton OptEmpaquetar 
         Caption         =   "Empaquetar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   300
         TabIndex        =   5
         Top             =   345
         Width           =   1545
      End
      Begin VB.Label lblFechaFinal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3600
         TabIndex        =   8
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label lblFechaInicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   7
         Top             =   1005
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   555
      Left            =   7815
      TabIndex        =   2
      Top             =   1620
      Width           =   1320
   End
   Begin VB.CommandButton cmdProcesa 
      Caption         =   "Procesa"
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   150
      Width           =   1350
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   270
      TabIndex        =   0
      Top             =   4125
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblElaborandoBackup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3645
      TabIndex        =   3
      Top             =   3810
      Width           =   60
   End
End
Attribute VB_Name = "Frm_backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdProcesa_Click()

    Dim fechai       As String

    Dim fechaf       As String

    Dim conta        As Integer

    Dim cantidad     As Long

    Dim my_file_open As String

    Dim myproceded   As Boolean

    Dim myprocedef   As Boolean

    Dim finicio      As String

    Dim ffinal       As String

    fechai = DTFinicio
    fechaf = DTFfinal
    Frm_backup.ProgressBar1.Visible = True

    'poner con fechas en el file
    idia = Mid(DTFinicio, 1, 2)
    imes = Mid(DTFinicio, 4, 2)
    iano = Mid(DTFinicio, 7, 4)
    finicio = idia & imes & iano
    fdia = Mid$(DTFfinal, 1, 2)
    fmes = Mid$(DTFfinal, 4, 2)
    fano = Mid(DTFfinal, 7, 4)
    ffinal = fdia & fmes & fano

    'consistencia mensual
    If imes <> fmes Then
        MsgBox "El mes final debe ser igual al anterior", 64, "Infomacion"
        Exit Sub
    Else

        If OptEmpaquetar.Value = True Then
            OptEmpaquetar.Enabled = False
            OptDesempaquetar.Enabled = False
            'aqui creamos directorio
            crea_directorio
            Call bkp_detalle(fechai, fechaf, conta)

            If conta = 0 Then
                MsgBox "Datos no encontrados", 64, "Infomacion"
                Exit Sub
            Else
                'para factura
                Call bkp_factura(fechai, fechaf)
                'para almacen
                Call bkp_almacen
    
                Call crear_rar(finicio, ffinal)
                'aqui enviar correo
                Call envio_correosBackup
                'aqui eliminamos directorio de empaquetado
                Call eliminar_empaquetado

            End If
   
        Else
            OptDesempaquetar.Enabled = False
            OptEmpaquetar.Enabled = False
            finicio = ""
            ffinal = ""
            'aqui nombre file a copiar
            Call fileBckp_Click

            If myfile = "" Then
                MsgBox "Escoge el Archivo a Copiar"
                Exit Sub
            Else
                Call copia_Decargar 'esto ok
                '       'aqui desempaquetamos
                desampaqueta_rar 'ok
                'para los controles respectivos
                anoi = Mid(myfile, 16, 4)
                mesi = Mid(myfile, 14, 2)
                diai = Mid(myfile, 12, 2)
                finicio = anoi & mesi & diai
                ffinal = Mid(myfile, 24, 4) & Mid(myfile, 22, 2) & Mid(myfile, 20, 2)

                Call control_detalle(finicio, ffinal, myproceded)
                Call control_factura(finicio, ffinal, myprocedef)

                If myproceded = False Or myprocedef = False Then
                    ProgressBar1.Visible = True
                    my_respuesta = MsgBox("Desea efetuar igualmente la transaccion?", vbYesNo, "Backup")

                    If my_respuesta = vbYes Then
                        Call Eli_Bckp_detalle(finicio, ffinal)
                        Call Eli_Bckp_factura(finicio, ffinal)
                        Eli_Bckp_almacen
             
                        'aqui efectua el bakcup
                        'para el detalle
              
                        Call read_cantidad_file_enviado("C:\DesempaquetaVi\bkpdetalle.txt", cantidad, fnum)
                        Call read_save_detalle("C:\DesempaquetaVi\bkpdetalle.txt", cantidad)
                        'para la factura
                        Call read_cantidad_file_enviado("C:\DesempaquetaVi\bkpfactura.txt", cantidad, fnum)

                        If cantidad <> 0 Then
                            Call read_save_factura("C:\DesempaquetaVi\bkpfactura.txt", cantidad)
                        Else
                            Exit Sub

                        End If

                        '              'para almacen
                        Call read_cantidad_file_enviado("C:\DesempaquetaVi\bkpAlmacen.txt", cantidad, fnum)

                        If cantidad <> 0 Then
                            Call read_save_almacen("C:\DesempaquetaVi\bkpAlmacen.txt", cantidad)
                        Else
                            Exit Sub

                        End If
                         
                        Call eliminar_desempaquetado
                    Else
                        Exit Sub

                    End If

                Else
                    ProgressBar1.Visible = True
                    'aqui efectua el bakcup
                    Eli_Bckp_almacen
                    'para el detalle
                    Call read_cantidad_file_enviado("C:\DesempaquetaVi\bkpdetalle.txt", cantidad, fnum)
                    Call read_save_detalle("C:\DesempaquetaVi\bkpdetalle.txt", cantidad)
                    'para la factura
                    Call read_cantidad_file_enviado("C:\DesempaquetaVi\bkpfactura.txt", cantidad, fnum)
                    Call read_save_factura("C:\DesempaquetaVi\bkpfactura.txt", cantidad)
                    'para almacen
                    Call read_cantidad_file_enviado("C:\DesempaquetaVi\bkpAlmacen.txt", cantidad, fnum)
                    Call read_save_almacen("C:\DesempaquetaVi\bkpAlmacen.txt", cantidad)
                        
                    Call eliminar_desempaquetado

                End If

            End If

        End If

    End If

    If OptEmpaquetar.Value = False And OptDesempaquetar.Value = False Then
        ProgressBar1.Visible = False
        MsgBox "Elija una Opcion", 65, "Informacion"
        Call cmdRefresh_Click

    End If

End Sub

Private Sub cmdRefresh_Click()
    Frm_backup.lblElaborandoBackup = ""
    OptEmpaquetar.Enabled = True
    OptDesempaquetar.Enabled = True
    Frm_backup.ProgressBar1.Visible = False
    fileBckp.Visible = False
    Dir1.Visible = False

End Sub

Private Sub cmdSalir_Click()
    Unload Me

End Sub

Private Sub Dir1_Change()

    Dim X As Integer
    
    fileBckp.path = Dir1.path
    Screen.MousePointer = vbHourglass

    For X = 0 To fileBckp.ListCount - 1
        fileBckp.Pattern = "*.rar"
    Next X

    Screen.MousePointer = vbDefault

End Sub

Private Sub fileBckp_Click()
    origen = Dir1.path & "\" & fileBckp.FileName
    myfile = fileBckp.FileName

End Sub

Private Sub Form_Load()
    DTFinicio.Value = Format(Now, "dd/mm/yyyy")
    DTFfinal.Value = Format(Now, "dd/mm/yyyy")
 
    '09/06/2017 kenyo
    Dir1.path = "C:\Users"
    Frm_backup.ProgressBar1.Visible = False
    fileBckp.Visible = False
    'fileBckp.path = "C:\Users\Administrador\Downloads\"
    '09/06/2017 kenyo
  
    Dir1.Visible = False

End Sub

Private Sub OptDesempaquetar_Click()
    lblFechaInicio.Visible = False
    lblFechaFinal.Visible = False
    DTFinicio.Visible = False
    DTFfinal.Visible = False
    OptEmpaquetar.Enabled = False
    fileBckp.Visible = True
    Dir1.Visible = True

End Sub

Public Sub eliminar_desempaquetado()

    On Error GoTo eliminar_desem

    Kill ("C:\DesempaquetaVi\bkpAlmacen.txt")
    Kill ("C:\DesempaquetaVi\bkpdetalle.txt")
    Kill ("C:\DesempaquetaVi\bkpfactura.txt")
    RmDir ("C:\DesempaquetaVi")
    Kill ("C:\" & myfile)
    Dir1.refresh

eliminar_desem:
    Exit Sub

End Sub

Public Sub eliminar_empaquetado()

    Dim fso As New Scripting.FileSystemObject

    On Error GoTo eliminar_em

    Kill ("C:\EmpaquetaVi\bkpAlmacen.txt")
    Kill ("C:\EmpaquetaVi\bkpdetalle.txt")
    Kill ("C:\EmpaquetaVi\bkpfactura.txt")
    '
    'fso.DeleteFolder ("C:\EmpaquetaVi")
    RmDir ("C:\EmpaquetaVi")
    Dir1.refresh

eliminar_em:
    Exit Sub

End Sub

Private Sub OptEmpaquetar_Click()
    DTFinicio.Visible = True
    DTFfinal.Visible = True
    lblFechaInicio.Visible = True
    lblFechaFinal.Visible = True

End Sub

