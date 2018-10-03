VERSION 5.00
Begin VB.Form tcodbar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Codigo Barras"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameBarcode 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   10215
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         FillColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   9915
         TabIndex        =   27
         Top             =   240
         Width           =   9945
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties And Printing"
         Height          =   4695
         Left            =   0
         TabIndex        =   1
         Top             =   1920
         Width           =   9975
         Begin VB.Frame Frame3 
            BackColor       =   &H80000004&
            Caption         =   "Type of Bar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   3765
            Begin VB.OptionButton optBar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Bar 39"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   18
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optBar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Bar 128"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   17
               Top             =   720
               Width           =   855
            End
            Begin VB.ComboBox cboBarSize 
               Height          =   315
               ItemData        =   "tcodbar.frx":0000
               Left            =   2430
               List            =   "tcodbar.frx":000D
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   300
               Width           =   975
            End
            Begin VB.ComboBox cboTextStyle 
               Height          =   315
               ItemData        =   "tcodbar.frx":0027
               Left            =   2430
               List            =   "tcodbar.frx":0037
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   680
               Width           =   975
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Bar Size"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   20
               Top             =   375
               Width           =   855
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Text Style"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   1560
               TabIndex        =   19
               Top             =   735
               Width           =   975
            End
         End
         Begin VB.Frame fr128 
            BackColor       =   &H80000004&
            Caption         =   "BarCode 128 Properties"
            Height          =   1185
            Left            =   3960
            TabIndex        =   10
            Top             =   240
            Width           =   4485
            Begin VB.CheckBox chkTextAlignment 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Bottom Align Caption"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   360
               TabIndex        =   13
               Top             =   720
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox checkBarCaption 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Bar With Caption"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   360
               TabIndex        =   12
               Top             =   360
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.CheckBox Check1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Top Align Caption"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2400
               TabIndex        =   11
               Top             =   360
               Value           =   1  'Checked
               Width           =   1935
            End
         End
         Begin VB.PictureBox picSample3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   857
            Left            =   4920
            ScaleHeight     =   825
            ScaleWidth      =   2175
            TabIndex        =   9
            Top             =   2640
            Width           =   2203
         End
         Begin VB.PictureBox picSample2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   857
            Left            =   7200
            ScaleHeight     =   825
            ScaleWidth      =   2175
            TabIndex        =   8
            Top             =   1680
            Width           =   2203
         End
         Begin VB.PictureBox picSample1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   857
            Left            =   4920
            ScaleHeight     =   825
            ScaleWidth      =   2175
            TabIndex        =   7
            Top             =   1680
            Width           =   2203
         End
         Begin VB.TextBox txtBCode 
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   6
            Text            =   "0.2"
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox txtBCode 
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   5
            Text            =   "0.2"
            Top             =   3840
            Width           =   495
         End
         Begin VB.CommandButton cmdPageSetup 
            Caption         =   "Page Setup"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton cmdPreview 
            Caption         =   "Print Preview"
            Height          =   375
            Left            =   2160
            TabIndex        =   3
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   4200
            Width           =   1695
         End
         Begin VB.Label lblProp 
            AutoSize        =   -1  'True
            Caption         =   "(2) Distance between barcodes :"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   3120
            Width           =   2310
         End
         Begin VB.Label lblProp 
            AutoSize        =   -1  'True
            Caption         =   "(a) Barcode 1 to Barcode 2 :              inch/es"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   25
            Top             =   3480
            Width           =   3180
         End
         Begin VB.Label lblProp 
            AutoSize        =   -1  'True
            Caption         =   "(b) Barcode 1 to Barcode 3 :              inch/es"
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   24
            Top             =   3840
            Width           =   3180
         End
         Begin VB.Label lblProp 
            AutoSize        =   -1  'True
            Caption         =   "(b) Height : XXXX inch/es"
            Height          =   195
            Index           =   8
            Left            =   600
            TabIndex        =   23
            Tag             =   "(b) Height : XXXX inch/es"
            Top             =   2760
            Width           =   1830
         End
         Begin VB.Label lblProp 
            AutoSize        =   -1  'True
            Caption         =   "(a) Width : XXXX inch/es"
            Height          =   195
            Index           =   9
            Left            =   600
            TabIndex        =   22
            Tag             =   "(a) Width : XXXX inch/es"
            Top             =   2400
            Width           =   1785
         End
         Begin VB.Label lblProp 
            AutoSize        =   -1  'True
            Caption         =   "(1) Paper Size :"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   21
            Top             =   2040
            Width           =   1080
         End
      End
   End
   Begin VB.Menu di234 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcodbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPageSetup_Click()
Dim MyLength        As Single
    MyLength = ShowPageSetupDlg
    'Call PaperSize
    If MyLength > -1 Then
        MyLength = (12210 / 8500) * 1000
        lblProp(9).Caption = Replace(lblProp(9).Tag, "XXXX", FarRightMargin / MyLength)
        lblProp(8).Caption = Replace(lblProp(8).Tag, "XXXX", FarTopMargin / MyLength)
    End If

End Sub

Private Sub cmdPreview_Click()
If Val(txtBCode(0)) = 0 Or Val(txtBCode(1)) = 0 Then
    MsgBox "Please input values in the Distance between barcodes. Thank you", vbInformation, "Information"
    Exit Sub
End If
If FarRightMargin = 0 Or FarTopMargin = 0 Then
    MsgBox "Please set the papersize using page setup. Thank you", vbInformation, "Information"
    Exit Sub
End If
    Load frmPreview
    frmPreview.ZOrder 0
    frmPreview.picPreview.Cls
    frmPreview.PaperSize
    Call CreateLabel(Me.picSample1, frmPreview.picPreview)

End Sub

Private Sub cmdPrint_Click()
If Val(txtBCode(0)) = 0 Or Val(txtBCode(1)) = 0 Then
    MsgBox "Please input values in the Distance between barcodes. Thank you", vbInformation, "Information"
    Exit Sub
End If
If FarRightMargin = 0 Or FarTopMargin = 0 Then
    MsgBox "Please set the papersize using page setup. Thank you", vbInformation, "Information"
    Exit Sub
End If
    Call CreateLabel(picSample1, Printer)
    Printer.EndDoc

End Sub

Private Sub di234_Click()
tcodbar.Hide
Unload tcodbar
End Sub
