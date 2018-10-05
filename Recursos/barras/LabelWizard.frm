VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form LabelWizard 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion Etiqueta"
   ClientHeight    =   5820
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   8070
   ControlBox      =   0   'False
   Icon            =   "LabelWizard.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   5280
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Grabar"
      Height          =   372
      Left            =   1320
      TabIndex        =   28
      Top             =   5280
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir"
      Height          =   372
      Left            =   240
      TabIndex        =   27
      Top             =   5280
      Width           =   972
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Preview (Portrait - A4 size)"
      Height          =   4812
      Left            =   4440
      TabIndex        =   22
      Top             =   240
      Width           =   3372
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4332
         Left            =   240
         ScaleHeight     =   76.465
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   53.975
         TabIndex        =   23
         Top             =   360
         Width           =   3060
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   4210
            Left            =   0
            ScaleHeight     =   73.819
            ScaleMode       =   6  'Millimeter
            ScaleWidth      =   51.065
            TabIndex        =   24
            Top             =   0
            Width           =   2920
            Begin VB.PictureBox alabel 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H80000008&
               Height          =   732
               Index           =   0
               Left            =   120
               ScaleHeight     =   705
               ScaleWidth      =   945
               TabIndex        =   26
               Top             =   960
               Visible         =   0   'False
               Width           =   972
            End
            Begin VB.PictureBox label 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H80000008&
               Height          =   732
               Index           =   0
               Left            =   120
               ScaleHeight     =   705
               ScaleWidth      =   945
               TabIndex        =   25
               Top             =   120
               Visible         =   0   'False
               Width           =   972
            End
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aplicar"
      Default         =   -1  'True
      Height          =   372
      Left            =   5040
      TabIndex        =   15
      Top             =   5280
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   372
      Left            =   6480
      TabIndex        =   14
      Top             =   5280
      Width           =   1332
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Control Hoja"
      Height          =   4812
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3972
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   4212
         Left            =   240
         ScaleHeight     =   4155
         ScaleWidth      =   3435
         TabIndex        =   1
         Top             =   360
         Width           =   3492
         Begin VB.HScrollBar HScroll6 
            Height          =   288
            Left            =   2124
            Max             =   10000
            Min             =   1
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2520
            Value           =   10
            Width           =   400
         End
         Begin VB.HScrollBar HScroll5 
            Height          =   288
            Left            =   2124
            Max             =   10000
            Min             =   1
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   2160
            Value           =   10
            Width           =   400
         End
         Begin VB.TextBox Text6 
            Height          =   288
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "10"
            Top             =   2520
            Width           =   492
         End
         Begin VB.TextBox Text5 
            Height          =   288
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "10"
            Top             =   2160
            Width           =   492
         End
         Begin VB.TextBox Text4 
            Height          =   288
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "1"
            Top             =   1200
            Width           =   492
         End
         Begin VB.TextBox Text3 
            Height          =   288
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "1"
            Top             =   1560
            Width           =   492
         End
         Begin VB.HScrollBar HScroll4 
            Height          =   288
            Left            =   2124
            Max             =   1000
            Min             =   1
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1200
            Value           =   1
            Width           =   400
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   288
            Left            =   2124
            Max             =   10000
            Min             =   1
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1560
            Value           =   1
            Width           =   400
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   288
            Left            =   1640
            Max             =   10000
            Min             =   1
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   600
            Value           =   1
            Width           =   400
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   288
            Left            =   1640
            Max             =   1000
            Min             =   1
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Value           =   1
            Width           =   400
         End
         Begin VB.TextBox Text2 
            Height          =   288
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "1"
            Top             =   600
            Width           =   492
         End
         Begin VB.TextBox Text1 
            Height          =   288
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "1"
            Top             =   240
            Width           =   492
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Page Margin left:"
            Height          =   252
            Left            =   240
            TabIndex        =   17
            Top             =   2520
            Width           =   1332
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Page Margin top:"
            Height          =   252
            Left            =   240
            TabIndex        =   16
            Top             =   2160
            Width           =   1332
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Row spacing:"
            Height          =   252
            Left            =   240
            TabIndex        =   9
            Top             =   1560
            Width           =   1332
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Column spacing:"
            Height          =   252
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   1332
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Columns:"
            Height          =   252
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   852
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Rows:"
            Height          =   252
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   732
         End
      End
   End
End
Attribute VB_Name = "LabelWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I           As Integer

Dim j           As Integer

Dim FileContent As String

Dim FileArray   As Variant

Private Sub Command1_Click()
    Unload Me

End Sub

Private Sub Command2_Click()

    tcxbarra.columns.Caption = Text1.Text
    tcxbarra.rows.Caption = Text2.Text
    tcxbarra.columnspacing.Caption = Text4.Text
    tcxbarra.rowspacing.Caption = Text3.Text
    tcxbarra.pagemargintop.Caption = Text5.Text
    tcxbarra.pagemarginleft.Caption = Text6.Text

    Unload Me

End Sub

Function ShowPreview()

    For I = 1 To label.UBound
        Unload label(I)
    Next I

    For I = 1 To alabel.UBound
        Unload alabel(I)
    Next I

    For I = 1 To Text2.Text 'rows
        Load label(label.count)

        'set properties
        With label(label.UBound)
            .Width = (tcxbarra.label.Width / 4) * 10
            .Height = (tcxbarra.label.Height / 4) * 10

            If I = 1 Then
                .Top = (Text5.Text / 4)
            Else
                .Top = label(label.UBound - 1).Top + label(label.UBound).Height + (Text3.Text / 4)

            End If

            .Left = Text6.Text / 4
            .Visible = True

        End With

        For j = 1 To Text1.Text 'columns
            Load alabel(alabel.count)

            'set properties
            With alabel(alabel.UBound)
                .Width = (tcxbarra.label.Width / 4) * 10
                .Height = (tcxbarra.label.Height / 4) * 10
                .Top = label(I).Top

                If j = 1 Then
                    .Left = (Text6.Text / 4)
                Else
                    .Left = alabel(alabel.UBound - 1).Left + alabel(alabel.UBound).Width + (Text4.Text / 4)

                End If

                .Visible = True

            End With

        Next j
    Next I
    
    Me.Caption = "Label Wizard - " & (Text1.Text * Text2.Text) & " label(s)"

End Function

Private Sub Command3_Click()

    On Error GoTo Err:
    
    With CommonDialog1
        .Filter = "Label Wizard Settings (*.lws)|*.lws|"
        .CancelError = True
        .ShowOpen
        'load data
        Open .FileName For Input As #1

        Do While Not EOF(1)
            Line Input #1, FileContent
            FileArray = Split(FileContent, "=")

            If FileArray(0) = "columns" Then
                Text1.Text = FileArray(1)
                HScroll1.Value = FileArray(1)

            End If

            If FileArray(0) = "rows" Then
                Text2.Text = FileArray(1)
                HScroll2.Value = FileArray(1)

            End If

            If FileArray(0) = "columnspacing" Then
                Text4.Text = FileArray(1)
                HScroll4.Value = FileArray(1)

            End If

            If FileArray(0) = "rowspacing" Then
                Text3.Text = FileArray(1)
                HScroll3.Value = FileArray(1)

            End If

            If FileArray(0) = "topmargin" Then
                Text5.Text = FileArray(1)
                HScroll5.Value = FileArray(1)

            End If

            If FileArray(0) = "leftmargin" Then
                Text6.Text = FileArray(1)
                HScroll6.Value = FileArray(1)

            End If

        Loop
        Close #1
        ShowPreview

    End With
    
Err:
    Exit Sub

End Sub

Private Sub Command4_Click()

    On Error GoTo Err:
    
    With CommonDialog1
        .Filter = "Label Wizard Settings (*.lws)|*.lws|"
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
        Open .FileName For Output As #1
        Print #1, "[LABEL WIZARD SETTINGS]"
        Print #1, "columns=" & Text1.Text
        Print #1, "rows=" & Text2.Text
        Print #1, "columnspacing=" & Text4.Text
        Print #1, "rowspacing=" & Text3.Text
        Print #1, "topmargin=" & Text5.Text
        Print #1, "leftmargin=" & Text6.Text
        Close #1

    End With
    
Err:
    Exit Sub

End Sub

Private Sub Form_Load()
     
    'Picture3.Width = Printer.ScaleWidth 'ancho
    'Picture3.Height = Printer.ScaleHeight 'ancho
    ShowPreview

End Sub

Private Sub Form_Unload(Cancel As Integer)

    PrintManager.Enabled = True
    'PrintManager.Show
    Unload Me

End Sub

Private Sub HScroll1_Change()
    Text1.Text = HScroll1.Value
    ShowPreview

End Sub

Private Sub HScroll2_Change()
    Text2.Text = HScroll2.Value
    ShowPreview

End Sub

Private Sub HScroll3_Change()
    Text3.Text = HScroll3.Value
    ShowPreview

End Sub

Private Sub HScroll4_Change()
    Text4.Text = HScroll4.Value
    ShowPreview

End Sub

Private Sub HScroll5_Change()
    Text5.Text = HScroll5.Value
    ShowPreview

End Sub

Private Sub HScroll6_Change()
    Text6.Text = HScroll6.Value
    ShowPreview

End Sub
