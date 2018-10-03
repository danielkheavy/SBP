VERSION 5.00
Begin VB.Form PrintManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   2835
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   5790
   ControlBox      =   0   'False
   Icon            =   "Print.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   2280
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Label Wizard"
      Height          =   372
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuracion Dispositivo"
      Height          =   1932
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5292
      Begin VB.ComboBox Combo3 
         Height          =   288
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   1452
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   3372
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   288
         Left            =   960
         Max             =   10000
         Min             =   1
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1200
         Value           =   1
         Width           =   400
      End
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "1"
         Top             =   1224
         Width           =   492
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1452
      End
      Begin VB.Label Label3 
         Caption         =   "Orientation:"
         Height          =   252
         Left            =   3600
         TabIndex        =   10
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Impresora:"
         Height          =   252
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Numero de copias:"
         Height          =   252
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Quality:"
         Height          =   252
         Left            =   1920
         TabIndex        =   7
         Top             =   960
         Width           =   612
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir..."
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   4320
      TabIndex        =   0
      Top             =   2280
      Width           =   1212
   End
End
Attribute VB_Name = "PrintManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I As Integer

Dim j As Integer

Dim SetCurrentTop

Dim Pr

Dim SetWidth

Dim SetHeight

Dim SetCurrentLeft

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    Dim count As Integer

    If MsgBox("Are you sure you would like to print now?", vbQuestion + vbYesNo, "Print") = vbYes Then

        Set tcxbarra.LabelImage.Picture = tcxbarra.LabelImage.Image

        'set printer device
        For Each Pr In Printers

            If Pr.DeviceName = Combo1.List(Combo1.ListIndex) Then
                Set Printer = Pr
                Exit For

            End If

        Next
        count = 1
        'set printer settings
        Printer.ScaleMode = vbCentimeters

        If Combo1.Text = "Draft" Then
            Printer.PrintQuality = vbPRPQDraft
        ElseIf Combo1.Text = "Low" Then
            Printer.PrintQuality = vbPRPQLow
        ElseIf Combo1.Text = "Medium" Then
            Printer.PrintQuality = vbPRPQMedium
        ElseIf Combo1.Text = "High" Then
            Printer.PrintQuality = vbPRPQHigh

        End If

        If Combo3.Text = "Portrait" Then
            Printer.Orientation = cdlPortrait
        ElseIf Combo3.Text = "Landscape" Then
            Printer.Orientation = cdlLandscape

        End If

        'get set label sizes
        SetWidth = tcxbarra.label.Width
        SetHeight = tcxbarra.label.Height

        For I = 1 To tcxbarra.rows.Caption 'rows

            If I = 1 Then
                SetCurrentTop = (tcxbarra.pagemargintop.Caption / 10)
            Else
                SetCurrentTop = (SetCurrentTop + SetHeight) + (tcxbarra.rowspacing.Caption / 10)

            End If

            'set columns
            For j = 1 To tcxbarra.columns.Caption 'columns

                If j = 1 Then
                    SetCurrentLeft = (tcxbarra.pagemarginleft.Caption / 10)
                Else
                    SetCurrentLeft = (SetCurrentLeft + SetWidth) + (tcxbarra.columnspacing.Caption / 10)

                End If

                'print the label
                If count <= Val(Text3) Then
                    Printer.PaintPicture tcxbarra.LabelImage.Picture, SetCurrentLeft, SetCurrentTop, SetWidth, SetHeight
                    count = count + 1
                Else
                    Exit For

                End If

            Next j
        Next I

        Printer.EndDoc
        Unload Me

    End If

End Sub

Private Sub Command3_Click()

    Me.Enabled = False
    LabelWizard.Text1.Text = tcxbarra.columns.Caption
    LabelWizard.Text2.Text = tcxbarra.rows.Caption
    LabelWizard.Text3.Text = tcxbarra.rowspacing.Caption
    LabelWizard.Text4.Text = tcxbarra.columnspacing.Caption
    LabelWizard.Text5.Text = tcxbarra.pagemargintop.Caption
    LabelWizard.Text6.Text = tcxbarra.pagemarginleft.Caption
    LabelWizard.HScroll1.Value = tcxbarra.columns.Caption
    LabelWizard.HScroll2.Value = tcxbarra.rows.Caption
    LabelWizard.HScroll3.Value = tcxbarra.rowspacing.Caption
    LabelWizard.HScroll4.Value = tcxbarra.columnspacing.Caption
    LabelWizard.HScroll5.Value = tcxbarra.pagemargintop.Caption
    LabelWizard.HScroll6.Value = tcxbarra.pagemarginleft.Caption
    LabelWizard.Show 1

End Sub

Private Sub Form_Load()

    For Each Pr In Printers

        Combo1.AddItem Pr.DeviceName
    Next
    Combo1.Text = Printer.DeviceName

    Combo2.AddItem "Draft"
    Combo2.AddItem "Low"
    Combo2.AddItem "Medium"
    Combo2.AddItem "High"
    Combo2.Text = "Medium"
    
    Combo3.AddItem "Portrait"
    Combo3.AddItem "Landscape"
    Combo3.Text = "Portrait"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    tcxbarra.Enabled = True
    'tcxbarra.Show
    Unload Me

End Sub

Private Sub HScroll3_Change()
    Text3.Text = HScroll3.Value

End Sub
