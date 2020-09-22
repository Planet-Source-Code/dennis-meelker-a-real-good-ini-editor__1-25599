VERSION 5.00
Begin VB.Form frmFonts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fonts"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmFonts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   3480
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&Edit Field"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Sections List"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.PictureBox picEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   5745
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton btnEditSame 
         Caption         =   "&Make Section Font The Same"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   5535
      End
      Begin VB.ComboBox cmbEditFont 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   120
         Width           =   4455
      End
      Begin VB.TextBox txtEditSize 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Text            =   "8"
         Top             =   600
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Example"
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   5535
         Begin VB.Label lblEditTest 
            Caption         =   "Test"
            Height          =   975
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.CheckBox chkEditBold 
         Caption         =   "&Bold"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox chkEditItalic 
         Caption         =   "&Italic"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Font name:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Font &Size:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   705
      End
   End
   Begin VB.PictureBox picSection 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   5745
      TabIndex        =   2
      Top             =   480
      Width           =   5775
      Begin VB.CommandButton btnSectionsSame 
         Caption         =   "&Make Edit Field-Font The Same"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   5535
      End
      Begin VB.CheckBox chkSectionItalic 
         Caption         =   "&Italic"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkSectionBold 
         Caption         =   "&Bold"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Example"
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
         Begin VB.Label lblSectionsTest 
            Caption         =   "Test"
            Height          =   975
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.TextBox txtSectionSize 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "8"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox cmbSectionFont 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Font &Size:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Font name:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnEditSame_Click()
With lblSectionsTest
    .FontName = cmbEditFont.Text
    .FontSize = txtEditSize.Text
    .FontBold = chkEditBold.Value
    .FontItalic = chkEditItalic.Value
End With

cmbSectionFont.Text = cmbEditFont.Text
txtSectionSize.Text = txtEditSize.Text
chkSectionBold.Value = chkEditBold.Value
chkSectionItalic.Value = chkEditItalic.Value
End Sub

Private Sub btnSectionsSame_Click()
With lblEditTest
    .FontName = cmbSectionFont.Text
    .FontSize = txtSectionSize.Text
    .FontBold = chkSectionBold.Value
    .FontItalic = chkSectionItalic.Value
End With

cmbEditFont.Text = cmbSectionFont.Text
txtEditSize.Text = txtSectionSize.Text
chkEditBold.Value = chkSectionBold.Value
chkEditItalic.Value = chkSectionItalic.Value
End Sub

Private Sub chkEditBold_Click()
UpdateTest
End Sub

Private Sub chkEditItalic_Click()
UpdateTest
End Sub

Private Sub chkSectionBold_Click()
UpdateTest
End Sub

Private Sub chkSectionItalic_Click()
UpdateTest
End Sub

Private Sub cmbEditFont_Click()
UpdateTest
End Sub

Private Sub cmbSectionFont_Click()
UpdateTest
End Sub

Private Sub Command1_Click()
With frmMain.lstSections.Font
    .Name = lblSectionsTest.FontName
    .Size = lblSectionsTest.FontSize
    .Bold = lblSectionsTest.FontBold
    .Italic = lblSectionsTest.FontItalic
End With

With frmMain.txtIni.Font
    .Name = lblEditTest.FontName
    .Size = lblEditTest.FontSize
    .Bold = lblEditTest.FontBold
    .Italic = lblEditTest.FontItalic
End With

frmMain.ResizeControls

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
For i = 0 To Screen.FontCount - 1
    cmbSectionFont.AddItem Screen.Fonts(i)
    cmbEditFont.AddItem Screen.Fonts(i)
Next i

cmbSectionFont.Text = frmMain.lstSections.FontName
cmbEditFont.Text = frmMain.txtIni.Font.Name

txtSectionSize.Text = frmMain.lstSections.FontSize
txtEditSize.Text = frmMain.txtIni.Font.Size

chkSectionBold.Value = IIf(frmMain.lstSections.FontBold, 1, 0)
chkSectionItalic.Value = IIf(frmMain.lstSections.FontItalic, 1, 0)

chkEditBold.Value = IIf(frmMain.txtIni.Font.Bold, 1, 0)
chkEditItalic.Value = IIf(frmMain.txtIni.Font.Italic, 1, 0)

UpdateTest
End Sub

Private Sub Option1_Click()
picSection.Visible = True
picEdit.Visible = False
End Sub

Private Sub Option2_Click()
picSection.Visible = False
picEdit.Visible = True
End Sub

Private Sub txtEditSize_Change()
UpdateTest
End Sub

Private Sub txtSectionSize_Change()
UpdateTest
End Sub

Public Function UpdateTest()
On Error Resume Next
With lblSectionsTest
    .FontName = cmbSectionFont.Text
    .FontSize = txtSectionSize.Text
    .FontBold = chkSectionBold.Value
    .FontItalic = chkSectionItalic.Value
End With

With lblEditTest
    .FontName = cmbEditFont.Text
    .FontSize = txtEditSize.Text
    .FontBold = chkEditBold.Value
    .FontItalic = chkEditItalic.Value
End With

End Function
