VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contact me"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3600
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   1485
         Left            =   1200
         Picture         =   "frmAbout.frx":000C
         Top             =   1080
         Width           =   2265
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "DMIni Editor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1245
         TabIndex        =   3
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Version 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1755
         TabIndex        =   2
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "© 2001, DMSoftware"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()
ShellExecute hwnd, "open", "mailto:meelkertje@hotmail.com", vbNull, vbNull, SW_SHOWNORMAL
End Sub

