VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "DMIni Editor"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10860
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Toolbar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10860
      TabIndex        =   6
      Top             =   0
      Width           =   10860
      Begin DMIni.Button Button5 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   0
         Width           =   375
         _extentx        =   661
         _extenty        =   661
         picture         =   "Form1.frx":0442
         linecolor2      =   8421504
      End
      Begin DMIni.Button Button3 
         Height          =   375
         Left            =   860
         TabIndex        =   9
         Top             =   0
         Width           =   375
         _extentx        =   661
         _extenty        =   661
         picture         =   "Form1.frx":0554
         linecolor2      =   8421504
      End
      Begin DMIni.Button Button2 
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   0
         Width           =   375
         _extentx        =   661
         _extenty        =   661
         picture         =   "Form1.frx":08A6
         linecolor2      =   8421504
      End
      Begin DMIni.Button Button1 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   375
         _extentx        =   661
         _extenty        =   661
         picture         =   "Form1.frx":0BF8
         linecolor2      =   8421504
      End
      Begin DMIni.Button Button4 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   0
         Width           =   375
         _extentx        =   661
         _extenty        =   661
         picture         =   "Form1.frx":0F4A
         linecolor2      =   8421504
      End
   End
   Begin VB.PictureBox split 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   2880
      ScaleHeight     =   7695
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.ListBox lstSections 
      Height          =   7665
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox txtTemp 
      Height          =   1815
      Left            =   8400
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":129C
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   5040
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox mid 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   2760
      MouseIcon       =   "Form1.frx":134A
      MousePointer    =   99  'Custom
      ScaleHeight     =   7695
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   360
      Width           =   105
   End
   Begin VB.ListBox lstTemp 
      Height          =   1815
      Left            =   8400
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox txtIni 
      Height          =   7695
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   13573
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":149C
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditInsertFile 
         Caption         =   "&Insert File"
         Shortcut        =   ^{INSERT}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsRefresh 
         Caption         =   "&Update Section List"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuOptionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsFonts 
         Caption         =   "&Fonts"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SectionDeleted As Boolean
Dim LastText As String
Dim OpenedFile As String
Dim Saved As Boolean

Dim Splitter As New SplitClass

Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
Select Case ButtonIndex
    Case 1
        mnuFileNew_Click
    Case 3
        mnuOpen_Click
    Case 4
        mnuSave_Click
    Case 6
        mnuPrint_Click
    Case 8
        mnuEditCut_Click
    Case 9
        mnuEditCopy_Click
    Case 10
        mnuEditPaste_Click
End Select
End Sub

Private Sub Button1_Click()
mnuFileNew_Click
End Sub

Private Sub Button2_Click()
mnuOpen_Click
End Sub

Private Sub Button3_Click()
If OpenedFile = "" Then
        mnuSaveAs_Click
    Else
        mnuSave_Click
End If
End Sub

Private Sub Button4_Click()
mnuPrint_Click
End Sub

Private Sub Button5_Click()
mnuEditFind_Click
End Sub

Private Sub Form_Load()
Dim FF
Dim ReadString As String

FF = FreeFile

If Len(Dir(AddSlash(App.Path) & "FontSettings.dat")) > 0 Then
    Open AddSlash(App.Path) & "FontSettings.dat" For Input As FF
        
        Line Input #FF, textline
        lstSections.FontBold = textline
        
        Line Input #FF, textline
        lstSections.FontItalic = textline
        
        Line Input #FF, textline
        lstSections.FontName = textline
        
        Line Input #FF, textline
        lstSections.FontSize = textline
        
        
        Line Input #FF, textline
        txtIni.Font.Bold = textline
        
        Line Input #FF, textline
        txtIni.Font.Italic = textline
        
        Line Input #FF, textline
        txtIni.Font.Name = textline
        
        Line Input #FF, textline
        txtIni.Font.Size = textline
        
    Close FF
End If

FF = FreeFile

If Len(Dir(AddSlash(App.Path) & "RecentFiles.dat")) > 0 Then
    mnuSep4.Visible = True
    Open AddSlash(App.Path) & "RecentFiles.dat" For Input As FF 'AddSlash(App.Path) & "RecentFiles.dat" For Input As FF
        Do Until EOF(FF)
            Line Input #FF, textline
            If Len(Dir(textline)) > 0 Then
                Load mnuRecent(mnuRecent.Count)
                mnuRecent(mnuRecent.Count - 1).Caption = textline
                mnuRecent(mnuRecent.Count - 1).Visible = True
            End If
        Loop
    Close FF
Else
    mnuSep4.Visible = False
End If

If mnuRecent.Count = 1 Then mnuSep4.Visible = False

Saved = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim FF
FF = FreeFile
Open AddSlash(App.Path) & "RecentFiles.dat" For Output As FF 'AddSlash(App.Path) & "RecentFiles.dat" For Input As FF
    For i = 1 To mnuRecent.Count - 1
        Print #FF, mnuRecent(i).Caption
    Next i
Close FF

FF = FreeFile

Open AddSlash(App.Path) & "FontSettings.dat" For Output As FF
        
        Print #FF, lstSections.FontBold
        Print #FF, lstSections.FontItalic
        Print #FF, lstSections.FontName
        Print #FF, lstSections.FontSize
        
        Print #FF, txtIni.Font.Bold
        Print #FF, txtIni.Font.Italic
        Print #FF, txtIni.Font.Name
        Print #FF, txtIni.Font.Size
        

Close FF

If Saved = False Then
    If MsgBox("Do you want to quit the program without saving your changes?", vbCritical + vbYesNo, frmMain.Caption) = vbNo Then
        If OpenedFile = "" Then
            mnuSaveAs_Click
        Else
            mnuSave_Click
        End If
    End If
End If
End Sub

Private Sub Form_Resize()
ResizeControls
End Sub

Private Sub lstSections_Click()
Dim FoundPos
Dim SelectedItem As String

SelectedItem = lstSections.List(lstSections.ListIndex)
LastItem = lstSections.List(lstSections.ListCount - 1)

FoundPos = InStr(1, txtIni.Text, LastItem, vbTextCompare)
txtIni.SelStart = FoundPos - 1

FoundPos = InStr(1, txtIni.Text, SelectedItem, vbTextCompare)

If FoundPos = 0 Then Exit Sub

txtIni.SelStart = FoundPos - 1
txtIni.SelLength = Len(SelectedItem)
'result = SendMessageBynum(txtIni.hwnd, EM_LINESCROLL, 0, txtIni.GetLineFromChar(txtIni.SelStart))
'Debug.Print result
txtIni.SetFocus

'txtIni.Find lstSections.List(lstSections.ListIndex), 0, 0, rtfNoHighlight
'dl& = SendMessageBynum(Text1.hwnd, EM_LINESCROLL, 0, CLng(VScroll1.Value - firstvisible%))

End Sub

Public Function RefreshSectionList()
Dim FF
FF = FreeFile

lstTemp.Clear

txtIni.SaveFile AddSlash(App.Path) & "tmp.ini", rtfText

Open AddSlash(App.Path) & "tmp.ini" For Input As FF
    Do Until EOF(FF)
        Line Input #FF, textline
        If Left(textline, 1) = "[" Then lstTemp.AddItem textline
    Loop
Close FF

Kill AddSlash(App.Path) & "tmp.ini"

For i = 0 To lstTemp.ListCount - 1
    If Not lstSections.List(i) = lstTemp.List(i) Then
        GoTo FillSectionList
    End If
Next i
Exit Function

FillSectionList:
lstSections.Clear
For i = 0 To lstTemp.ListCount - 1
    lstSections.AddItem lstTemp.List(i)
Next i
End Function

Private Sub lstSections_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblMessage = "Here you can see all the sections in the opened ini file"
End Sub

Private Sub mid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
split.Left = txtIni.Left - (mid.Width / 2)
split.Top = lstSections.Top
split.Visible = True
End Sub

Private Sub mid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And mid.Visible = True Then
    split.Move x + txtIni.Left - 115 ', lstsections.Top, 3, lstsections.Height
End If
End Sub

Private Sub mid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
split.Visible = False
lstSections.Width = split.Left - lstSections.Left
mid.Left = split.Left
txtIni.Left = split.Left + mid.Width
txtIni.Width = (Me.ScaleWidth - txtIni.Left)

End Sub

Private Sub mnuEdit_Click()
If txtIni.SelLength = 0 Then
    mnuEditCopy.Enabled = False
    mnuEditCut.Enabled = False
    mnuEditDelete.Enabled = False
Else
    mnuEditCopy.Enabled = True
    mnuEditCut.Enabled = True
    mnuEditDelete.Enabled = True
End If
End Sub

Private Sub mnuEditAll_Click()
txtIni.SelStart = 0
txtIni.SelLength = Len(txtIni.Text)
End Sub

Private Sub mnuEditCopy_Click()
Clipboard.Clear
Clipboard.SetText txtIni.SelText
End Sub

Private Sub mnuEditCut_Click()
Clipboard.Clear
Clipboard.SetText txtIni.SelText
txtIni.SelText = ""
End Sub

Private Sub mnuEditDelete_Click()
txtIni.SelText = ""
End Sub

Private Sub mnuEditFind_Click()
frmFind.Show
End Sub

Private Sub mnuEditInsertFile_Click()
On Error GoTo Canceled

CDlg.DialogTitle = "Instert File"
CDlg.Filter = "All Files(*.*)|*.*|Ini Files(*.ini)|*.ini"
CDlg.FilterIndex = 2
CDlg.ShowOpen

txtTemp.Text = ""
txtTemp.LoadFile CDlg.filename

txtIni.SelText = txtTemp.Text


Canceled:
End Sub

Private Sub mnuEditPaste_Click()
txtIni.SelText = Clipboard.GetText
RefreshSectionList
End Sub



Private Sub mnuEditReplace_Click()
frmReplace.Show
End Sub

Private Sub mnuEditUndo_Click()
SendMessage txtIni.hwnd, EM_UNDO, 0, 0
End Sub

Private Sub mnuFileNew_Click()
If Saved = False Then
    If MsgBox("Do you want save your changes to the current document?", vbCritical + vbYesNo, frmMain.Caption) = vbYes Then
        If OpenedFile = "" Then
            mnuSaveAs_Click
        Else
            mnuSave_Click
        End If
    End If
End If

txtIni.Text = ""
lstSections.Clear
lstTemp.Clear
OpenedFile = ""
Saved = True
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
On Error GoTo Canceled
Randomize
Dim Rand As Integer

CDlg.DialogTitle = "Open"
CDlg.Filter = "All Files(*.*)|*.*|Ini Files(*.ini)|*.ini"
CDlg.FilterIndex = 2
CDlg.ShowOpen

lstSections.Clear
lstTemp.Clear
txtIni.LoadFile CDlg.filename

Saved = True

'Add item to recent file list

mnuSep4.Visible = True

RefreshSectionList

For i = 1 To mnuRecent.Count - 1
    If mnuRecent(i).Caption = CDlg.filename Then Exit Sub
Next i

If mnuRecent.Count = 5 Then
    Rand = Int((4 * Rnd) + 1)
    mnuRecent(Rand).Caption = CDlg.filename
    mnuRecent(Rand).Visible = True
Else
    Load mnuRecent(mnuRecent.Count)
    mnuRecent(mnuRecent.Count - 1).Caption = CDlg.filename
    mnuRecent(mnuRecent.Count - 1).Visible = True
End If
    


Canceled:
Exit Sub
End Sub

Private Sub mnuOptionsFonts_Click()
frmFonts.Show vbModal, Me
End Sub

Private Sub mnuOptionsRefresh_Click()
RefreshSectionList
End Sub

Private Sub mnuPrint_Click()
On Error GoTo Canceled

CDlg.DialogTitle = "Print"

CDlg.ShowPrinter

Printer.Copies = CDlg.Copies
Printer.FontBold = txtIni.Font.Bold
Printer.FontItalic = txtIni.Font.Italic
Printer.FontUnderline = txtIni.Font.Underline
Printer.FontSize = txtIni.Font.Size
Printer.FontName = txtIni.Font.Name

Printer.Print txtIni.Text
Printer.EndDoc

Canceled:
End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub mnuRecent_Click(index As Integer)
lstSections.Clear
lstTemp.Clear
txtIni.LoadFile mnuRecent(index).Caption
opened = mnuRecent(index).Caption

Saved = True

RefreshSectionList
End Sub

Private Sub mnuSave_Click()
txtIni.SaveFile OpenedFile, rtfText
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo Canceled
Randomize
Dim Rand As Integer

CDlg.DialogTitle = "Save as..."
CDlg.Filter = "All Files(*.*)|*.*|Ini File(*.ini)|*.ini"
CDlg.FilterIndex = 2
CDlg.ShowSave

OpenedFile = CDlg.filename

txtIni.SaveFile CDlg.filename, rtfText

Saved = True

'Add item to recent file list

mnuSep4.Visible = True

For i = 1 To mnuRecent.Count - 1
    If mnuRecent(i).Caption = CDlg.filename Then Exit Sub
Next i

If mnuRecent.Count = 5 Then
    Rand = Int((4 * Rnd) + 1)
    mnuRecent(Rand).Caption = CDlg.filename
    mnuRecent(Rand).Visible = True
Else
    Load mnuRecent(mnuRecent.Count)
    mnuRecent(mnuRecent.Count - 1).Caption = CDlg.filename
    mnuRecent(mnuRecent.Count - 1).Visible = True
End If
    
RefreshSectionList

Canceled:
End Sub

Private Sub txtIni_Change()
Saved = False
End Sub

Private Sub txtIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then
    If LastText = "]" Or LastText = "[" Then RefreshSectionList
End If
End Sub

Private Sub txtIni_KeyUp(KeyCode As Integer, Shift As Integer)
Dim FoundPos As Variant
Dim StringToFind As String
If KeyCode = 221 Then
    '"]" key
    RefreshSectionList
End If
End Sub
Public Function AddSlash(Path As String) As String
'On Error Resume Next
If Right(Path, 1) = "\" Or Right(Path, 1) = "/" Then
    AddSlash = Path
Else
    AddSlash = Path & "\"
End If
End Function




Public Function ResizeControls()
On Error Resume Next
lstSections.Height = Me.ScaleHeight - 300
txtIni.Width = Me.ScaleWidth - lstSections.Width - 100
txtIni.Height = Me.ScaleHeight - 300
mid.Height = txtIni.Height
split.Height = txtIni.Height
End Function
