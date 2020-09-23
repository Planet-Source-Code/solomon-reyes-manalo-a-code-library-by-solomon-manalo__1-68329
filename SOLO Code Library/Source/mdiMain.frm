VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00404040&
   Caption         =   "SOLO Code Library"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   10410
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   210
      Top             =   1680
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   6330
      ScaleHeight     =   7575
      ScaleWidth      =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   4080
      Begin VB.Frame fraMid 
         Height          =   7035
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   3975
         Begin MSComctlLib.ImageList imlTreview 
            Left            =   1320
            Top             =   2760
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":08CA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":0C64
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":0FFE
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   2070
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":1398
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageCombo cboLanguage 
            Height          =   360
            Left            =   60
            TabIndex        =   5
            Top             =   210
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   0
            BackColor       =   16777215
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "mdiMain.frx":1732
            Locked          =   -1  'True
            Text            =   "Language List"
            ImageList       =   "ImageList1"
         End
         Begin MSComctlLib.TreeView tvLibrary 
            Height          =   6345
            Left            =   60
            TabIndex        =   4
            Top             =   630
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   11192
            _Version        =   393217
            Indentation     =   617
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imlTreview"
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "mdiMain.frx":1894
         End
      End
      Begin VB.Frame fraTop 
         Height          =   435
         Left            =   60
         TabIndex        =   2
         Top             =   -90
         Width           =   3975
      End
   End
   Begin CodeLibrary.McToolBar mctoolMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   16
      ButtonsWidth    =   40
      ButtonsHeight   =   40
      ButtonsPerRow   =   20
      HoverColor      =   16744576
      BackGradient    =   3
      ButtonsMode     =   4
      ButtonsBackColor=   14807794
      ButtonsPerRow_Chev=   20
      ButtonToolTipIcon1=   1
      Button_Type1    =   1
      ButtonCaption2  =   ""
      ButtonIcon2     =   "mdiMain.frx":1BAE
      ButtonToolTipText2=   "New Code File (Empty)"
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   ""
      ButtonIcon3     =   "mdiMain.frx":25A8
      ButtonToolTipText3=   "Language Wizard"
      ButtonToolTipIcon3=   1
      ButtonCaption4  =   ""
      ButtonIcon4     =   "mdiMain.frx":2FA2
      ButtonToolTipText4=   "Category Wizard"
      ButtonToolTipIcon4=   1
      ButtonToolTipIcon5=   1
      Button_Type5    =   1
      ButtonCaption6  =   ""
      ButtonIcon6     =   "mdiMain.frx":387C
      ButtonToolTipText6=   "Save All"
      ButtonToolTipIcon6=   1
      ButtonCaption7  =   ""
      ButtonIcon7     =   "mdiMain.frx":4276
      ButtonToolTipText7=   "Save"
      ButtonToolTipIcon7=   1
      ButtonToolTipIcon8=   1
      Button_Type8    =   1
      ButtonCaption9  =   ""
      ButtonIcon9     =   "mdiMain.frx":4C70
      ButtonToolTipText9=   "Select ALL"
      ButtonToolTipIcon9=   1
      ButtonCaption10 =   ""
      ButtonIcon10    =   "mdiMain.frx":566A
      ButtonToolTipText10=   "Copy"
      ButtonToolTipIcon10=   1
      ButtonCaption11 =   ""
      ButtonIcon11    =   "mdiMain.frx":6064
      ButtonToolTipText11=   "Cut"
      ButtonToolTipIcon11=   1
      ButtonCaption12 =   ""
      ButtonIcon12    =   "mdiMain.frx":6A5E
      ButtonToolTipText12=   "Paste"
      ButtonToolTipIcon12=   1
      ButtonToolTipIcon13=   1
      Button_Type13   =   1
      ButtonCaption14 =   ""
      ButtonIcon14    =   "mdiMain.frx":7458
      ButtonToolTipText14=   "Security Settings"
      ButtonToolTipIcon14=   1
      ButtonCaption15 =   ""
      ButtonIcon15    =   "mdiMain.frx":7E52
      ButtonToolTipText15=   "About this Software"
      ButtonToolTipIcon15=   1
      ButtonCaption16 =   ""
      ButtonIcon16    =   "mdiMain.frx":884C
      ButtonToolTipText16=   "Exit Application"
      ButtonToolTipIcon16=   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   8175
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   882
            Text            =   "User:"
            TextSave        =   "User:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1217
            MinWidth        =   882
            Text            =   "DBUser"
            TextSave        =   "DBUser"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   2858
            MinWidth        =   882
            Text            =   "Active Document:  "
            TextSave        =   "Active Document:  "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12789
            MinWidth        =   6174
            Text            =   "Path & Name"
            TextSave        =   "Path & Name"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save As.."
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "Edit"
      Begin VB.Menu mSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mOther 
      Caption         =   "Other"
      Begin VB.Menu mLW 
         Caption         =   "Language Wizard"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mCW 
         Caption         =   "Category Wizard"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mSS 
         Caption         =   "Security Settings"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Folder As New FolderSysObject
Dim File As New FileSysObject

Private Sub cboLanguage_Change()
On Error Resume Next
Call GetCategories(Me.cboLanguage.SelectedItem.Text)
Call GetCategories(Me.cboLanguage.Text)
End Sub

Private Sub cboLanguage_Click()
On Error Resume Next
Call GetCategories(Me.cboLanguage.SelectedItem.Text)
Call GetCategories(Me.cboLanguage.Text)
End Sub

Private Sub mAbout_Click()
Call mctoolMain_Click(15)
End Sub

Private Sub mCopy_Click()
Call mctoolMain_Click(10)
End Sub

Private Sub mctoolMain_Click(ByVal ButtonIndex As Long)
Select Case ButtonIndex
       Case 2:
               Call NewDoc
       Case 3:
               frmLanguage.Show 1
       Case 4:
               frmCategory.Show 1
       Case 7:
               frmSaveCode.Show 1
       Case 6:
               If (ActiveForm Is Nothing) Then
                  Exit Sub
               Else
                  'if nofile name means NEW
                  If IsNoFileName(Me.ActiveForm.Caption) = True Then
                     frmSaveCode.txtname.Text = Me.ActiveForm.Caption
                     frmSaveCode.Show 1
                  Else
                     If Me.ActiveForm.txtdoc.bChange = True Then
                        If SaveFileDoc(DATAPATH, Me.ActiveForm.pathf, _
                            Me.ActiveForm.txtdoc.TextRTF) = False Then
                            MsgBox "Unable to save File.", vbCritical, "Unhandled File I/O Error"
                        End If
                        Me.ActiveForm.txtdoc.bChange = False
                      End If
                  End If
               End If
       Case 9:
               If (ActiveForm Is Nothing) Then
                  Exit Sub
               Else
                  Me.ActiveForm.txtdoc.SelectALL
               End If
       Case 10:
               If (ActiveForm Is Nothing) Then
                  Exit Sub
               Else
                  Me.ActiveForm.txtdoc.Copy
               End If
       Case 11:
               If (ActiveForm Is Nothing) Or Me.ActiveForm.txtdoc.SelText = "" Then
                  Exit Sub
               Else
                  Me.ActiveForm.txtdoc.Cut
               End If
        Case 12:
               If (ActiveForm Is Nothing) Then
                  Exit Sub
               Else
                  Me.ActiveForm.txtdoc.Paste
               End If
        Case 14:
               frmsecurity.Show 1
        Case 15:
               about.Show 1
        Case 16:
               End
       
End Select
End Sub

Private Sub mCut_Click()
Call mctoolMain_Click(11)
End Sub

Private Sub mCW_Click()
Call mctoolMain_Click(4)
End Sub

Private Sub MDIForm_Load()
GetLanguages
Call Timer1_Timer
End Sub

Private Sub mExit_Click()
Call mctoolMain_Click(16)
End Sub

Private Sub mLW_Click()
Call mctoolMain_Click(3)
End Sub

Private Sub mNew_Click()
Call mctoolMain_Click(2)
End Sub

Private Sub mPaste_Click()
Call mctoolMain_Click(12)
End Sub

Private Sub mSave_Click()
Call mctoolMain_Click(6)
End Sub

Private Sub mSaveAs_Click()
Call mctoolMain_Click(7)
End Sub

Private Sub mSelectAll_Click()
Call mctoolMain_Click(9)
End Sub

Private Sub mSS_Click()
Call mctoolMain_Click(14)
End Sub

Private Sub Picture1_Resize()
On Error Resume Next
    Me.fraMid.Width = Me.fraTop.Width
    Me.fraMid.Height = Me.Picture1.Height - 380
    Me.tvLibrary.Height = Me.fraMid.Height - 700
End Sub

Public Sub GetLanguages()
Dim Languages() As String
Dim keyx As String
Dim I As Integer

On Error Resume Next
Languages = Folder.GetDirectories(DATAPATH)
With Me.cboLanguage.ComboItems
     .Clear
     For I = 1 To UBound(Languages)
         .Add , Languages(I), Languages(I), 1
     Next
End With
Me.tvLibrary.Nodes(1).Expanded = True
End Sub

Public Sub GetCategories(LanguageName As String)
Dim Categories() As String
Dim keys As String
Dim I As Integer

Categories = Folder.GetDirectories(DATAPATH & LanguageName & "\")
With Me.tvLibrary.Nodes
     .Clear
     'Create Root Node
     Set itn = .Add(, , "root", "Console Root", 1)
     Set itn = Nothing
     
     For I = 1 To UBound(Categories)
         keys = LanguageName & Categories(I)
         Set itn = .Add("root", tvwChild, keys, Categories(I), 2)
         Call GetFiles(LanguageName, Categories(I))
     Next
     Set itn = Nothing
End With
Me.tvLibrary.Nodes(1).Expanded = True
End Sub

Public Sub GetFiles(LanguageName, Category As String)
Dim cFiles() As String
Dim keyx As String
Dim I As Integer

cFiles = File.GetFiles(DATAPATH & LanguageName & "\" & Category & "\")
On Error Resume Next
Set itn = Nothing
With Me.tvLibrary.Nodes

     For I = 1 To UBound(cFiles)
         keyx = LanguageName & Category & cFiles(I)
         Set itn = .Add(LanguageName & Category, tvwChild, keyx, cFiles(I), 3)
     Next
     Set itn = Nothing
     
End With
End Sub

Private Sub Timer1_Timer()
'toolbars
     If (Me.ActiveForm Is Nothing) Then
        Me.mctoolMain.SetButtonValue 6, BTN_Enabled, False
        Me.mctoolMain.SetButtonValue 7, BTN_Enabled, False
        Me.mctoolMain.SetButtonValue 9, BTN_Enabled, False
        Me.mctoolMain.SetButtonValue 10, BTN_Enabled, False
        Me.mctoolMain.SetButtonValue 11, BTN_Enabled, False
        Me.mctoolMain.SetButtonValue 12, BTN_Enabled, False
        Me.mSave.Enabled = False
        Me.mSaveAs.Enabled = False
        Me.mSelectAll.Enabled = False
        Me.mCut.Enabled = False
        Me.mCopy.Enabled = False
        Me.mPaste.Enabled = False
        Me.StatusBar1.Panels(4).Text = ""
     Else
        Me.StatusBar1.Panels(4).Text = ""
            If Me.ActiveForm.pathf = "" Then
                Me.StatusBar1.Panels(4).Text = Me.ActiveForm.Caption
            Else
                Me.StatusBar1.Panels(4).Text = Me.ActiveForm.pathf
            End If
        Me.mctoolMain.SetButtonValue 6, BTN_Enabled, Me.ActiveForm.txtdoc.bChange
        Me.mctoolMain.SetButtonValue 7, BTN_Enabled, Me.ActiveForm.txtdoc.bChange
        Me.mSave.Enabled = True
        Me.mSaveAs.Enabled = True
        If Me.ActiveForm.txtdoc.Text = "" Then
            Me.mctoolMain.SetButtonValue 9, BTN_Enabled, False
            Me.mctoolMain.SetButtonValue 10, BTN_Enabled, False
            Me.mctoolMain.SetButtonValue 11, BTN_Enabled, False
            Me.mSelectAll.Enabled = False
            Me.mCut.Enabled = False
            Me.mCopy.Enabled = False
        Else
            Me.mctoolMain.SetButtonValue 9, BTN_Enabled, True
            Me.mctoolMain.SetButtonValue 10, BTN_Enabled, True
            Me.mctoolMain.SetButtonValue 11, BTN_Enabled, True
            Me.mSelectAll.Enabled = True
            Me.mCut.Enabled = True
            Me.mCopy.Enabled = True
        End If
        Me.mctoolMain.SetButtonValue 12, BTN_Enabled, True
        Me.mPaste.Enabled = True
     End If
End Sub

Private Sub tvLibrary_Click()
Me.Timer1.Enabled = False
End Sub

Private Sub tvLibrary_DblClick()
Dim tmp As String

tmp = Me.StatusBar1.Panels(4).Text
If tmp = Empty Or InStr(tmp, fExt) = 0 Then
   Exit Sub
End If
If OpenFileDoc(DATAPATH, tmp) = False Then
   MsgBox "File Not Found.", vbCritical, "Error"
End If
End Sub

Private Sub tvLibrary_NodeClick(ByVal Node As MSComctlLib.Node)
Me.StatusBar1.Panels(4).Text = Replace(Node.FullPath, "Console Root", Me.cboLanguage.Text)
End Sub

Public Sub refreshTreeview()
Call cboLanguage_Click
Call cboLanguage_Change
End Sub
