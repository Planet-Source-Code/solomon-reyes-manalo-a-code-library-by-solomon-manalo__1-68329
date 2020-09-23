VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOLO CODE Library"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5760
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":08CA
   ScaleHeight     =   2250
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   4560
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   1140
      Width           =   1155
      Begin CodeLibrary.lvButtons_H lvOK 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   30
         TabIndex        =   2
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "OK"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockHover       =   1
         cGradient       =   16711680
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmLogin.frx":12B4
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmLogin.frx":164E
      End
      Begin CodeLibrary.lvButtons_H lvCancel 
         Height          =   375
         Left            =   30
         TabIndex        =   3
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockHover       =   1
         cGradient       =   16711680
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmLogin.frx":17B0
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmLogin.frx":1B4A
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   6
      Top             =   1140
      Width           =   4455
      Begin VB.TextBox txtpass 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   630
         Width           =   3165
      End
      Begin VB.TextBox txtname 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   0
         Top             =   240
         Width           =   3165
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   345
         Left            =   1080
         Top             =   600
         Width           =   3225
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   345
         Left            =   1080
         Top             =   210
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   930
      Width           =   5805
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   5760
      TabIndex        =   4
      Top             =   0
      Width           =   5760
      Begin VB.Image Image3 
         Height          =   420
         Left            =   300
         Picture         =   "frmLogin.frx":1CAC
         Top             =   360
         Width           =   420
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   4920
         Picture         =   "frmLogin.frx":2696
         Top             =   90
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   90
         Picture         =   "frmLogin.frx":3560
         Top             =   60
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Folder As New FolderSysObject
Dim File As New FileSysObject


Private Sub Form_Load()
If fso.FolderExists(DATAPATH) = False Then
   fso.CreateFolder DATAPATH
End If
If FindSecurityFile = False Then frmComment.Show 1, Me
End Sub

Private Sub lvCancel_Click()
Unload Me: End
End Sub

Private Sub lvOK_Click()
Dim key() As String
If txtname.Text = "" Then
   MsgBox "A Username is required.", vbCritical, "Username Error"
   Call HL(txtname)
   Exit Sub
End If
If txtpass.Text = "" Then
   MsgBox "A Password is required.", vbCritical, "Password Error"
   Call HL(txtpass)
   Exit Sub
End If

key = Split(ReadSecurityFile, "//")
If txtname.Text = crypt.DeCode(key(0)) And txtpass.Text = crypt.DeCode(key(1)) Then
   mdiMain.StatusBar1.Panels(2).Text = crypt.DeCode(key(0)) & " "
   mdiMain.Show
   Unload Me
Else
   MsgBox "Sorry, Incompatible Username and Password.", vbCritical, "Error"
   Call HL(txtname)
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lvOK_Click
End Sub


Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lvOK_Click
End Sub
