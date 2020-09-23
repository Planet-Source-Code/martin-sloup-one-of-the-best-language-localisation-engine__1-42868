VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "(language test)"
   ClientHeight    =   4575
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "(cancel)"
      Height          =   375
      Index           =   1
      Left            =   2940
      TabIndex        =   6
      Top             =   4080
      Width           =   1635
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "(ok)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test: (this text is not modifed)"
      Height          =   2235
      Left            =   1920
      TabIndex        =   3
      Top             =   1740
      Width           =   2655
      Begin VB.ListBox lstTestList 
         Height          =   1230
         ItemData        =   "frmTest.frx":0000
         Left            =   60
         List            =   "frmTest.frx":0010
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblT3 
         Caption         =   "Label1"
         Height          =   495
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Caption         =   "(test label)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.ListBox lstLang 
      Height          =   2205
      Left            =   60
      TabIndex        =   1
      Top             =   1740
      Width           =   1755
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   60
      ScaleHeight     =   1275
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
   Begin VB.Label lblLang 
      Caption         =   "Language: (this text is not modifed)"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1500
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "(file)"
      Begin VB.Menu mnuFile_New 
         Caption         =   "(new)"
      End
      Begin VB.Menu mnuFile_Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Open 
         Caption         =   "(open)"
      End
      Begin VB.Menu mnuFile_Line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Save 
         Caption         =   "(save)"
      End
      Begin VB.Menu mnuFile_SaveAs 
         Caption         =   "(save as)"
      End
      Begin VB.Menu mnuFile_Line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Option 
         Caption         =   "(preferences)"
      End
      Begin VB.Menu mnuFile_Line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "(exit)"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Array for Languages Dll, which is loaded in lstLang
'This array include full path and full name of language dll
Dim arrLanguages() As String

Private Sub cmdAction_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    'This Read all files and from each lang file read info varible
    'Language_Name and add this to the lstLang
    arrLanguages = LoadAllLangInfo(AppDir & "language\", lstLang)
    'Select first language in lstLang if it's possible
    If lstLang.ListCount > 0 Then
        lstLang.ListIndex = 0
        'Load selected file
        LoadLang arrLanguages(lstLang.ListIndex)
        'Apply language file on all item in a form including form
        SetFormLang Me
        'Load picture from language file!!!
        Set pic.Picture = LoadLangPicture("pic.jpg", resJPEG)
        'Set varibles to lblT3
        lblT3.Caption = GetLangVar(Me, "Test3", CStr((6 / 2)))
    End If
End Sub

Private Sub lstLang_Click()
    'Load selected file
    LoadLang arrLanguages(lstLang.ListIndex)
    'Apply language file on all item in a form including form
    SetFormLang Me
    'Load picture from language file!!!
    Set pic.Picture = LoadLangPicture("pic.jpg", resJPEG)
    'Set varibles to lblT3
    lblT3.Caption = GetLangVar(Me, "Test3", CStr((6 / 2)))
End Sub

Function AppDir() As String
    If Len(App.Path) = 3 Then
        AppDir = App.Path
    Else
        AppDir = App.Path & "\"
    End If
End Function

Private Sub mnuFile_Exit_Click()
    Unload Me
End Sub
