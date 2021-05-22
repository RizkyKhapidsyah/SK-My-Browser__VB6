VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form From1 
   BackColor       =   &H00000000&
   Caption         =   "My Browser"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   Icon            =   "browser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComCtl2.Animation Animation2 
      Height          =   1455
      Left            =   7680
      TabIndex        =   11
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2566
      _Version        =   327681
      BackColor       =   0
      FullWidth       =   273
      FullHeight      =   97
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   615
      Left            =   3000
      TabIndex        =   10
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1085
      _Version        =   327681
      AutoPlay        =   -1  'True
      BackColor       =   0
      FullWidth       =   297
      FullHeight      =   41
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   8400
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2940
      Visible         =   0   'False
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2593
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "22/05/2021"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "8:09"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Go Search"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go Home"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go Forward"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Go Back"
      Height          =   375
      Left            =   240
      MaskColor       =   &H0000C0C0&
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   10695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   11895
      ExtentX         =   20981
      ExtentY         =   10186
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save&As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "From1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveflag As Integer


Private Sub Command1_Click()
WebBrowser1.GoBack
End Sub

Private Sub Command2_Click()
WebBrowser1.GoForward
End Sub

Private Sub Command3_Click()
WebBrowser1.GoHome
End Sub

Private Sub Command4_Click()
WebBrowser1.GoSearch
End Sub

Private Sub Command5_Click()
WebBrowser1.Refresh

End Sub

Private Sub Form_Load()
WebBrowser1.MenuBar = True
Animation1.Open (App.Path + "\mybrowser.avi")
Animation2.Open (App.Path + "\working.avi")

Animation2.Visible = False
'from1.WindowState=
'WebBrowser1.Resizable = True
WebBrowser1.RegisterAsBrowser = True
WebBrowser1.CausesValidation = True
'From1.Icon = (App.Path + "\butterfly.ico")


End Sub

Private Sub Form_Resize()
   'owser1WebBrowser1.Width = Me.ScaleWidth - (WebBrowser1.Left * 2)
   'WebBrowser1.Height = Me.ScaleHeight - WebBrowser1.Top - WebBrowser1.Left
   'txtAddress.Width = WebBr.Width - txtAddress.Left
End Sub

Private Sub mnuabout_Click()
MsgBox "Hi,Friend.This Browser is made by Indrait Adhya.Contact him at indra_to_u@yahoo.com"


End Sub

Private Sub mnuClose_Click()
Unload Me

End Sub

Private Sub mnuMax_Click()
From1.Width = Screen.Width
From1.Height = Screen.Height










End Sub

Private Sub mnuNormal_Click()

End Sub

Private Sub mnuOpen_Click()

 Dim sFile As String


    With CommonDialog1
        
        'set the flags and attributes of the
        'common dialog control
        '.Filter = "All files(*.*)|(*.*)"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        Text1 = sFile
        WebBrowser1.Navigate CommonDialog1.FileName
        
        
    End With
    WebBrowser1.AddressBar = True
    
    
    'To Do
    'process the opened file
End Sub

Private Sub mnuprint_Click()
On Error GoTo error:
CommonDialog1.ShowPrinter
Exit Sub

error:

End Sub

Private Sub mnuSaveAs_Click()
CommonDialog1.ShowSave
Dim filenum As Integer
   CommonDialog1.Action = 2
   If CommonDialog1.FileName <> "" Then
      Screen.MousePointer = 11
      filenum = FreeFile
      Open CommonDialog1.FileName For Output As filenum
      
        Print #filenum, 'editor.TextRTF
        'editor.SetFocus
      
      Close (filenum)
      saveflag = 0
      Screen.MousePointer = 0
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      WebBrowser1.Navigate Text1.Text
End If

End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If (InStr(1, URL, "xxx") <> 0) Then
MsgBox "Access denied to URL: " & URL
Cancel = True
End If
'If WebBrowser1.Offline = False Then
  ' Animation1.Play
'Else
   'Animation1.Stop
'End If
If WebBrowser1.Busy = True Then
   Animation2.Visible = True
   Animation2.Play
   
Else
   Animation2.Visible = False
   Animation2.Stop
End If

End Sub

Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)

   Const CSC_NAVIGATEFORWARD = 1
   Const CSC_NAVIGATEBACK = 2

   Select Case Command
      Case Is = CSC_NAVIGATEFORWARD
         Command2.Enabled = Enable
      Case Is = CSC_NAVIGATEBACK
         Command1.Enabled = Enable
         
   End Select



End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Label3.Caption = 0
'Label3.Caption = (Progress * 100) / ProgressMax
'If Label3.Caption = 0 Then
'ProgressBar1.Enabled = True
'ProgressBar1.Scrolling = ccScrollingSmooth
'End If
'If Label3.Caption = 100 Then
'ProgressBar1.Enabled = False
'End If


'DoEvents
If WebBrowser1.Busy = True Then
   Animation2.Visible = True
   Animation2.Play
   
Else
   Animation2.Visible = False
   Animation2.Stop
End If

WebBrowser1.RegisterAsBrowser = True


End Sub

