VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Caption         =   "Video Downloader "
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   6975
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   6300
      TabIndex        =   19
      Top             =   90
      Width           =   600
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   60
      TabIndex        =   18
      Top             =   870
      Width           =   6825
   End
   Begin VB.TextBox txtClip 
      Height          =   285
      Left            =   5670
      TabIndex        =   17
      Top             =   4410
      Visible         =   0   'False
      Width           =   675
   End
   Begin ComctlLib.StatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   5160
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7937
            MinWidth        =   7937
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Downloading Information"
      Height          =   2115
      Left            =   60
      TabIndex        =   5
      Top             =   2940
      Width           =   4245
      Begin ComctlLib.ProgressBar PB1 
         Height          =   285
         Left            =   1290
         TabIndex        =   14
         Top             =   1680
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Progress"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label lblPercent 
         Caption         =   "???"
         Height          =   255
         Left            =   1290
         TabIndex        =   13
         Top             =   1350
         Width           =   2805
      End
      Begin VB.Label Label123 
         Caption         =   "Percent"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1380
         Width           =   1005
      End
      Begin VB.Label lblSaved 
         Caption         =   "???"
         Height          =   255
         Left            =   1290
         TabIndex        =   11
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Saved"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label lblRemaining 
         Caption         =   "???"
         Height          =   255
         Left            =   1290
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Remaining"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lblVidSize 
         Caption         =   "???"
         Height          =   255
         Left            =   1290
         TabIndex        =   7
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Video Size"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   390
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Download Video"
      Height          =   2085
      Left            =   4365
      TabIndex        =   4
      Top             =   2970
      Width           =   2565
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   7605
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7035
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   120
      Width           =   5025
   End
   Begin VB.Label lblVidName 
      Caption         =   "???"
      Height          =   255
      Left            =   1290
      TabIndex        =   3
      Top             =   540
      Width           =   5595
   End
   Begin VB.Label Label2 
      Caption         =   "Video Name"
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Youtube URL"
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   1035
   End
   Begin VB.Menu mnu 
      Caption         =   "&popup"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu sap1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ResetControls()
    txtURL.Text = ""
    lblVidName.Caption = ""
    lblVidID = ""
    lblVidSize = ""
    lblRemaining = ""
    lblSaved = ""
    lblPercent = ""
    lblSpeed = ""
    PB1.Value = 0
End Sub

Private Sub cmdDownload_Click()
Dim i As Integer
    For i = 0 To List1.ListCount - 1
        DownloadVideo GetVideoInfo(List1.List(i), Inet1), VideoName & ".flv"
    Next i
End Sub

Private Sub Command1_Click()
Dim str1 As String
    If InStr(txtURL.Text, "youtube.com/watch") Then
        str1 = Left(txtURL.Text, 42)
        List1.AddItem str1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnu
    End If
End Sub

Private Sub mnuClear_Click()
    List1.Clear
End Sub

Private Sub mnuPaste_Click()
    Dim strClipData As String
    
    strClipData = Clipboard.GetText(vbCFText)
    
    If InStr(strClipData, "youtube.com/watch") Then
        txtClip.Text = Clipboard.GetText(vbCFText)
        txtURL.Text = Left(txtClip.Text, 42)
        List1.AddItem txtURL.Text
    End If
End Sub

Private Sub txtURL_Change()
    txtClip.Text = Clipboard.GetText(vbCFText)
    txtURL.Text = Left(txtClip.Text, 42)
End Sub

