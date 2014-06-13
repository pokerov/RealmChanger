VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realm Changer"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":628A
   ScaleHeight     =   6780
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Interval        =   1
      Left            =   9120
      Top             =   480
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   9120
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00000000&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   4400
      Width           =   375
   End
   Begin VB.TextBox txtRealmPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   4400
      Width           =   4335
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtRealms 
      Height          =   495
      Left            =   7320
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstRealms 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2340
      Left            =   2280
      TabIndex        =   0
      Top             =   1700
      Width           =   4815
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtRealm 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "set realmlist"
      Top             =   5000
      Width           =   4815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Specially designed for ClearType™ that runs perfectly over Microsoft® Windows® XP or Vista"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   6240
      Width           =   4815
   End
   Begin VB.Label lS 
      BackStyle       =   0  'Transparent
      Caption         =   "Changed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   3000
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "WoW realm path:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   4140
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2008 Chris FAKAR"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected realm:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   4755
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Available server realms"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim sFN As String

Private Sub cmdBrowse_Click()
    On Error Resume Next
    
    With CD
        .DialogTitle = "Open realm..."
        .Filter = "realmlist.wtf|realmlist.wtf"
        .FilterIndex = 1
        .ShowOpen
    End With
    
    txtRealmPath.Text = CD.FileName
    SaveSetting "RealmChanger", "WOWRealm", "RealmPath", txtRealmPath.Text
End Sub

Private Sub cmdChange_Click()
    On Error GoTo a
    
    With FSO
        sFN = txtRealmPath.Text
        Set strFile = .CreateTextFile(sFN, True)
        strFile.Write (txtRealm.Text)
    End With
    lS.Visible = True
    tmrHide.Enabled = True
    
a:
    If Len(Err.Description) = 0 Then Exit Sub
    MsgBox Err.Description, vbCritical, "Error"
    Exit Sub
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    
    lstRealms.Clear
    With FSO
        sFN = App.Path & "\servers.wtf"
        Set sText = .OpenTextFile(sFN, ForReading)
        With sText
            Do Until .AtEndOfStream
                sabc = (txtRealms.Text & .ReadLine)
                lstRealms.AddItem sabc
            Loop
        End With
    End With
    lstRealms.Selected(0) = True
    
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    
    With FSO
        sFN = App.Path & "\servers.wtf"
        Set sText = .OpenTextFile(sFN, ForReading)
        With sText
            Do Until .AtEndOfStream
                sabc = (txtRealms.Text & .ReadLine)
                lstRealms.AddItem sabc
            Loop
        End With
    End With
    lstRealms.Selected(0) = True
    
    txtRealmPath.Text = GetSetting("RealmChanger", "WOWRealm", "RealmPath", "C:\Program Files\World of Warcraft\realm.wrf")
    
    
    
EH:
    If Len(Err.Description) = 0 Then Exit Sub
    MsgBox Err.Description, vbCritical, "Error"
    Exit Sub
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lS.Left = (Me.ScaleWidth - lS.Width) / 2
    Label1.Left = (Me.ScaleWidth - Label1.Width) / 2
    lstRealms.Left = (Me.ScaleWidth - lstRealms.Width) / 2
    Label4.Left = lstRealms.Left
    txtRealmPath.Left = lstRealms.Left
    Label2.Left = lstRealms.Left
    txtRealm.Left = lstRealms.Left
    cmdRefresh.Left = lstRealms.Left
    cmdChange.Left = txtRealm.Left + txtRealm.Width - cmdChange.Width
    cmdBrowse.Left = txtRealmPath.Left + txtRealmPath.Width + 120
    Label3.Left = (Me.ScaleWidth - Label3.Width) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lstRealms_Click()
    txtRealm.Text = lstRealms.Text
End Sub

Private Sub tmrCheck_Timer()
    On Error Resume Next
    If Len(txtRealmPath.Text) = 0 Then
        cmdChange.Enabled = False
    Else
        cmdChange.Enabled = True
    End If
End Sub

Private Sub tmrHide_Timer()
    On Error Resume Next
    If lS.Visible = True Then
        lS.Visible = False
        tmrHide.Enabled = False
    End If
End Sub

Private Sub txtRealmPath_Change()
    On Error Resume Next
    SaveSetting "RealmChanger", "WOWRealm", "RealmPath", txtRealmPath.Text
End Sub
