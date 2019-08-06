VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form StartingForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'なし
   Caption         =   "Start Menu"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "メイリオ"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   5250
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   1080
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton Command3 
         Caption         =   "Main Menu"
         Height          =   615
         Left            =   2400
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   390
         Left            =   3300
         TabIndex        =   12
         Top             =   660
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Select DB"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "メイリオ"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   960
         TabIndex        =   14
         Top             =   480
         Width           =   4215
         Begin VB.Label Label6 
            BackStyle       =   0  '透明
            Caption         =   "Connection DataBase is"
            Height          =   735
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3855
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'なし
      Height          =   1575
      Left            =   600
      Picture         =   "StartingForm.frx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   8175
      TabIndex        =   10
      Top             =   480
      Width           =   8175
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H008080FF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      MaskColor       =   &H000000FF&
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1200
      TabIndex        =   8
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox Text3 
         Appearance      =   0  'ﾌﾗｯﾄ
         BorderStyle     =   0  'なし
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2400
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         Caption         =   "Now Instance is"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Caption         =   "LOGIN"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   405
      IMEMode         =   3  'ｵﾌ固定
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3840
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "_"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   8880
      TabIndex        =   18
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Presented by Nihon Software Knowledge Corp"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "Ver 1.2.0"
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "Pass"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "ID"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "SQL Server LOGIN"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "StartingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_activate()

'インスタンス情報取得---------------------------------------

    instance = CreateObject("WScript.Network").ComputerName
    Text3.Text = instance
    
'-----------------------------------------------------------
    
End Sub


Private Sub Command1_Click() 'ログインボタン押下時----------
    On Error GoTo SQLERR

'■IDパスワード入力確認-------------------------------------
    
    If Text1.Text = "" And Text2.Text = "" Then
    MsgBox "Error: ID、パスワードが入力されていません。"
    Exit Sub
    End If
    
    Lid = StartingForm.Text1.Text
    Lpass = StartingForm.Text2.Text
'■---------------------------------------------------------
    
    
'■Datacombo1表示設定---------------------------------------
    
    '▼ログイン情報をフォームから取得----------------
    cn.Open "Provider=SQLOLEDB;Data Source=" & Text3.Text & ";" & "Initial Catalog=master; User ID=" & Lid & ";Password=" & Lpass & ";"
    
    ssql = "select name from sysdatabases"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c6 = rs.Clone 'rsを使い回す為、現時点のクローンをtmpに作成
    
    '▲----------------------------------------------
    
    '▼ログイン成功時：挙動設定----------------------
    
    Set DataCombo1.RowSource = rs_c6
    DataCombo1.ListField = "Name"
    DataCombo1.Refresh
    
    Call RSclose
    Frame1.Visible = True
    DataCombo1.SetFocus
    
    '▲----------------------------------------------
    
'■---------------------------------------------------------
    
SQLERR:
    If cn.State = 0 Then
        MsgBox "Error: ID、パスワードを確認してください｡ "
    End If
        MsgBox "インスタンスへの接続に失敗しました。"
    Exit Sub
End Sub '---------------------------------------------------



Private Sub Command2_Click() '×ボタン押下時----------------
    Call Unload(Me)
End Sub '---------------------------------------------------



Private Sub Command3_Click() 'MainMenuボタン押下時----------
    
    If DataCombo1.Text = "" Or DataCombo1.Text = "Select DB" Then
        
        MsgBox "Error: DBの認識に失敗しました。"
        
    Else '▼DB接続成功時----------------------------
        
        selDB = StartingForm.DataCombo1.Text
        cnstr = "Driver={SQL Server};SERVER=" & Text3.Text & ";" & "DATABASE=" & selDB & ";UID=" & Lid & ";PWD=" & Lpass & ";"

        Call Unload(Me)
        Form0.Show
        
    End If
End Sub '----------------------------------------------------


Private Sub Form_Unload(Cancel As Integer)
    Call CNclose
    Call RSclose
End Sub


Private Sub Label8_Click() '■■■Administratorデバッグ用■■■
    selDB = StartingForm.DataCombo1.Text
'    cnstr = "Driver={SQL Server};SERVER=" & "UGCH-003-4" & ";" & "DATABASE=" & "SPSDATA" & ";UID=" & "sa" & ";PWD=" & "5473nsk0036" & ";"
'   20141127 oledb対応
    cnstr = "Provider=SQLOLEDB;Data Source=" & "UGCH-003-4" & ";" & "Initial Catalog=" & "SPSDATA" & "; User ID=" & "sa" & ";Password=" & "5473nsk0036" & ";"
    Call Unload(Me)
    Form0.Show
End Sub '■■■■■■Administratorデバッグ用ここまで■■■■■■
