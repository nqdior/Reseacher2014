VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form Form0 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "SOFIT SQL Researcher"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "メイリオ"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form0"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   15588.86
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command7 
      Caption         =   "Exit (&E)"
      Height          =   735
      Left            =   7440
      TabIndex        =   5
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search (&S)"
      Height          =   735
      Left            =   7440
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
      Height          =   4335
      Left            =   960
      OleObjectBlob   =   "MainMenu.frx":27A2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calculation (&C)"
      Height          =   735
      Left            =   7440
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton command1 
      Caption         =   "JoinMenu (&F)"
      Height          =   735
      Left            =   7440
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   6720
      TabIndex        =   8
      Top             =   600
      Width           =   3375
      Begin VB.CommandButton Command4 
         Caption         =   "Report (&G)"
         Height          =   735
         Left            =   720
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   5295
      Begin VB.CommandButton Command6 
         Caption         =   "↓"
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "↑"
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "Sort"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//--MainMenu概要説明---------------------------------------20141201--//
'
'   基本的動作
'   表示時にGridにテーブル一覧が表示され、Dblクリックでデータ表示
'
'   "SPSDATA"DBに接続時のみ帳票メニューが表示される
'
'//-------------------------------------------------------------------//



Private Sub Form_activate()
On Error GoTo SQLERR
    
'■起動時動作---------------------------------------------
    
    Form0.MousePointer = 11
        
    '▼ログインDBがSPSDATAの際帳票機能表示------------
    
    If selDB = "SPSDATA" Then
        Frame1.Height = 4575
    End If
    
    '▲-----------------------------------------------
        
    Call CNclose
    Call RSclose
    ssql = ""
    Nowform = "form0"


    '▼UIから作成したSQLをEditor及びTDBGridへ渡す------
    
    cn.Open cnstr
    ssql = "SELECT name as 'テーブル名' FROM sysobjects WHERE xtype = 'u'"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Set rs_c0 = rs.Clone
    Set TDBGrid1.DataSource = rs_c0
    TDBGrid1.Refresh
    
    '▲-----------------------------------------------
        
'■-------------------------------------------------------

    Form0.MousePointer = 0

SQLERR:
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Call CNclose
    Call RSclose
    
End Sub


Private Sub Command1_Click() 'JoinMenu移行

    Call CNclose
    Call RSclose
    ssql = ""

    Form1.Show (1)
    
End Sub


Private Sub Command2_Click() 'SearchMenu移行
    
    Call CNclose
    Call RSclose
    
    Form2.Show (1)

End Sub


Private Sub Command3_Click() 'CalcMenu移行
    
    Call CNclose
    Call RSclose
    
    Form3.Show (1)
    
End Sub


Private Sub Command4_Click() 'ReportMenu移行
    
    Call CNclose
    Call RSclose
    
    Form4.Show (1)

End Sub


Private Sub Command5_Click() 'テーブル一覧ソート機能（↑）
        
    Call RSclose

    ssql = "SELECT name as 'テーブル名' FROM sysobjects WHERE xtype = 'u' order by name"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c0 = rs.Clone
    Set TDBGrid1.DataSource = rs_c0
    TDBGrid1.Refresh
        
End Sub


Private Sub Command6_Click() 'テーブル一覧ソート機能（↓）
        
    Call RSclose

    ssql = "SELECT name as 'テーブル名' FROM sysobjects WHERE xtype = 'u' order by name desc"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c0 = rs.Clone
    Set TDBGrid1.DataSource = rs_c0
    TDBGrid1.Refresh
        
End Sub


Private Sub Command7_Click() 'JoinMenu移行

Call Unload(Me)

End Sub


Private Sub TDBgrid1_DblClick() 'TDBグリッドからForm表示
On Error GoTo SQLERR
    
'■唯一form0からEditorへ連携のあるプロシージャ■

    '▼TDBGridから値取得--------------------------
    
    retcode = nullEmpty(TDBGrid1.Columns(0).Value)
    ssql = "Select * From " & retcode
    
    '▲-------------------------------------------

    '▼作成したssqlを表示-------------------------
    
    pSql = ssql
    EForm.Text1.Text = ssql
    
    Call CNclose
    Call RSclose
    
    EForm.Show (1)
    
    '▲-------------------------------------------
    
SQLERR:
    Exit Sub
End Sub '-----------------------------------------------
