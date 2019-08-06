VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'ﾌﾗｯﾄ
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Search Menu"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "メイリオ"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4690
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   6345
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   2300
   End
   Begin VB.ComboBox Combo3 
      Height          =   420
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   420
      Left            =   2640
      TabIndex        =   4
      Text            =   "前方"
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   420
      Left            =   2640
      TabIndex        =   2
      Text            =   "="
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Text            =   " Keyword2"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   3
      Text            =   " Keyword1"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reseach (F5)"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   390
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   688
      _Version        =   393216
      Text            =   "Column Name"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   688
      _Version        =   393216
      Text            =   "Table Name "
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SavingTable"
      Height          =   2175
      Left            =   3840
      TabIndex        =   11
      Top             =   240
      Width           =   2175
      Begin VB.CommandButton Command3 
         Caption         =   "ClearTable"
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "TableCall"
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "Search Menu"
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "　～"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim oControl As Control

'ver1.1joint対応/20141121■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub Command2_Click()
On Error GoTo SQLERR
        jointflg = "1"
        If tableX = "" Then
            MsgBox ("保持されているテーブルがありません")
            Exit Sub
        End If
        
        '▼sSqlにOrderByが含まれているか判定し、存在すればOrderBy以降を切り捨て
        'ver0.9から追加
        
        tmp = InStr(1, tableX, "order")
        If tmp <> 0 Then
            tableX = Left(tableX, tmp - 2)
        End If
        'ソートによる不具合発生時はこちらのコメント削除にて対応可能。
        'その場合ソート時のSQLは別フォームから移行時のSQL+Orderとなる。
        '▲---------------------------------------------------------------------
        
        Combo3.Clear
        
    '▼rs,cn状態確認及び閉じる処理----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        If cn.State <> 0 Then
        cn.Close
        End If
    '▲--------------------------------------------------------

        cn.Open cnstr
                
        '▼DataCombo2にカラム一覧を表示する--------------------------
        ssql = tableX
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c1 = rs.Clone 'rsを使い回す為、現時点のクローンをtmpに作成
  
        For i = 0 To rs.Fields.Count - 1
            Combo3.AddItem rs.Fields(i).Name
        Next i
        
        '▲------------------------------------------------------------
                    
            '▼rs状態確認及び処理----------------------------------------
            If rs.State <> 0 Then
            rs.Close
            End If
            '▲----------------------------------------------------------
            
    Combo3.Visible = True
    Text3.Text = tableXnm
    Text3.Visible = True
    
    tableXstr = "(" & tableX & ") as " & tableXnm
SQLERR:
    Exit Sub
End Sub
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

Private Sub Command3_Click()
jointflg = "0"
tableXstr = tableX
Call Form_activate
End Sub

Public Sub Form_activate()
On Error GoTo SQLERR
    '▼rs,cn状態確認及び閉じる処理----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        If cn.State <> 0 Then
        cn.Close
        End If
    '▲--------------------------------------------------------
        Combo1.Clear: Combo1.Text = "="
        Text1.Text = " Keyword1"
        Text2.Text = " Keyword2": Text2.Visible = False
        Combo2.Clear: Combo2.Text = 前方: Combo2.Visible = False
        
        cn.Open cnstr
        
        '■ver1.1 joint対応/20141121 ■■■■■■■■■■■■■■■■■■■■■■
            If tableXnm <> "" Then
                Label3.Caption = tableXnm
            Else: Label3.Caption = "Saving Nothing"
            End If
            If Text3.Visible = True Then
                Text3.Text = ""
                Text3.Visible = False
            End If
            If Combo3.Visible = True Then
                Combo3.Clear
                Combo3.Visible = False
            End If
        
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        
        '▼DataCombo2にカラム一覧を表示する----------------------------
        If DataCombo1.Text <> "" And DataCombo1.Text <> "Table Name" Then
            ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo1.Text & "')"
            rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
            Set rs_c2 = rs.Clone
            
            '▼rs状態確認及び処理----------------------------------------
            If rs.State <> 0 Then
            rs.Close
            End If
            '▲----------------------------------------------------------
                    
            Set DataCombo2.RowSource = rs_c2
            DataCombo2.ListField = "Name"
            DataCombo2.Refresh
        End If
        
        '▲-------------------------------------------------------------
        '▼DataCombo1にテーブル一覧を表示する--------------------------
        ssql = "SELECT name FROM sysobjects WHERE xtype = 'u'"
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c1 = rs.Clone 'rsを使い回す為、現時点のクローンをtmpに作成
        
        '▼rs状態確認及び処理----------------------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        '▲----------------------------------------------------------
        
        Set DataCombo1.RowSource = rs_c1
        DataCombo1.ListField = "Name"
        DataCombo1.Refresh
        
        '▲------------------------------------------------------------
        
        '▼Combo1にアイテム追加----------------------------------------
        Combo1.AddItem ("="): Combo1.AddItem ("<"): Combo1.AddItem (">"):
        Combo1.AddItem ("<="): Combo1.AddItem (">="): Combo1.AddItem ("<>"):
        Combo1.AddItem ("bet"): Combo1.AddItem ("like")
        '▲------------------------------------------------------------
        '▼Combo2にアイテム追加----------------------------------------
        Combo2.AddItem ("前方"): Combo2.AddItem ("後方"): Combo2.AddItem ("部分"):
        '▲------------------------------------------------------------
SQLERR:
    Exit Sub
End Sub

'▼フォーム終了時-------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)

    '▼rs,cn状態確認及び閉じる処理-----------------------------
    If rs.State <> 0 Then
    rs.Close
    End If
    
    If cn.State <> 0 Then
    cn.Close
    End If
    '▲--------------------------------------------------------
    
    ssql = ""
    pSql = ""
    
End Sub
'▲----------------------------------------------------------------------


'▼F5キー押下時アクション---------------------
Private Sub Datacombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command1_Click
    End If
End Sub
Private Sub Datacombo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command1_Click
    End If
End Sub



Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command1_Click
    End If
End Sub
Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command1_Click
    End If
End Sub
'▲------------------------------------------------------------

Private Sub DataCombo1_LostFocus()
On Error GoTo SQLERR
        '▼Form2へデータ引渡し用---------------------------------------
        table_n = DataCombo1.Text
        '▲------------------------------------------------------------
         
        '▼rs状態確認及び処理----------------------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        
        '▲----------------------------------------------------------
         
        '▼DataCombo2にカラム一覧を表示する----------------------------
        ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo1.Text & "')"
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c2 = rs.Clone
        
            '▼rs状態確認及び処理----------------------------------------
            If rs.State <> 0 Then
            rs.Close
            End If
            '▲----------------------------------------------------------
                
        Set DataCombo2.RowSource = rs_c2
        DataCombo2.ListField = "Name"
        DataCombo2.Refresh
        
        '▲-------------------------------------------------------------
SQLERR:
    Exit Sub
End Sub
        '▼特殊条件時各オブジェクト表示変更----------------------------
Private Sub Combo1_LostFocus()
        If Combo1.Text = "bet" Then
            Label1.Visible = True
            Text2.Visible = True
            Combo2.Visible = False
        ElseIf Combo1.Text = "like" Then
            Label1.Visible = False
            Text2.Visible = False
            Combo2.Visible = True
        Else
            Label1.Visible = False
            Text2.Visible = False
            Combo2.Visible = False
        End If
End Sub
        '▲-------------------------------------------------------------
        
'ver1.1 joint対応/20141121■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub Command1_Click()
On Error GoTo SQLERR
        
        If Text3.Visible = True Then
            DataCombo1.Text = tableXstr
            DataCombo2.Text = Combo3.Text
        End If
        
        '▼検索条件を指定するか否か-------------------------------------
        If Combo1.Text = "null" Then
                ssql = "Select * From " & DataCombo1.Text & " where " & DataCombo2.Text & " is null"
        ElseIf Text1.Text = "" Or Text1.Text = " KeyWord1" Or DataCombo2.Text = "Column Name" Or DataCombo2.Text = "" Then
'        '条件文もしくはカラム選択がなされていない場合テーブルを表示するだけのSQL発行
                ssql = "Select * From " & DataCombo1.Text
        ElseIf Combo1.Text = "like" Then
            If Combo2.Text = "前方" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '" & Text1.Text & "%'"
            ElseIf Combo2.Text = "後方" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '%" & Text1.Text & "'"
            ElseIf Combo2.Text = "部分" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '%" & Text1.Text & "%'"
            End If
        ElseIf Combo1.Text <> "=" And Combo1.Text <> "bet" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '" & Text1.Text & "'"
        ElseIf Combo1.Text = "=" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " = '" & Text1.Text & "'"
        ElseIf Combo1.Text = "bet" Then
                ssql = "Select * From " & DataCombo1.Text & " where " & DataCombo2.Text & " between '" & Text1.Text & "' and '" & Text2.Text & "'"
        End If
        
        '▲-------------------------------------------------------------
        '▼rs状態確認及び処理----------------------------------------
        If rs.State <> 0 Then
        rs.Close
                                                                                                                                                                                         End If
        '▲----------------------------------------------------------
        '▼UIから作成したSQLをForm2及びTDBGridへ渡す--------------------
        jointflg = "0"
        pSql = ssql '受け渡し用
        Form2.Text1.Text = ssql
        Nowform = "form1"
        Form2.Show (1)
        '▲-------------------------------------------------------------


'Private Sub Command1_Click()
'On Error GoTo SQLERR
'
'        '▼検索条件を指定するか否か-------------------------------------
'        If Text1.Text = "" Or Text1.Text = " KeyWord1" Or DataCombo2.Text = "Column Name" Or DataCombo2.Text = "" Then
'        '条件文もしくはカラム選択がなされていない場合テーブルを表示するだけのSQL発行
'                ssql = "Select * From " & DataCombo1.Text
'        ElseIf Combo1.Text = "like" Then
'            If Combo2.Text = "前方" Then
'                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '" & Text1.Text & "%'"
'            ElseIf Combo2.Text = "後方" Then
'                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '%" & Text1.Text & "'"
'            ElseIf Combo2.Text = "部分" Then
'                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '%" & Text1.Text & "%'"
'            End If
'        ElseIf Combo1.Text <> "=" And Combo1.Text <> "bet" Then
'                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '" & Text1.Text & "'"
'        ElseIf Combo1.Text = "=" Then
'                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " = '" & Text1.Text & "'"
'        ElseIf Combo1.Text = "bet" Then
'                ssql = "Select * From " & DataCombo1.Text & " where " & DataCombo2.Text & " between '" & Text1.Text & "' and '" & Text2.Text & "'"
'        End If
'
'        '▲-------------------------------------------------------------
'        '▼rs状態確認及び処理----------------------------------------
'        If rs.State <> 0 Then
'        rs.Close
'        End If
'        '▲----------------------------------------------------------
'        '▼UIから作成したSQLをForm2及びTDBGridへ渡す--------------------
'        pSql = ssql '受け渡し用
'        Form2.Text1.Text = ssql
'        Nowform = "form1"
'        Form2.Show (1)
'        '▲-------------------------------------------------------------
'SQLERR:
'    Exit Sub
'End Sub

SQLERR:
    Exit Sub
End Sub
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■


'▼各Textboxに関わる挙動--------------------------------------------------------
        '▼Text1に関わる挙動----------------------------------------------------
Private Sub text1_GotFocus()
        If Text1.Text = " Keyword1" Then
            Text1.Text = ""
        End If
End Sub
        '▲----------------------------------------------------------------------

        '▼Text2に関わる挙動----------------------------------------------------
Private Sub Text2_Gotfocus()
        If Text2.Text = " Keyword2" Then
            Text2.Text = ""
        End If
End Sub
        '▲----------------------------------------------------------------------

