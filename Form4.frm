VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Report Menu"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "メイリオ"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6345
   StartUpPosition =   2  '画面の中央
   Begin MSMask.MaskEdBox text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/MM/dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1041
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "メイリオ"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####/##/##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reseach (F5)"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin MSMask.MaskEdBox text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/MM/dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1041
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "メイリオ"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "Report Menu"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "EndDate"
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "StartDate"
      BeginProperty Font 
         Name            =   "メイリオ"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim table_nm As String

Private Sub Command1_Click()
    
    Dim startd As String
    Dim endd As String
    
    tablenm = "uriage" 'テーブル名
    
    If Text3.Text = "" Then '
        Text3.Text = Text2.Text
    ElseIf Text3.Text < Text2.Text Then
        MsgBox "指定範囲に誤りがあります"
        Text2.SetFocus
    End If

        ssql = "select * from " & tablenm & " "

    If Len(Text2.Text) <= 0 Then '集計有∧期間指定無
        MsgBox "範囲が指定されていません"
        Text2.SetFocus
        Exit Sub
    Else '集計有∧期間指定有

        ssql1 = "select syoku.syokuto + isnull(buy.buyto,0) as [総合計],syoku.sya_cd as [社員ID] ,syoku.sya_nm as [社員名] ,syoku.syokuto as [食堂合計], " _
        & "syoku.syokusu as [食数] ,buy.buyto as [売店合計],buy.buysu as [品数] ,syoku.kai_cd as [会社コード] from " _
        & "(SELECT syain.sya_cd,syain.sya_nm,sum(uriage.kingaku) as syokuTo,count(kingaku) as syokusu,syain.kai_cd " _
        & "FROM syohin LEFT OUTER JOIN (uriage INNER JOIN syain ON uriage.sya_cd = syain.sya_cd) ON uriage.hin_cd = syohin.hin_cd " _
        & "where syohin.bmn_cd in ('10','20','30') and eigyo_date between '" & Text2.Text & "' and '" & Text3.Text & "' " _
        & "group by syain.sya_cd,syain.sya_nm,syain.kai_cd) as Syoku LEFT OUTER JOIN (SELECT syain.sya_cd,sum(uriage.kingaku) as buyTo,count(kingaku) as buysu " _
        & "FROM syohin LEFT OUTER JOIN (uriage INNER JOIN syain ON uriage.sya_cd = syain.sya_cd) ON uriage.hin_cd = syohin.hin_cd " _
        & "where syohin.bmn_cd in ('11','21','31') and eigyo_date between '" & Text2.Text & "' and '" & Text3.Text & "' " _
        & "group by syain.sya_cd) as Buy on Syoku.sya_cd = Buy.sya_cd order by syoku.kai_cd,syoku.sya_cd" _
        
        Form4.Visible = False
        frmDBTable.Show
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

'
Private Sub Form_Unload(Cancel As Integer)

''▼接続を閉じ自身を非表示にしform1を表示-----------------------
''    ▼rs,cn状態確認及び閉じる処理-----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        If cn.State <> 0 Then
        cn.Close
        End If
'    ▲--------------------------------------------------------
    ssql = ""
    pSql = ""

End Sub
''▲------------------------------------------------------------
