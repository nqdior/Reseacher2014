VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "vsview3.OCX"
Begin VB.Form frmDBTable 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Print Menu"
   ClientHeight    =   13425
   ClientLeft      =   3120
   ClientTop       =   2250
   ClientWidth     =   17760
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   13425
   ScaleWidth      =   17760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command2 
      Caption         =   "次頁(&N)"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "前頁(&B)"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdToPrinter 
      Caption         =   "印刷(P)"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "縮小(&O)"
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "拡大(&I)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4860
      Width           =   1215
   End
   Begin vsViewLib.vsPrinter vsPrinter1 
      Height          =   4695
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      _Version        =   196608
      _ExtentX        =   8916
      _ExtentY        =   8281
      _StockProps     =   229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "メイリオ"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "メイリオ"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Zoom            =   50
      ZoomStep        =   10
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2700
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   3.6701e-38
   End
End
Attribute VB_Name = "frmDBTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const mintControlMargin = 60
Private fntHeader As New StdFont
Private fntBody As New StdFont


'表を印刷します。
Private Sub SetTableData(strDbPath As String)
Dim szTitle As Integer
Dim szCalc As Integer
Dim strFormat As String
Dim strHeader As String
Dim lngHeaderShade As Long
Dim strBody As String
Dim x As Integer
Dim i As Integer
Dim kai As Integer
Dim cn As New ADODB.Connection
Dim rs As ADODB.Recordset
Dim breakp As String
Dim Total As Long
Dim kaiTotal As Long

    
    'データベースをオープン =====================================================================================================
    
    cn.ConnectionString = cnstr
    cn.Open
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.Open ssql1, cn, adOpenStatic, adLockReadOnly

    '============================================================================================================================
    
        With vsPrinter1

        '表組みヘッダの書式
        strFormat = "<+1500|^1800|^1800|^1400|^1400|^1400|^1400|<~1400;"
        '表組みヘッダの文字
        strHeader = "会社コード|社員ID|社員名|食堂金額|食数|売店金額|品数|合計;"

        '表組み本文の文字
        lngHeaderShade = RGB(175, 238, 238)
        '段落の後の間隔
        .SpaceAfter = 60
        '段落の前の間隔
        .SpaceBefore = .SpaceAfter
        '印刷開始
        .StartDoc
        '表の罫線
        .TableBorder = tbTop
        '罫線の太さ
        .PenWidth = 0
        '表組みの行位置
        .TextAlign = taCenterTop
        '表組みの設定
        Me.Caption = "表を作成中..."
        .AddTable strFormat, strHeader, "", lngHeaderShade
        
        
    '表示設定部分===============================================================================================================
        If rs.EOF Then
        MsgBox "データが存在しません"
        Else
        rs.MoveLast
        kai = rs.RecordCount
        rs.MoveFirst
        End If
        
        For x = 1 To kai
            If x < kai Then
                
                If x = 720 Then
                    MsgBox ("")
                End If
                
                If breakp = "" Then
                    If rs.Fields("売店合計") <> "" Then '売店合計がNullレコード時
                        strBody = rs.Fields("会社コード") & "|" & rs.Fields("社員ID") & "|" & rs.Fields("社員名") & "|" & "\" & Format(rs.Fields("食堂合計"), "#,###,###") & "|" _
                        & Format(rs.Fields("食数"), "#,###,###") & "|" & "\" & Format(rs.Fields("売店合計"), "#,###,###") & "|" & Format(rs.Fields("品数"), "#,###,###") & "|" & "\" & Format(rs.Fields("総合計"), "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
    
                        Total = Total + rs.Fields("総合計")
                        kaiTotal = kaiTotal + rs.Fields("総合計")
                        breakp = rs.Fields("会社コード")
                        rs.MoveNext
                        
                    Else                                    '売店合計が存在する時
                        strBody = rs.Fields("会社コード") & "|" & rs.Fields("社員ID") & "|" & rs.Fields("社員名") & "|" & "\" & Format(rs.Fields("食堂合計"), "#,###,###") & "|" _
                        & Format(rs.Fields("食数"), "#,###,###") & "|" & Format(rs.Fields("売店合計"), "#,###,###") & "|" & Format(rs.Fields("品数"), "#,###,###") & "|" & "\" & Format(rs.Fields("総合計"), "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                                            
                        kaiTotal = kaiTotal + rs.Fields("総合計")
                        breakp = rs.Fields("会社コード")
                        rs.MoveNext
                                            
                    End If
                    
                Else
                    
                    If breakp <> rs.Fields("会社コード") Then '会社コード表示時判定
                        strBody = "" & "|" & "" & "|" & "" & "|" & "" & "|" _
                        & "" & "|" & breakp & "|" & "合計" & "|" & "\" & Format(kaiTotal, "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                        breakp = rs.Fields("会社コード")
                        kaiTotal = 0
                        
                    ElseIf rs.Fields("売店合計") <> "" Then '売店合計がNullレコード時
                        strBody = rs.Fields("会社コード") & "|" & rs.Fields("社員ID") & "|" & rs.Fields("社員名") & "|" & "\" & Format(rs.Fields("食堂合計"), "#,###,###") & "|" _
                        & Format(rs.Fields("食数"), "#,###,###") & "|" & "\" & Format(rs.Fields("売店合計"), "#,###,###") & "|" & Format(rs.Fields("品数"), "#,###,###") & "|" & "\" & Format(rs.Fields("総合計"), "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
    
                        Total = Total + rs.Fields("総合計")
                        kaiTotal = kaiTotal + rs.Fields("総合計")
                        breakp = rs.Fields("会社コード")
                        rs.MoveNext
                        
                    Else                                    '売店合計が存在する時
                        strBody = rs.Fields("会社コード") & "|" & rs.Fields("社員ID") & "|" & rs.Fields("社員名") & "|" & "\" & Format(rs.Fields("食堂合計"), "#,###,###") & "|" _
                        & Format(rs.Fields("食数"), "#,###,###") & "|" & Format(rs.Fields("売店合計"), "#,###,###") & "|" & Format(rs.Fields("品数"), "#,###,###") & "|" & "\" & Format(rs.Fields("総合計"), "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                                            
                        kaiTotal = kaiTotal + rs.Fields("総合計")
                        breakp = rs.Fields("会社コード")
                        rs.MoveNext
                                            
                    End If
                    
                End If
                
            Else '最終レコードに辿り着いた時
                                      
                    strBody = rs.Fields("会社コード") & "|" & rs.Fields("社員ID") & "|" & rs.Fields("社員名") & "|" & "\" & Format(rs.Fields("食堂合計"), "#,###,###") & "|" _
                    & Format(rs.Fields("食数"), "#,###,###") & "|" & Format(rs.Fields("売店合計"), "#,###,###") & "|" & Format(rs.Fields("品数"), "#,###,###") & "|" & "\" & Format(rs.Fields("総合計"), "#,###,###") & ";"
                    .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                    Total = Total + rs.Fields("総合計")
                    
                    strBody = "" & "|" & "" & "|" & "" & "|" & "" & "|" _
                    & "" & "|" & breakp & "|" & "合計" & "|" & "\" & Format(kaiTotal, "#,###,###") & ";"
                    .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                    breakp = rs.Fields("会社コード")

                    
                    strBody = "" & "|" & "" & "|" & "" & "|" _
                    & "" & "|" & "" & "|" & "" & "|" & "総合計" & "|" & "\" & Format(Total, "#,###,###") & ";"
                    .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                            
            End If

            
        Next x
       
        .EndDoc
    End With
    
    '============================================================================================================================
        
        
   'フォントの設定
    With fntHeader
        .Name = "メイリオ"
        .Size = 10
        .Bold = True
    End With
    With fntBody
        .Name = "メイリオ"
        .Size = 10
        .Bold = False
    End With
        
End Sub

Private Sub cmdToPrinter_Click() '=========================以下各種設定================================================
    
    'ドキュメント名を設定
    vsPrinter1.FileName = "個人別利用詳細印刷"
    '印刷
    vsPrinter1.PrintDoc

End Sub

Private Sub cmdZoom_Click(index As Integer)
Const ZoomParam = 25
    
    Screen.MousePointer = vbHourglass
    
    With vsPrinter1
        Select Case index
            Case 0
                .Zoom = .Zoom + ZoomParam
                'Zoomの許容範囲を外れる場合、[拡大]ボタンを使用不可能に設定
                If .Zoom > 150 - ZoomParam Then
                    cmdZoom(0).Enabled = False
                End If
                '[縮小]ボタンを使用可能に設定
                If Not cmdZoom(1).Enabled Then
                    cmdZoom(1).Enabled = True
                End If
                Case 1
                .Zoom = .Zoom - ZoomParam
                'Zoomの許容範囲を外れる場合、[縮小]ボタンを使用不可能に設定
                If .Zoom <= 0 + ZoomParam Then
                    cmdZoom(1).Enabled = False
                End If
                '[拡大]ボタンを使用可能に設定
                If Not cmdZoom(0).Enabled Then
                    cmdZoom(0).Enabled = True
                End If
        End Select
        Me.Caption = "Zoom " & .Zoom & "%"
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub Command1_Click()
            If vsPrinter1.PreviewPage > 1 Then
                vsPrinter1.PreviewPage = vsPrinter1.PreviewPage - 1
            End If
End Sub

Private Sub Command2_Click()
            If vsPrinter1.PreviewPage < vsPrinter1.CurrentPage Then
                vsPrinter1.PreviewPage = vsPrinter1.PreviewPage + 1
            End If
End Sub

Private Sub Form_Load() '----------------------------------------------------------------
Dim strDataBasepath As String
    
    Screen.MousePointer = vbHourglass
    With vsPrinter1
        '各ページに印刷されるフッタを設定
        .Footer = "|- %d -|"
        '各ページに印刷されるヘッダを設定
        .Header = Format(Form4.Text2.Text, "####/##/##") & " ～ " & Format(Form4.Text3.Text, "####/##/##") & "|個人別利用詳細一覧|"
        'ヘッダとフッタの印刷に使用するフォント
        .HdrFontName = "メイリオ"
        .HdrFontSize = 10
        'マウスをドラッグすることによりページのプレビューをスクロール
        .MouseScroll = True
        '各ページの周りに描かれるページ枠を設定
        .PageBorder = pbNone
        'Printerコントロールの出力を全て画面へ
        .Preview = True
        'プレビュー画面の表示モードを画面の色に設定
        .PreviewMode = pmScreen
        'プレビュー画面の縮尺率
        .Zoom = 100
        'マウスによるズームを禁止
        .MouseZoom = False
        '用紙サイズをＡ４に設定
        If .PaperSizes(vbPRPSA4) = True Then
            .PaperSize = vbPRPSA4
        Else
            MsgBox "用紙サイズを設定できませんでした。", vbExclamation
            Exit Sub
        End If
        '用紙方向を横に設定
        
            .Orientation = orLandscape
            If .Error <> 0 Or .Orientation = orPortrait Then
                MsgBox "用紙方向を設定できませんでした。", vbExclamation
                Exit Sub
            End If

        Left = (Screen.Width - Width) \ 2
        Top = (Screen.Height - Height) \ 2
        Show
        DoEvents

            '表を印刷
            Call SetTableData(strDataBasepath)
'        End If
    End With
    Me.Caption = "Zoom 100%"
    Screen.MousePointer = vbDefault
    
    
    

End Sub '-----------------------------------------------------------------------------------------


'コントロールのサイズをフォームサイズに合わせます。-----------------------------------------------
Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then Exit Sub
    With cmdToPrinter
        .Left = Me.ScaleWidth - mintControlMargin - .Width
        .Top = Me.ScaleHeight - mintControlMargin - .Height
    End With
    With Command1
        .Left = cmdZoom(0).Left + cmdZoom(0).Width + mintControlMargin / 2 + 1500
        .Top = Me.ScaleHeight - mintControlMargin - .Height
    End With
    With Command2
        .Left = cmdZoom(0).Left + cmdZoom(0).Width + mintControlMargin / 2 + 2750
        .Top = Me.ScaleHeight - mintControlMargin - .Height
    End With
    With cmdZoom(0)
        .Left = mintControlMargin
        .Top = Me.ScaleHeight - mintControlMargin - .Height
    End With
    With cmdZoom(1)
        .Left = cmdZoom(0).Left + cmdZoom(0).Width + mintControlMargin / 2
        .Top = Me.ScaleHeight - mintControlMargin - .Height
    End With
    With vsPrinter1
        .Left = mintControlMargin
        .Top = mintControlMargin
        .Height = Me.ScaleHeight - cmdZoom(0).Height - 3 * mintControlMargin
        .Width = Me.ScaleWidth - 2 * mintControlMargin
    End With

End Sub '------------------------------------------------------------------------------------------

'指定位置に指定の長さで縦線または横線を引きます。
Sub PutLine(leftPos As Single, topPos As Single, pLineSize As String, TateYoko As String)
    Dim previewf As Variant
    With previewf.vsPrinter1
        Select Case TateYoko
            Case "|"   '縦
                .DrawLine leftPos, topPos, leftPos, topPos + pLineSize
            
            Case "-"   '横
                .DrawLine leftPos, topPos, leftPos + pLineSize, topPos
        End Select
    End With
End Sub
