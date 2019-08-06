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
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   13425
   ScaleWidth      =   17760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command2 
      Caption         =   "����(&N)"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�O��(&B)"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdToPrinter 
      Caption         =   "���(P)"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "�k��(&O)"
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "�g��(&I)"
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
         Name            =   "���C���I"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���C���I"
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


'�\��������܂��B
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

    
    '�f�[�^�x�[�X���I�[�v�� =====================================================================================================
    
    cn.ConnectionString = cnstr
    cn.Open
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.Open ssql1, cn, adOpenStatic, adLockReadOnly

    '============================================================================================================================
    
        With vsPrinter1

        '�\�g�݃w�b�_�̏���
        strFormat = "<+1500|^1800|^1800|^1400|^1400|^1400|^1400|<~1400;"
        '�\�g�݃w�b�_�̕���
        strHeader = "��ЃR�[�h|�Ј�ID|�Ј���|�H�����z|�H��|���X���z|�i��|���v;"

        '�\�g�ݖ{���̕���
        lngHeaderShade = RGB(175, 238, 238)
        '�i���̌�̊Ԋu
        .SpaceAfter = 60
        '�i���̑O�̊Ԋu
        .SpaceBefore = .SpaceAfter
        '����J�n
        .StartDoc
        '�\�̌r��
        .TableBorder = tbTop
        '�r���̑���
        .PenWidth = 0
        '�\�g�݂̍s�ʒu
        .TextAlign = taCenterTop
        '�\�g�݂̐ݒ�
        Me.Caption = "�\���쐬��..."
        .AddTable strFormat, strHeader, "", lngHeaderShade
        
        
    '�\���ݒ蕔��===============================================================================================================
        If rs.EOF Then
        MsgBox "�f�[�^�����݂��܂���"
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
                    If rs.Fields("���X���v") <> "" Then '���X���v��Null���R�[�h��
                        strBody = rs.Fields("��ЃR�[�h") & "|" & rs.Fields("�Ј�ID") & "|" & rs.Fields("�Ј���") & "|" & "\" & Format(rs.Fields("�H�����v"), "#,###,###") & "|" _
                        & Format(rs.Fields("�H��"), "#,###,###") & "|" & "\" & Format(rs.Fields("���X���v"), "#,###,###") & "|" & Format(rs.Fields("�i��"), "#,###,###") & "|" & "\" & Format(rs.Fields("�����v"), "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
    
                        Total = Total + rs.Fields("�����v")
                        kaiTotal = kaiTotal + rs.Fields("�����v")
                        breakp = rs.Fields("��ЃR�[�h")
                        rs.MoveNext
                        
                    Else                                    '���X���v�����݂��鎞
                        strBody = rs.Fields("��ЃR�[�h") & "|" & rs.Fields("�Ј�ID") & "|" & rs.Fields("�Ј���") & "|" & "\" & Format(rs.Fields("�H�����v"), "#,###,###") & "|" _
                        & Format(rs.Fields("�H��"), "#,###,###") & "|" & Format(rs.Fields("���X���v"), "#,###,###") & "|" & Format(rs.Fields("�i��"), "#,###,###") & "|" & "\" & Format(rs.Fields("�����v"), "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                                            
                        kaiTotal = kaiTotal + rs.Fields("�����v")
                        breakp = rs.Fields("��ЃR�[�h")
                        rs.MoveNext
                                            
                    End If
                    
                Else
                    
                    If breakp <> rs.Fields("��ЃR�[�h") Then '��ЃR�[�h�\��������
                        strBody = "" & "|" & "" & "|" & "" & "|" & "" & "|" _
                        & "" & "|" & breakp & "|" & "���v" & "|" & "\" & Format(kaiTotal, "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                        breakp = rs.Fields("��ЃR�[�h")
                        kaiTotal = 0
                        
                    ElseIf rs.Fields("���X���v") <> "" Then '���X���v��Null���R�[�h��
                        strBody = rs.Fields("��ЃR�[�h") & "|" & rs.Fields("�Ј�ID") & "|" & rs.Fields("�Ј���") & "|" & "\" & Format(rs.Fields("�H�����v"), "#,###,###") & "|" _
                        & Format(rs.Fields("�H��"), "#,###,###") & "|" & "\" & Format(rs.Fields("���X���v"), "#,###,###") & "|" & Format(rs.Fields("�i��"), "#,###,###") & "|" & "\" & Format(rs.Fields("�����v"), "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
    
                        Total = Total + rs.Fields("�����v")
                        kaiTotal = kaiTotal + rs.Fields("�����v")
                        breakp = rs.Fields("��ЃR�[�h")
                        rs.MoveNext
                        
                    Else                                    '���X���v�����݂��鎞
                        strBody = rs.Fields("��ЃR�[�h") & "|" & rs.Fields("�Ј�ID") & "|" & rs.Fields("�Ј���") & "|" & "\" & Format(rs.Fields("�H�����v"), "#,###,###") & "|" _
                        & Format(rs.Fields("�H��"), "#,###,###") & "|" & Format(rs.Fields("���X���v"), "#,###,###") & "|" & Format(rs.Fields("�i��"), "#,###,###") & "|" & "\" & Format(rs.Fields("�����v"), "#,###,###") & ";"
                        .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                                            
                        kaiTotal = kaiTotal + rs.Fields("�����v")
                        breakp = rs.Fields("��ЃR�[�h")
                        rs.MoveNext
                                            
                    End If
                    
                End If
                
            Else '�ŏI���R�[�h�ɒH�蒅������
                                      
                    strBody = rs.Fields("��ЃR�[�h") & "|" & rs.Fields("�Ј�ID") & "|" & rs.Fields("�Ј���") & "|" & "\" & Format(rs.Fields("�H�����v"), "#,###,###") & "|" _
                    & Format(rs.Fields("�H��"), "#,###,###") & "|" & Format(rs.Fields("���X���v"), "#,###,###") & "|" & Format(rs.Fields("�i��"), "#,###,###") & "|" & "\" & Format(rs.Fields("�����v"), "#,###,###") & ";"
                    .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                    Total = Total + rs.Fields("�����v")
                    
                    strBody = "" & "|" & "" & "|" & "" & "|" & "" & "|" _
                    & "" & "|" & breakp & "|" & "���v" & "|" & "\" & Format(kaiTotal, "#,###,###") & ";"
                    .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                    breakp = rs.Fields("��ЃR�[�h")

                    
                    strBody = "" & "|" & "" & "|" & "" & "|" _
                    & "" & "|" & "" & "|" & "" & "|" & "�����v" & "|" & "\" & Format(Total, "#,###,###") & ";"
                    .AddTable strFormat, strHeader, strBody, lngHeaderShade, vbWhite, True
                            
            End If

            
        Next x
       
        .EndDoc
    End With
    
    '============================================================================================================================
        
        
   '�t�H���g�̐ݒ�
    With fntHeader
        .Name = "���C���I"
        .Size = 10
        .Bold = True
    End With
    With fntBody
        .Name = "���C���I"
        .Size = 10
        .Bold = False
    End With
        
End Sub

Private Sub cmdToPrinter_Click() '=========================�ȉ��e��ݒ�================================================
    
    '�h�L�������g����ݒ�
    vsPrinter1.FileName = "�l�ʗ��p�ڍ׈��"
    '���
    vsPrinter1.PrintDoc

End Sub

Private Sub cmdZoom_Click(index As Integer)
Const ZoomParam = 25
    
    Screen.MousePointer = vbHourglass
    
    With vsPrinter1
        Select Case index
            Case 0
                .Zoom = .Zoom + ZoomParam
                'Zoom�̋��e�͈͂��O���ꍇ�A[�g��]�{�^�����g�p�s�\�ɐݒ�
                If .Zoom > 150 - ZoomParam Then
                    cmdZoom(0).Enabled = False
                End If
                '[�k��]�{�^�����g�p�\�ɐݒ�
                If Not cmdZoom(1).Enabled Then
                    cmdZoom(1).Enabled = True
                End If
                Case 1
                .Zoom = .Zoom - ZoomParam
                'Zoom�̋��e�͈͂��O���ꍇ�A[�k��]�{�^�����g�p�s�\�ɐݒ�
                If .Zoom <= 0 + ZoomParam Then
                    cmdZoom(1).Enabled = False
                End If
                '[�g��]�{�^�����g�p�\�ɐݒ�
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
        '�e�y�[�W�Ɉ�������t�b�^��ݒ�
        .Footer = "|- %d -|"
        '�e�y�[�W�Ɉ�������w�b�_��ݒ�
        .Header = Format(Form4.Text2.Text, "####/##/##") & " �` " & Format(Form4.Text3.Text, "####/##/##") & "|�l�ʗ��p�ڍ׈ꗗ|"
        '�w�b�_�ƃt�b�^�̈���Ɏg�p����t�H���g
        .HdrFontName = "���C���I"
        .HdrFontSize = 10
        '�}�E�X���h���b�O���邱�Ƃɂ��y�[�W�̃v���r���[���X�N���[��
        .MouseScroll = True
        '�e�y�[�W�̎���ɕ`�����y�[�W�g��ݒ�
        .PageBorder = pbNone
        'Printer�R���g���[���̏o�͂�S�ĉ�ʂ�
        .Preview = True
        '�v���r���[��ʂ̕\�����[�h����ʂ̐F�ɐݒ�
        .PreviewMode = pmScreen
        '�v���r���[��ʂ̏k�ڗ�
        .Zoom = 100
        '�}�E�X�ɂ��Y�[�����֎~
        .MouseZoom = False
        '�p���T�C�Y���`�S�ɐݒ�
        If .PaperSizes(vbPRPSA4) = True Then
            .PaperSize = vbPRPSA4
        Else
            MsgBox "�p���T�C�Y��ݒ�ł��܂���ł����B", vbExclamation
            Exit Sub
        End If
        '�p�����������ɐݒ�
        
            .Orientation = orLandscape
            If .Error <> 0 Or .Orientation = orPortrait Then
                MsgBox "�p��������ݒ�ł��܂���ł����B", vbExclamation
                Exit Sub
            End If

        Left = (Screen.Width - Width) \ 2
        Top = (Screen.Height - Height) \ 2
        Show
        DoEvents

            '�\�����
            Call SetTableData(strDataBasepath)
'        End If
    End With
    Me.Caption = "Zoom 100%"
    Screen.MousePointer = vbDefault
    
    
    

End Sub '-----------------------------------------------------------------------------------------


'�R���g���[���̃T�C�Y���t�H�[���T�C�Y�ɍ��킹�܂��B-----------------------------------------------
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

'�w��ʒu�Ɏw��̒����ŏc���܂��͉����������܂��B
Sub PutLine(leftPos As Single, topPos As Single, pLineSize As String, TateYoko As String)
    Dim previewf As Variant
    With previewf.vsPrinter1
        Select Case TateYoko
            Case "|"   '�c
                .DrawLine leftPos, topPos, leftPos, topPos + pLineSize
            
            Case "-"   '��
                .DrawLine leftPos, topPos, leftPos + pLineSize, topPos
        End Select
    End With
End Sub