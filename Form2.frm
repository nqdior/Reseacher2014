VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Data Editor"
   ClientHeight    =   11790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18495
   BeginProperty Font 
      Name            =   "���C���I"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   11790
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   14107.55
   StartUpPosition =   2  '��ʂ̒���
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   14760
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���C���I"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FontSize"
      Height          =   975
      Left            =   10320
      TabIndex        =   16
      Tag             =   "invis"
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command6 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "���C���I"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   19
         Tag             =   "invis"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "���C���I"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Tag             =   "invis"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Reset"
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Output (&O)"
      Height          =   615
      Left            =   17280
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TableSave"
      Height          =   975
      Left            =   12960
      TabIndex        =   11
      Tag             =   "invis"
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         Height          =   495
         Left            =   2880
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Save"
         Height          =   495
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   240
         TabIndex        =   12
         Text            =   "TableName"
         Top             =   360
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2280
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reseach (F5)"
      Height          =   1070
      Left            =   17280
      TabIndex        =   2
      Top             =   10800
      Width           =   1095
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
      Height          =   9375
      Left            =   120
      OleObjectBlob   =   "Form2.frx":27A2
      TabIndex        =   0
      Top             =   1200
      Width           =   18255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Tag             =   "invis"
      Top             =   120
      Width           =   10095
      Begin VB.ComboBox DataCombo3 
         Height          =   390
         Left            =   5880
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�~��"
         Height          =   375
         Left            =   7920
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "invis"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox DataCombo1 
         Height          =   390
         Left            =   240
         TabIndex        =   6
         Tag             =   "invis"
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox DataCombo2 
         Height          =   390
         Left            =   3000
         TabIndex        =   5
         Tag             =   "invis"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Sort (&S)"
         Height          =   495
         Left            =   8760
         TabIndex        =   4
         Tag             =   "invis"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�~��"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "invis"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�~��"
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "invis"
         Top             =   360
         Width           =   855
      End
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   10800
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form2.frx":4FCB
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'ver1.1 joint�Ή�/20141121������������������������������������������������
Private Sub Text2_Gotfocus()
    Text2.Text = ""
End Sub

Private Sub Command8_Click()

'��sSql��OrderBy���܂܂�Ă��邩���肵�A���݂����OrderBy�ȍ~��؂�̂�
'ver0.9����ǉ�

If Text2.Text Like "*[a-z,A-Z]*" Then
Else
    MsgBox "�e�[�u�����͔��p�p���݂̂œ��͂��Ă�������"
    Exit Sub
End If

tmp = InStr(1, tableX, "order")
If tmp = 0 Then
    tableX = Text1.Text
Else
    tableX = Left(Text1.Text, tmp - 2)
End If

'�\�[�g�ɂ��s��������͂�����̃R�����g�폜�ɂđΉ��\�B
'���̏ꍇ�\�[�g����SQL�͕ʃt�H�[������ڍs����SQL+Order�ƂȂ�B
'��---------------------------------------------------------------------

tableXnm = Text2.Text
MsgBox ("�\�����̃e�[�u����ۑ����܂����B" & vbCrLf & "�e�[�u���� " & tableXnm)


End Sub
'��������������������������������������������������������������������������

Private Sub Command1_Click()
Text1.Text = tableX
Call Command3_Click
End Sub



Private Sub Form_activate()
On Error GoTo SQLERR


    Adodc1.ConnectionString = cnstr
    Adodc1.RecordSource = ssql
    TDBGrid1.DataSource = Adodc1
    TDBGrid1.Refresh
    
''��20141125 ver1.2���data editor�̕\�����@�ύX----------------
'Call Command3_Click
''��------------------------------------------------------------

'    '��rs��Ԋm�F�y�я���----------------------------------------
'    If cn.State = 0 Then
'    cn.Open cnstr
'    End If
'    If rs.State = 0 Then
'    ssql = Text1.Text
'    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
'    End If
'    '��----------------------------------------------------------
'
''��rs_c3�Ƀf�[�^���i�[��rs����ɂ���---------------------------
'    Set rs_c3 = rs.Clone
'        Set TDBGrid1.DataSource = Nothing
'        Set TDBGrid1.DataSource = rs_c3
'    TDBGrid1.Refresh
'    rs.Close
'    ssql = ""
''��------------------------------------------------------------


'���\�[�g�p�R���{�{�b�N�X�̒l��TDBGrid���擾-----------------
    For i = 0 To TDBGrid1.Columns.Count - 1
         DataCombo1.AddItem (TDBGrid1.Columns(i).Caption)
    Next i
    For i = 0 To TDBGrid1.Columns.Count - 1
         DataCombo2.AddItem (TDBGrid1.Columns(i).Caption)
    Next i
    For i = 0 To TDBGrid1.Columns.Count - 1
         DataCombo3.AddItem (TDBGrid1.Columns(i).Caption)
    Next i
'��------------------------------------------------------------

SQLERR:
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo SQLERR
'����ʃ��T�C�Y���̊e�I�u�W�F�N�g�ʒu�ύX----------------------
    TDBGrid1.Width = Me.ScaleWidth - 200
    TDBGrid1.Height = Me.ScaleHeight - 2550
    Text1.Top = TDBGrid1.Height + 1350
    Command3.Top = TDBGrid1.Height + 1350
    Command3.Left = Command3.Left + (TDBGrid1.Width - Text1.Width - 835.24)
    Text1.Width = TDBGrid1.Width - 835.24
    
    Command4.Left = TDBGrid1.Width - 770
    If Command4.Left <= 13000 Then  '����R�}���h���\�[�g�Q�Ɋ�������\�[�g�Q���\��
        For Each ctl In Me.Controls
        If (ctl.Tag = "invis") Then
        ctl.Visible = False
        End If
        Next ctl
    End If
    If Command4.Left > 13001 Then '�\�[�g�Q�ւ̊����Ȃ��Ȃ�΍ĕ\��
        For Each ctl In Me.Controls
        If (ctl.Tag = "invis") Then
        ctl.Visible = True
        End If
        Next ctl
    End If
'��-----------------------------------------------------------
SQLERR:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
''���ڑ�������g���\���ɂ�form1��\��-----------------------
''    ��rs,cn��Ԋm�F�y�ѕ��鏈��-----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        If cn.State <> 0 Then
        cn.Close
        End If
'    ��--------------------------------------------------------

    '��ver1.1 joint�Ή�/20141121 ��������������������������������������������
        
        If tableX <> "" Then
            Form1.Label3.Caption = tableXnm
        Else: Form1.Label3.Caption = "Saving Nothing"
        End If
        If Form1.Text3.Visible = True Then
            Form1.Text3.Text = ""
            Form1.Text3.Visible = False
        End If
        If Form1.Combo3.Visible = True Then
            Form1.Combo3.Clear
            Form1.Combo3.Visible = False
        End If
        
        If Nowform = "form1" Then
            Form1.DataCombo1.Text = "Table Name"
            Form1.DataCombo2.Text = "Column Name"
        End If
    '������������������������������������������������������������������������
    
    ssql = ""
    pSql = ""
    DataCombo1.Clear
    DataCombo2.Clear
    DataCombo3.Clear
    Form2.Visible = False
    
    If Nowform = "form0" Then
        Form0.Visible = True
    ElseIf Nowform = "form1" Then
        Call Form1.Form_activate
        Form1.Visible = True
    ElseIf Nowform = "form3" Then
        Call Form3.Form_activate
        Form3.Visible = True
    ElseIf Nowform = "form5" Then
        Call Form5.Form_activate
        Form5.Visible = True
    End If
End Sub
'��------------------------------------------------------------
'�������T�C�Y�ύX�{�^��----------------------------------------
Private Sub Command5_Click()
TDBGrid1.Font.Size = TDBGrid1.Font.Size + 1
End Sub

Private Sub Command6_Click()
TDBGrid1.Font.Size = TDBGrid1.Font.Size - 1
End Sub

Private Sub Command7_Click()
TDBGrid1.Font.Size = 9
End Sub
'��------------------------------------------------------------

'��F5�L�[�������A�N�V����-ERR001-------------------------------
'�����̂��@�\���Ă��Ȃ����ߋC�����邱�Ɓ���������
Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command3_Click
    End If
End Sub
'��------------------------------------------------------------


'��F5�L�[�������A�N�V����-ERR001��U�ɍ쐬---------------------
Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command3_Click
    End If
End Sub
'��------------------------------------------------------------
    
'����ł�SQL���s-----------------------------------------------
Private Sub Command3_Click()
Form2.MousePointer = 11
On Error GoTo SQLERR

    Adodc1.ConnectionString = cnstr
    Adodc1.RecordSource = ssql
    TDBGrid1.DataSource = Adodc1
    TDBGrid1.Refresh
'
'    '��rs��Ԋm�F�y�я���--------------------------------------
'    If cn.State = 0 Then
'        cn.Open cnstr
'    End If
'    If rs.State <> 0 Then
'        rs.Close
'    End If
'    '��--------------------------------------------------------
'
'    ssql = Form2.Text1.Text
'    rs.Open ssql, cn ', adOpenStatic, adLockOptimistic, adCmdText
'    Set rs_c4 = rs.Clone
'    rs.Close
'    Set TDBGrid1.DataSource = rs_c4
'    TDBGrid1.Refresh
'
'���\�[�g�p�R���{�{�b�N�X�̒l��TDBGrid���擾-----------------
    DataCombo1.Clear
    For i = 0 To TDBGrid1.Columns.Count - 1
         DataCombo1.AddItem (TDBGrid1.Columns(i).Caption)
    Next i
    DataCombo2.Clear
    For i = 0 To TDBGrid1.Columns.Count - 1
         DataCombo2.AddItem (TDBGrid1.Columns(i).Caption)
    Next i
    DataCombo3.Clear
    For i = 0 To TDBGrid1.Columns.Count - 1
         DataCombo3.AddItem (TDBGrid1.Columns(i).Caption)
    Next i
'��------------------------------------------------------------
    
    ssql = ""
    Form2.MousePointer = 0
    
SQLERR:
    If Err.Number <> 0 Then
        Form2.MousePointer = 0
        MsgBox "ErrorNo: " & Err.Number & vbCrLf & "ErrorMassage: " & Err.Description
        Exit Sub
    End If
End Sub
'��------------------------------------------------------------

Private Sub Command2_Click()
On Error GoTo SQLERR

'��sSql��OrderBy���܂܂�Ă��邩���肵�A���݂����OrderBy�ȍ~��؂�̂�
ssql = Text1.Text
tmp = InStr(1, ssql, "order")
If tmp = 0 Then
    pSql = Text1.Text
Else
    pSql = Left(Text1.Text, tmp - 2)
End If
'�\�[�g�ɂ��s��������͂�����̃R�����g�폜�ɂđΉ��\�B
'���̏ꍇ�\�[�g����SQL�͕ʃt�H�[������ڍs����SQL+Order�ƂȂ�B
'��---------------------------------------------------------------------

'���\�[�g�����d���󔒔���--------------------------------------
    If DataCombo1.Text = "" And DataCombo2.Text = "" And DataCombo3.Text = "" Then
        MsgBox ("�\�[�g�L�[���ݒ肳��Ă��܂���")
        Exit Sub
    ElseIf DataCombo1.Text = DataCombo2.Text And DataCombo1.Text <> "" Then
        MsgBox ("�\�[�g�L�[���d�����Ă��܂�")
        Exit Sub
    ElseIf DataCombo1.Text = DataCombo3.Text And DataCombo1.Text <> "" Then
        MsgBox ("�\�[�g�L�[���d�����Ă��܂�")
        Exit Sub
    ElseIf DataCombo2.Text = DataCombo3.Text And DataCombo2.Text <> "" Then
        MsgBox ("�\�[�g�L�[���d�����Ă��܂�")
        Exit Sub
    End If
'��------------------------------------------------------------

Set TDBGrid1.DataSource = Nothing

'���\�[�g���s�������L�q----------------------------------------
    If DataCombo1.Text <> "" Then
        ssql = pSql & " order by " & DataCombo1.Text
        If Check1.Value = 1 Then
               ssql = ssql & " desc"
        End If
    Else: Exit Sub
    End If
    If DataCombo2.Text <> "" Then
        ssql = ssql & ", " & DataCombo2.Text
        If Check2.Value = 1 Then
               ssql = ssql & " desc"
        End If
    Else
    End If
    If DataCombo3.Text <> "" Then
        ssql = ssql & ", " & DataCombo3.Text
        If Check3.Value = 1 Then
               ssql = ssql & " desc"
        End If
    Else
    End If
    
    '��rs��Ԋm�F�y�я���----------------------------------------
    If rs.State <> 0 Then
    rs.Close
    End If
    If cn.State = 0 Then
    cn.Open cnstr
    End If
    '��----------------------------------------------------------
    
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c5 = rs.Clone
    Set TDBGrid1.DataSource = rs_c5
    TDBGrid1.Refresh
    Text1.Text = ssql
    ssql = ""
'��------------------------------------------------------------

''    ��rs,cn��Ԋm�F�y�ѕ��鏈��-----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
'    ��--------------------------------------------------------

SQLERR:
    Exit Sub
End Sub


Private Sub Command4_Click()
On Error GoTo SQLERR

Dim myfile As String 'CDL�A�h���X�ۑ��p
Dim rscount As Integer '�s�J�E���g�p

CommonDialog1.Filter = "÷��(*.csv)|*.csv|���ׂ�(*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = cdlOFNOverwritePrompt
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
myfile = CommonDialog1.FileName

'��rs��Ԋm�F�y�я���----------------------------------------
If rs.State <> 0 Then
rs.Close
End If
If cn.State = 0 Then
cn.Open cnstr
End If
'��----------------------------------------------------------

ssql = Text1.Text
rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText

'���ŏI���R�[�h�̃J�E���g------------------------------------
rs.MoveLast
rscount = rs.RecordCount - 1
rs.MoveFirst
'��----------------------------------------------------------

'�����ږ��o��------------------------------------------------
Open myfile For Output As #1
    For i = 0 To rs.Fields.Count - 1
        Print #1, rs.Fields(i).Name & ",";
    Next i
    Print #1, vbCrLf;
'���f�[�^�o��
For x = 0 To rscount
    For i = 0 To rs.Fields.Count - 1
        Print #1, rs.Fields(i) & ",";
    Next i
    Print #1, vbCrLf;
    rs.MoveNext
Next x
Close #1
'��-----------------------------------------------------------
'��-----------------------------------------------------------
''    ��rs,cn��Ԋm�F�y�ѕ��鏈��-----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
'    ��--------------------------------------------------------

SQLERR:
    Exit Sub
End Sub

