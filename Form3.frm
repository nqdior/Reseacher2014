VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "Calculation Menu"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "���C���I"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8065.283
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   5751
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SavingTable"
      Height          =   2175
      Left            =   3960
      TabIndex        =   38
      Top             =   1680
      Width           =   2175
      Begin VB.CommandButton Command2 
         Caption         =   "TableCall"
         Height          =   495
         Left            =   360
         TabIndex        =   33
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ClearTable"
         Height          =   495
         Left            =   360
         TabIndex        =   34
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Group Key"
      Height          =   2895
      Left            =   240
      TabIndex        =   36
      Top             =   1680
      Width           =   2775
      Begin VB.ComboBox Combo7 
         Height          =   390
         Left            =   240
         TabIndex        =   2
         Tag             =   "syukei"
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo6 
         Height          =   390
         Left            =   240
         TabIndex        =   1
         Tag             =   "syukei"
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo8 
         Height          =   390
         Left            =   240
         TabIndex        =   3
         Tag             =   "syukei"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo9 
         Height          =   390
         Left            =   240
         TabIndex        =   4
         Tag             =   "syukei"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo10 
         Height          =   390
         Left            =   240
         TabIndex        =   5
         Tag             =   "syukei"
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   390
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   390
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   390
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Height          =   390
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo6 
         Height          =   390
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reseach (F5)"
      Height          =   615
      Left            =   120
      TabIndex        =   32
      Top             =   8640
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   390
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   688
      _Version        =   393216
      Text            =   "Table Name "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���C���I"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Column"
      Height          =   3375
      Left            =   240
      TabIndex        =   37
      Top             =   4800
      Width           =   5895
      Begin VB.ComboBox Combo11 
         Height          =   390
         Left            =   240
         TabIndex        =   12
         Tag             =   "syukei"
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo12 
         Height          =   390
         Left            =   240
         TabIndex        =   16
         Tag             =   "syukei"
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo13 
         Height          =   390
         Left            =   240
         TabIndex        =   20
         Tag             =   "syukei"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo14 
         Height          =   390
         Left            =   240
         TabIndex        =   24
         Tag             =   "syukei"
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo15 
         Height          =   390
         Left            =   240
         TabIndex        =   28
         Tag             =   "syukei"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  '�ׯ�
         Height          =   390
         Left            =   2640
         TabIndex        =   14
         Text            =   "sum"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3600
         TabIndex        =   15
         Text            =   " New Column Name 1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  '�ׯ�
         Height          =   390
         Left            =   2640
         TabIndex        =   18
         Text            =   "sum"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  '�ׯ�
         Height          =   390
         Left            =   2640
         TabIndex        =   22
         Text            =   "sum"
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox Combo4 
         Appearance      =   0  '�ׯ�
         Height          =   390
         Left            =   2640
         TabIndex        =   26
         Text            =   "sum"
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox Combo5 
         Appearance      =   0  '�ׯ�
         Height          =   390
         Left            =   2640
         TabIndex        =   30
         Text            =   "sum"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3600
         TabIndex        =   19
         Text            =   " New Column Name 2"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   3600
         TabIndex        =   23
         Text            =   " New Column Name 3"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   3600
         TabIndex        =   27
         Text            =   " New Column Name 4"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   3600
         TabIndex        =   31
         Text            =   " New Column Name 5"
         Top             =   2760
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DataCombo7 
         Height          =   390
         Left            =   240
         TabIndex        =   13
         Tag             =   "syukei"
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo8 
         Height          =   390
         Left            =   240
         TabIndex        =   17
         Tag             =   "syukei"
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo9 
         Height          =   390
         Left            =   240
         TabIndex        =   21
         Tag             =   "syukei"
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo10 
         Height          =   390
         Left            =   240
         TabIndex        =   25
         Tag             =   "syukei"
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo11 
         Height          =   390
         Left            =   240
         TabIndex        =   29
         Tag             =   "syukei"
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   393216
         Text            =   "Column Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���C���I"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "Calculation Menu"
      BeginProperty Font 
         Name            =   "���C���I"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   35
      Top             =   360
      Width           =   1875
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ver1.1joint�Ή�/20141121������������������������������������������������������
Private Sub Command2_Click()
'On Error GoTo SQLERR
        jointflg = "1"
        If tableX = "" Then
            MsgBox ("�ێ�����Ă���e�[�u��������܂���")
            Exit Sub
        End If
        
        '��sSql��OrderBy���܂܂�Ă��邩���肵�A���݂����OrderBy�ȍ~��؂�̂�
        'ver0.9����ǉ�
        
        tmp = InStr(1, tableX, "order")
        If tmp <> 0 Then
            tableX = Left(tableX, tmp - 2)
        End If
        '�\�[�g�ɂ��s��������͂�����̃R�����g�폜�ɂđΉ��\�B
        '���̏ꍇ�\�[�g����SQL�͕ʃt�H�[������ڍs����SQL+Order�ƂȂ�B
        '��---------------------------------------------------------------------
        
        '��ver1.1 joint�Ή�/20141121��������������������������������������������
                Combo6.Clear 'combobox�N���A
                Combo6.Clear: Combo7.Clear: Combo8.Clear: Combo9.Clear: Combo10.Clear:
                Combo11.Clear: Combo12.Clear: Combo13.Clear: Combo14.Clear: Combo15.Clear:
        '������������������������������������������������������������������������
        
        
    '��rs,cn��Ԋm�F�y�ѕ��鏈��----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        If cn.State <> 0 Then
        cn.Close
        End If
    '��--------------------------------------------------------

        cn.Open cnstr
                
        '��Combo�ɃJ�����ꗗ��\������--------------------------
        ssql = tableX
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c1 = rs.Clone 'rs���g���񂷈ׁA�����_�̃N���[����tmp�ɍ쐬
  
        For i = 0 To rs.Fields.Count - 1
            Combo6.AddItem rs.Fields(i).Name
        Next i
        For i = 0 To rs.Fields.Count - 1
            Combo7.AddItem rs.Fields(i).Name
        Next i
        For i = 0 To rs.Fields.Count - 1
            Combo8.AddItem rs.Fields(i).Name
        Next i
        For i = 0 To rs.Fields.Count - 1
            Combo9.AddItem rs.Fields(i).Name
        Next i
                For i = 0 To rs.Fields.Count - 1
            Combo10.AddItem rs.Fields(i).Name
        Next i
        For i = 0 To rs.Fields.Count - 1
            Combo11.AddItem rs.Fields(i).Name
        Next i
                For i = 0 To rs.Fields.Count - 1
            Combo12.AddItem rs.Fields(i).Name
        Next i
        For i = 0 To rs.Fields.Count - 1
            Combo13.AddItem rs.Fields(i).Name
        Next i
                For i = 0 To rs.Fields.Count - 1
            Combo14.AddItem rs.Fields(i).Name
        Next i
        For i = 0 To rs.Fields.Count - 1
            Combo15.AddItem rs.Fields(i).Name
        Next i
        
        '��------------------------------------------------------------
                    
            '��rs��Ԋm�F�y�я���----------------------------------------
            If rs.State <> 0 Then
            rs.Close
            End If
            '��----------------------------------------------------------
            
        For Each ctl In Me.Controls
        If (ctl.Tag = "syukei") Then
            ctl.Visible = True
            ctl.Text = "Column Name"
        End If
        Next ctl
        
    Text6.Text = tableXnm
    Text6.Visible = True
    
    tableXstr = "(" & tableX & ") as " & tableXnm
    
'SQLERR:
'    Exit Sub
End Sub
'��������������������������������������������������������������������������������

Private Sub Command3_Click()
jointflg = "0"
Text6.Visible = False
tableXstr = tableX
'20141125 2����s�ɂ��Ή��@��ŏC����������������������������������������������
Call Form_activate
Call Form_activate
'joint�Ή�������������������������������������������������������������������������
End Sub

Public Sub Form_activate()
On Error GoTo SQLERR
    
    '��rs,cn��Ԋm�F�y�ѕ��鏈��----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        If cn.State <> 0 Then
        cn.Close
        End If
    '��--------------------------------------------------------
            
            DataCombo1.Text = "Table Name"
            
        '��ver1.1 joint�Ή�/20141121 ��������������������������������������������
            If tableXnm <> "" Then
                Label3.Caption = tableXnm
            Else: Label3.Caption = "Saving Nothing"
            End If
            
            If Combo6.Visible = True Then
                For Each ctl In Me.Controls
                If (ctl.Tag = "syukei") Then
                    ctl.Clear
                    ctl.Visible = False
                End If
                Next ctl
            End If

        '������������������������������������������������������������������������
        
        cn.Open cnstr
        'cn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Form1.CommonDialog1.FileName & ";" ver0.1 �c�[
         
        '��DataCombo1�Ƀe�[�u���ꗗ��\������--------------------------
        ssql = "SELECT name FROM sysobjects WHERE xtype = 'u'"
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c1 = rs.Clone
        
            '��rs��Ԋm�F�y�я���----------------------------------------
            If rs.State <> 0 Then
            rs.Close
            End If
            '��----------------------------------------------------------
        
        Set DataCombo1.RowSource = rs_c1
        DataCombo1.ListField = "Name"
        DataCombo1.Refresh
            
        '��------------------------------------------------------------

        '��Combo1�ɃA�C�e���ǉ�-���ׂ������v�C��with��---------------------
        Combo1.AddItem ("sum"): Combo1.AddItem ("max"): Combo1.AddItem ("min"):
        Combo1.AddItem ("avg"): Combo1.AddItem ("count"):
        '��------------------------------------------------------------
        '��Combo2�ɃA�C�e���ǉ�----------------------------------------
        Combo2.AddItem ("sum"): Combo2.AddItem ("max"): Combo2.AddItem ("min"):
        Combo2.AddItem ("avg"): Combo2.AddItem ("count"):
        '��------------------------------------------------------------
        '��Combo3�ɃA�C�e���ǉ�----------------------------------------
        Combo3.AddItem ("sum"): Combo3.AddItem ("max"): Combo3.AddItem ("min"):
        Combo3.AddItem ("avg"): Combo3.AddItem ("count"):
        '��------------------------------------------------------------
        '��Combo4�ɃA�C�e���ǉ�----------------------------------------
        Combo4.AddItem ("sum"): Combo4.AddItem ("max"): Combo4.AddItem ("min"):
        Combo4.AddItem ("avg"): Combo4.AddItem ("count"):
        '��------------------------------------------------------------
        '��Combo5�ɃA�C�e���ǉ�----------------------------------------
        Combo5.AddItem ("sum"): Combo5.AddItem ("max"): Combo5.AddItem ("min"):
        Combo5.AddItem ("avg"): Combo5.AddItem ("count"):
        '��------------------------------------------------------------
SQLERR:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '��rs,cn��Ԋm�F�y�ѕ��鏈��--------------------------------------
    If rs.State <> 0 Then
    rs.Close
    End If
    
    If cn.State <> 0 Then
    cn.Close
    End If
    '��--------------------------------------------------------
    
End Sub


'
'��F5�L�[�������A�N�V����---------------------
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
Private Sub text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command1_Click
    End If
End Sub
Private Sub text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command1_Click
    End If
End Sub
Private Sub text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Call Command1_Click
    End If
End Sub
'��------------------------------------------------------------

Private Sub DataCombo1_LostFocus()
On Error GoTo SQLERR

        '��Form2�փf�[�^���n���p---------------------------------------
        table_n = DataCombo1.Text
        '��------------------------------------------------------------
         
        '��rs��Ԋm�F�y�я���----------------------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        '��----------------------------------------------------------
         
        '��DataCombo2�ɃJ�����ꗗ��\������-���ׂ������v�C��with��-------
        ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo1.Text & "')"
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c2 = rs.Clone
        rs.Close
                
        Set DataCombo2.RowSource = rs_c2: DataCombo2.ListField = "Name"
            DataCombo2.Refresh
        Set DataCombo3.RowSource = rs_c2: DataCombo3.ListField = "Name"
            DataCombo3.Refresh
        Set DataCombo4.RowSource = rs_c2: DataCombo4.ListField = "Name"
            DataCombo4.Refresh
        Set DataCombo5.RowSource = rs_c2: DataCombo5.ListField = "Name"
            DataCombo5.Refresh
        Set DataCombo6.RowSource = rs_c2: DataCombo6.ListField = "Name"
            DataCombo6.Refresh
        Set DataCombo7.RowSource = rs_c2: DataCombo7.ListField = "Name"
            DataCombo7.Refresh
        Set DataCombo8.RowSource = rs_c2: DataCombo8.ListField = "Name"
            DataCombo8.Refresh
        Set DataCombo9.RowSource = rs_c2: DataCombo9.ListField = "Name"
            DataCombo9.Refresh
        Set DataCombo10.RowSource = rs_c2: DataCombo10.ListField = "Name"
            DataCombo10.Refresh
        Set DataCombo11.RowSource = rs_c2: DataCombo11.ListField = "Name"
            DataCombo11.Refresh
'
        '��-------------------------------------------------------------

SQLERR:
    Exit Sub
End Sub

'ver1.1 joint�Ή�/20141121������������������������������������������������������������������������������������������������������������������
Private Sub Command1_Click()
On Error GoTo SQLERR

    '���d���m�F---------------------------------------------------------
    '���ǋL�\�聁������������������������������������������������������
    '��-----------------------------------------------------------------
        If jointflg = "1" Then
            DataCombo1.Text = tableXstr
        End If
        If Text6.Visible = True Then
            DataCombo2.Text = Combo6.Text
            DataCombo3.Text = Combo7.Text
            DataCombo4.Text = Combo8.Text
            DataCombo5.Text = Combo9.Text
            DataCombo6.Text = Combo10.Text
            DataCombo7.Text = Combo11.Text
            DataCombo8.Text = Combo12.Text
            DataCombo9.Text = Combo13.Text
            DataCombo10.Text = Combo14.Text
            DataCombo11.Text = Combo15.Text
        End If
        
    '��sSql�쐬����---------------------------------------------------------
        '���\������------------------------------------------------------------
        If DataCombo2.Text <> "" And DataCombo2.Text <> "Column Name" Then
            cSql1 = "select " & DataCombo2.Text
        End If
        If DataCombo3.Text <> "" And DataCombo3.Text <> "Column Name" Then
            cSql1 = cSql1 & "," & DataCombo3.Text
        End If
        If DataCombo4.Text <> "" And DataCombo4.Text <> "Column Name" Then
            cSql1 = cSql1 & "," & DataCombo4.Text
        End If
        If DataCombo5.Text <> "" And DataCombo5.Text <> "Column Name" Then
            cSql1 = cSql1 & "," & DataCombo5.Text
        End If
        If DataCombo6.Text <> "" And DataCombo6.Text <> "Column Name" Then
            cSql1 = cSql1 & "," & DataCombo6.Text
        End If
        '��------------------------------------------------------------------
        '���W�v����---------------------------------------------------------
        If Text1.Text <> "" And Text1.Text <> " New Column Name 1" Then '���̏���
            cSql1 = cSql1 & ", " & Combo1.Text & "(" & DataCombo7.Text & ") as '" & Text1.Text & "' "
        End If
        If Text2.Text <> "" And Text2.Text <> " New Column Name 2" Then '���̏���
            cSql1 = cSql1 & "," & Combo2.Text & "(" & DataCombo8.Text & ") as '" & Text2.Text & "' "
        End If
        If Text3.Text <> "" And Text3.Text <> " New Column Name 3" Then '���̏���
            cSql1 = cSql1 & "," & Combo3.Text & "(" & DataCombo9.Text & ") as '" & Text3.Text & "' "
        End If
        If Text4.Text <> "" And Text4.Text <> " New Column Name 4" Then '���̏���
            cSql1 = cSql1 & "," & Combo4.Text & "(" & DataCombo10.Text & ") as '" & Text4.Text & "' "
        End If
        If Text5.Text <> "" And Text5.Text <> " New Column Name 5" Then '���̏���
            cSql1 = cSql1 & "," & Combo5.Text & "(" & DataCombo11.Text & ") as '" & Text5.Text & "' "
        End If
        '��------------------------------------------------------------------
        '���e�[�u���ݒ�------------------------------------------------------
        cSql1 = cSql1 & " from " & DataCombo1.Text
        '��------------------------------------------------------------------
        '���W�v��------------------------------------------------------
        If DataCombo2.Text <> "" And DataCombo2.Text <> "Column Name" Then
            cSql1 = cSql1 & " group by " & DataCombo2.Text
        End If
        If DataCombo3.Text <> "" And DataCombo3.Text <> "Column Name" Then
            cSql1 = cSql1 & "," & DataCombo3.Text
        End If
        If DataCombo4.Text <> "" And DataCombo4.Text <> "Column Name" Then
            cSql1 = cSql1 & "," & DataCombo4.Text
        End If
        If DataCombo5.Text <> "" And DataCombo5.Text <> "Column Name" Then
            cSql1 = cSql1 & "," & DataCombo5.Text
        End If
        If DataCombo6.Text <> "" And DataCombo6.Text <> "Column Name" Then
            cSql1 = cSql1 & "," & DataCombo6.Text
        End If
        
        ssql = cSql1
        '��------------------------------------------------------------------
    '��-----------------------------------------------------------------
    
        '��rs��Ԋm�F�y�я���----------------------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        '��----------------------------------------------------------
        '��UI����쐬����SQL��Form2�y��TDBGrid�֓n��--------------------
        pSql = ssql '�󂯓n���p
        jointflg = "0"
        Form2.Text1.Text = ssql
        Nowform = "form3"
        Form2.Show (1)

        '��-------------------------------------------------------------

SQLERR:
    Exit Sub
End Sub
'
'Private Sub Command1_Click()
'On Error GoTo SQLERR
'
'    '���d���m�F---------------------------------------------------------
'    '���ǋL�\�聁������������������������������������������������������
'    '��-----------------------------------------------------------------
'
'    '��sSql�쐬����---------------------------------------------------------
'        '���\������------------------------------------------------------------
'        If DataCombo2.Text <> "" And DataCombo2.Text <> "Column Name" Then
'            cSql1 = "select " & DataCombo2.Text
'        End If
'        If DataCombo3.Text <> "" And DataCombo3.Text <> "Column Name" Then
'            cSql1 = cSql1 & "," & DataCombo3.Text
'        End If
'        If DataCombo4.Text <> "" And DataCombo4.Text <> "Column Name" Then
'            cSql1 = cSql1 & "," & DataCombo4.Text
'        End If
'        If DataCombo5.Text <> "" And DataCombo5.Text <> "Column Name" Then
'            cSql1 = cSql1 & "," & DataCombo5.Text
'        End If
'        If DataCombo6.Text <> "" And DataCombo6.Text <> "Column Name" Then
'            cSql1 = cSql1 & "," & DataCombo6.Text
'        End If
'        '��------------------------------------------------------------------
'        '���W�v����---------------------------------------------------------
'        If Text1.Text <> "" And Text1.Text <> " New Column Name 1" Then '���̏���
'            cSql1 = cSql1 & ", " & Combo1.Text & "(" & DataCombo7.Text & ") as '" & Text1.Text & "' "
'        End If
'        If Text2.Text <> "" And Text2.Text <> " New Column Name 2" Then '���̏���
'            cSql1 = cSql1 & "," & Combo2.Text & "(" & DataCombo8.Text & ") as '" & Text2.Text & "' "
'        End If
'        If Text3.Text <> "" And Text3.Text <> " New Column Name 3" Then '���̏���
'            cSql1 = cSql1 & "," & Combo3.Text & "(" & DataCombo9.Text & ") as '" & Text3.Text & "' "
'        End If
'        If Text4.Text <> "" And Text4.Text <> " New Column Name 4" Then '���̏���
'            cSql1 = cSql1 & "," & Combo4.Text & "(" & DataCombo10.Text & ") as '" & Text4.Text & "' "
'        End If
'        If Text5.Text <> "" And Text5.Text <> " New Column Name 5" Then '���̏���
'            cSql1 = cSql1 & "," & Combo5.Text & "(" & DataCombo11.Text & ") as '" & Text5.Text & "' "
'        End If
'        '��------------------------------------------------------------------
'        '���e�[�u���ݒ�------------------------------------------------------
'        cSql1 = cSql1 & " from " & DataCombo1.Text
'        '��------------------------------------------------------------------
'        '���W�v��------------------------------------------------------
'        If DataCombo2.Text <> "" And DataCombo2.Text <> "Column Name" Then
'            cSql1 = cSql1 & " group by " & DataCombo2.Text
'        End If
'        If DataCombo3.Text <> "" And DataCombo3.Text <> "Column Name" Then
'            cSql1 = cSql1 & "," & DataCombo3.Text
'        End If
'        If DataCombo4.Text <> "" And DataCombo4.Text <> "Column Name" Then
'            cSql1 = cSql1 & "," & DataCombo4.Text
'        End If
'        If DataCombo5.Text <> "" And DataCombo5.Text <> "Column Name" Then
'            cSql1 = cSql1 & "," & DataCombo5.Text
'        End If
'        If DataCombo6.Text <> "" And DataCombo6.Text <> "Column Name" Then
'            cSql1 = cSql1 & "," & DataCombo6.Text
'        End If
'
'        ssql = cSql1
'        '��------------------------------------------------------------------
'    '��-----------------------------------------------------------------
'
'        '��rs��Ԋm�F�y�я���----------------------------------------
'        If rs.State <> 0 Then
'        rs.Close
'        End If
'        '��----------------------------------------------------------
'        '��UI����쐬����SQL��Form2�y��TDBGrid�֓n��--------------------
'        pSql = ssql '�󂯓n���p
'        Form2.Text1.Text = ssql
'        Nowform = "form3"
'        Form2.Show (1)
'
'        '��-------------------------------------------------------------
'
'SQLERR:
'    Exit Sub
'End Sub
'������������������������������������������������������������������������������������������������������������������������������������������������



'���eTextbox�Ɋւ�鋓��--------------------------------------------------------
        '��Text1�Ɋւ�鋓��----------------------------------------------------
Private Sub text1_GotFocus()
        If Text1.Text = " New Column Name 1" Then
            Text1.Text = ""
        End If
End Sub
        '��----------------------------------------------------------------------

        '��Text2�Ɋւ�鋓��----------------------------------------------------
Private Sub Text2_Gotfocus()
        If Text2.Text = " New Column Name 2" Then
            Text2.Text = ""
        End If
End Sub
        '��----------------------------------------------------------------------

        '��Text3�Ɋւ�鋓��----------------------------------------------------
Private Sub text3_GotFocus()
        If Text3.Text = " New Column Name 3" Then
            Text3.Text = ""
        End If
End Sub
        '��----------------------------------------------------------------------

        '��Text4�Ɋւ�鋓��----------------------------------------------------
Private Sub text4_GotFocus()
        If Text4.Text = " New Column Name 4" Then
            Text4.Text = ""
        End If
End Sub
        '��----------------------------------------------------------------------
        
        '��Text5�Ɋւ�鋓��----------------------------------------------------
Private Sub text5_GotFocus()
        If Text5.Text = " New Column Name 5" Then
            Text5.Text = ""
        End If
End Sub
        '��----------------------------------------------------------------------
'��------------------------------------------------------------------------------
