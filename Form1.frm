VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   Appearance      =   0  '�ׯ�
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "Search Menu"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "���C���I"
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
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   6345
   StartUpPosition =   2  '��ʂ̒���
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
      Text            =   "�O��"
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '�ׯ�
      Height          =   420
      Left            =   2640
      TabIndex        =   2
      Text            =   "="
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "���C���I"
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
         Name            =   "���C���I"
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
         Name            =   "���C���I"
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
         Name            =   "���C���I"
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
      BackStyle       =   0  '����
      Caption         =   "Search Menu"
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
      TabIndex        =   8
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "�@�`"
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

'ver1.1joint�Ή�/20141121������������������������������������������������������
Private Sub Command2_Click()
On Error GoTo SQLERR
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
        
        Combo3.Clear
        
    '��rs,cn��Ԋm�F�y�ѕ��鏈��----------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        If cn.State <> 0 Then
        cn.Close
        End If
    '��--------------------------------------------------------

        cn.Open cnstr
                
        '��DataCombo2�ɃJ�����ꗗ��\������--------------------------
        ssql = tableX
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c1 = rs.Clone 'rs���g���񂷈ׁA�����_�̃N���[����tmp�ɍ쐬
  
        For i = 0 To rs.Fields.Count - 1
            Combo3.AddItem rs.Fields(i).Name
        Next i
        
        '��------------------------------------------------------------
                    
            '��rs��Ԋm�F�y�я���----------------------------------------
            If rs.State <> 0 Then
            rs.Close
            End If
            '��----------------------------------------------------------
            
    Combo3.Visible = True
    Text3.Text = tableXnm
    Text3.Visible = True
    
    tableXstr = "(" & tableX & ") as " & tableXnm
SQLERR:
    Exit Sub
End Sub
'��������������������������������������������������������������������������������

Private Sub Command3_Click()
jointflg = "0"
tableXstr = tableX
Call Form_activate
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
        Combo1.Clear: Combo1.Text = "="
        Text1.Text = " Keyword1"
        Text2.Text = " Keyword2": Text2.Visible = False
        Combo2.Clear: Combo2.Text = �O��: Combo2.Visible = False
        
        cn.Open cnstr
        
        '��ver1.1 joint�Ή�/20141121 ��������������������������������������������
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
        
        '������������������������������������������������������������������������
        
        '��DataCombo2�ɃJ�����ꗗ��\������----------------------------
        If DataCombo1.Text <> "" And DataCombo1.Text <> "Table Name" Then
            ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo1.Text & "')"
            rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
            Set rs_c2 = rs.Clone
            
            '��rs��Ԋm�F�y�я���----------------------------------------
            If rs.State <> 0 Then
            rs.Close
            End If
            '��----------------------------------------------------------
                    
            Set DataCombo2.RowSource = rs_c2
            DataCombo2.ListField = "Name"
            DataCombo2.Refresh
        End If
        
        '��-------------------------------------------------------------
        '��DataCombo1�Ƀe�[�u���ꗗ��\������--------------------------
        ssql = "SELECT name FROM sysobjects WHERE xtype = 'u'"
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c1 = rs.Clone 'rs���g���񂷈ׁA�����_�̃N���[����tmp�ɍ쐬
        
        '��rs��Ԋm�F�y�я���----------------------------------------
        If rs.State <> 0 Then
        rs.Close
        End If
        '��----------------------------------------------------------
        
        Set DataCombo1.RowSource = rs_c1
        DataCombo1.ListField = "Name"
        DataCombo1.Refresh
        
        '��------------------------------------------------------------
        
        '��Combo1�ɃA�C�e���ǉ�----------------------------------------
        Combo1.AddItem ("="): Combo1.AddItem ("<"): Combo1.AddItem (">"):
        Combo1.AddItem ("<="): Combo1.AddItem (">="): Combo1.AddItem ("<>"):
        Combo1.AddItem ("bet"): Combo1.AddItem ("like")
        '��------------------------------------------------------------
        '��Combo2�ɃA�C�e���ǉ�----------------------------------------
        Combo2.AddItem ("�O��"): Combo2.AddItem ("���"): Combo2.AddItem ("����"):
        '��------------------------------------------------------------
SQLERR:
    Exit Sub
End Sub

'���t�H�[���I����-------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)

    '��rs,cn��Ԋm�F�y�ѕ��鏈��-----------------------------
    If rs.State <> 0 Then
    rs.Close
    End If
    
    If cn.State <> 0 Then
    cn.Close
    End If
    '��--------------------------------------------------------
    
    ssql = ""
    pSql = ""
    
End Sub
'��----------------------------------------------------------------------


'��F5�L�[�������A�N�V����---------------------
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
         
        '��DataCombo2�ɃJ�����ꗗ��\������----------------------------
        ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo1.Text & "')"
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c2 = rs.Clone
        
            '��rs��Ԋm�F�y�я���----------------------------------------
            If rs.State <> 0 Then
            rs.Close
            End If
            '��----------------------------------------------------------
                
        Set DataCombo2.RowSource = rs_c2
        DataCombo2.ListField = "Name"
        DataCombo2.Refresh
        
        '��-------------------------------------------------------------
SQLERR:
    Exit Sub
End Sub
        '������������e�I�u�W�F�N�g�\���ύX----------------------------
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
        '��-------------------------------------------------------------
        
'ver1.1 joint�Ή�/20141121������������������������������������������������������������������������������������������������������������������
Private Sub Command1_Click()
On Error GoTo SQLERR
        
        If Text3.Visible = True Then
            DataCombo1.Text = tableXstr
            DataCombo2.Text = Combo3.Text
        End If
        
        '�������������w�肷�邩�ۂ�-------------------------------------
        If Combo1.Text = "null" Then
                ssql = "Select * From " & DataCombo1.Text & " where " & DataCombo2.Text & " is null"
        ElseIf Text1.Text = "" Or Text1.Text = " KeyWord1" Or DataCombo2.Text = "Column Name" Or DataCombo2.Text = "" Then
'        '�������������̓J�����I�����Ȃ���Ă��Ȃ��ꍇ�e�[�u����\�����邾����SQL���s
                ssql = "Select * From " & DataCombo1.Text
        ElseIf Combo1.Text = "like" Then
            If Combo2.Text = "�O��" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '" & Text1.Text & "%'"
            ElseIf Combo2.Text = "���" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '%" & Text1.Text & "'"
            ElseIf Combo2.Text = "����" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '%" & Text1.Text & "%'"
            End If
        ElseIf Combo1.Text <> "=" And Combo1.Text <> "bet" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '" & Text1.Text & "'"
        ElseIf Combo1.Text = "=" Then
                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " = '" & Text1.Text & "'"
        ElseIf Combo1.Text = "bet" Then
                ssql = "Select * From " & DataCombo1.Text & " where " & DataCombo2.Text & " between '" & Text1.Text & "' and '" & Text2.Text & "'"
        End If
        
        '��-------------------------------------------------------------
        '��rs��Ԋm�F�y�я���----------------------------------------
        If rs.State <> 0 Then
        rs.Close
                                                                                                                                                                                         End If
        '��----------------------------------------------------------
        '��UI����쐬����SQL��Form2�y��TDBGrid�֓n��--------------------
        jointflg = "0"
        pSql = ssql '�󂯓n���p
        Form2.Text1.Text = ssql
        Nowform = "form1"
        Form2.Show (1)
        '��-------------------------------------------------------------


'Private Sub Command1_Click()
'On Error GoTo SQLERR
'
'        '�������������w�肷�邩�ۂ�-------------------------------------
'        If Text1.Text = "" Or Text1.Text = " KeyWord1" Or DataCombo2.Text = "Column Name" Or DataCombo2.Text = "" Then
'        '�������������̓J�����I�����Ȃ���Ă��Ȃ��ꍇ�e�[�u����\�����邾����SQL���s
'                ssql = "Select * From " & DataCombo1.Text
'        ElseIf Combo1.Text = "like" Then
'            If Combo2.Text = "�O��" Then
'                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '" & Text1.Text & "%'"
'            ElseIf Combo2.Text = "���" Then
'                ssql = "Select * From " & DataCombo1.Text & " Where " & DataCombo2.Text & " " & Combo1.Text & " '%" & Text1.Text & "'"
'            ElseIf Combo2.Text = "����" Then
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
'        '��-------------------------------------------------------------
'        '��rs��Ԋm�F�y�я���----------------------------------------
'        If rs.State <> 0 Then
'        rs.Close
'        End If
'        '��----------------------------------------------------------
'        '��UI����쐬����SQL��Form2�y��TDBGrid�֓n��--------------------
'        pSql = ssql '�󂯓n���p
'        Form2.Text1.Text = ssql
'        Nowform = "form1"
'        Form2.Show (1)
'        '��-------------------------------------------------------------
'SQLERR:
'    Exit Sub
'End Sub

SQLERR:
    Exit Sub
End Sub
'������������������������������������������������������������������������������������������������������������������������������������������������


'���eTextbox�Ɋւ�鋓��--------------------------------------------------------
        '��Text1�Ɋւ�鋓��----------------------------------------------------
Private Sub text1_GotFocus()
        If Text1.Text = " Keyword1" Then
            Text1.Text = ""
        End If
End Sub
        '��----------------------------------------------------------------------

        '��Text2�Ɋւ�鋓��----------------------------------------------------
Private Sub Text2_Gotfocus()
        If Text2.Text = " Keyword2" Then
            Text2.Text = ""
        End If
End Sub
        '��----------------------------------------------------------------------

