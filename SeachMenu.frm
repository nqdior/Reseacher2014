VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form2 
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
   Icon            =   "SeachMenu.frx":0000
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
      Left            =   480
      TabIndex        =   6
      Top             =   3720
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oControl As Control

Public Sub Form_activate()
On Error GoTo SQLERR

    Call CNclose
    Call RSclose
    jointflg = "0"
    
'���t�H�[�����[�h���eOBJ�ݒ�----------------------------

    '���e������������Ԃ�----------------------
    
    Combo1.Clear: Combo1.Text = "=": DataCombo2.Text = "Column Name"
    Text1.Text = " Keyword1": Text2.Text = " Keyword2"
    Text2.Visible = False: Label1.Visible = False
    Combo2.Clear: Combo2.Text = �O��: Combo2.Visible = False
    
    '��----------------------------------------
    
    '��Joint�󋵊m�F�E�Ή�---------------------
    
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
    
    '��----------------------------------------
    
        
    '��DataCombo1�Ƀe�[�u���ꗗ��\������------
    
    cn.Open cnstr
    ssql = "SELECT name FROM sysobjects WHERE xtype = 'u'"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c1 = rs.Clone
    
    Call RSclose
    
    With DataCombo1
        Set .RowSource = rs_c1: .ListField = "Name": .Refresh: .Text = "Table Name"
    End With
    
    '��-----------------------------------------
    
    
    '��DataCombo2�ɃJ�����ꗗ��\������---------
    
    If DataCombo1.Text <> "" And DataCombo1.Text <> "Table Name" Then
        ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo1.Text & "')"
        rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c2 = rs.Clone
                
        Call RSclose
                
        With DataCombo2
            Set .RowSource = rs_c2: .ListField = "Name": .Refresh
        End With
    End If
    
    '��------------------------------------------
    
    
    '��Combo1�ɃA�C�e���ǉ�----------------------
    
    With Combo1
        .AddItem ("="): .AddItem ("<"): .AddItem (">"):
        .AddItem ("<="): .AddItem (">="): .AddItem ("<>"):
        .AddItem ("bet"): .AddItem ("like"): .AddItem ("null")
    End With
    
    '��------------------------------------------
    
    
    '��Combo2�ɃA�C�e���ǉ�----------------------
    
    With Combo2
        .AddItem ("�O��"): .AddItem ("���"): .AddItem ("����"):
    End With
    
    '��------------------------------------------

'��------------------------------------------------------

SQLERR:
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer) '�t�H�[���I����
    
    Call RSclose: Call CNclose
    ssql = "": pSql = ""
    
End Sub '-------------------------------------------------


Private Sub Command1_Click() '���s�{�^��������
On Error GoTo SQLERR

    'Joint�Ή�------------------------
    If Text3.Visible = True Then
        DataCombo1.Text = tableXstr
        DataCombo2.Text = Combo3.Text
    End If
    '---------------------------------
    
    
'��SQL���쐬����---------------------------------------------------------

    '�������������w�肷�邩�ۂ�-------------------------------------
    
    If Combo1.Text = "null" Then
            ssql = "Select * From " & DataCombo1.Text & " where " & DataCombo2.Text & " is null"
            
    ElseIf Combo1.Text = "Nnull" Then
            ssql = "Select * From " & DataCombo1.Text & " where " & DataCombo2.Text & " is not null"
            
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

'��----------------------------------------------------------------------

    Call RSclose
    
    '��UI����쐬����SQL��EFORM�y��TDBGrid�֓n��--------------------
    jointflg = "0"
    pSql = ssql '�󂯓n���p
    EForm.Text1.Text = ssql
    Nowform = "form2"
    EForm.Show (1)
    '��-------------------------------------------------------------

SQLERR:
    Exit Sub
End Sub
'������������������������������������������������������������������������������������������������������������������������������������������������

Private Sub Command2_Click() 'Table Call
On Error GoTo SQLERR
    
    '��TableX �L������-------------------------------
    If tableX = "" Then
        MsgBox ("�ێ�����Ă���e�[�u��������܂���")
        Exit Sub
    End If
    '��----------------------------------------------
    
    Call TableCall
    Combo3.Clear: Combo3.Text = "Column Name"
    
    Call RSclose
    Call CNclose

'��Form�ݒ�-----------------------------------------------------------

    cn.Open cnstr
          
    Combo3.Visible = True
    Text3.Text = tableXnm
    Text3.Visible = True
    
    '��DataCombo2�ɃJ�����ꗗ��\������--------------------------
    
    ssql = tableX
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c1 = rs.Clone 'rs���g���񂷈ׁA�����_�̃N���[����tmp�ɍ쐬
    
    For i = 0 To rs.Fields.Count - 1
    Combo3.AddItem rs.Fields(i).Name
    Next i
    
    '��----------------------------------------------------------
'��--------------------------------------------------------------------
        
    Call RSclose
    tableXstr = "(" & tableX & ") as " & tableXnm

SQLERR:
    Exit Sub
End Sub
'��������������������������������������������������������������������������������

Private Sub Command3_Click()
    
    '���ǂݍ��݃e�[�u������--
    If jointflg = "0" Then
        MsgBox "�ǂݍ��ݍσe�[�u��������܂���"
        Exit Sub
    End If
    '��----------------------
    
    jointflg = "0"
    tableXstr = tableX
    Call Form_activate
End Sub


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


Private Sub DataCombo1_Lostfocus() '�e�[�u���I���㏈��
On Error GoTo SQLERR
         
    Call RSclose

    '��DataCombo2�ɃJ�����ꗗ��\������------------------------
    
    ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo1.Text & "')"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c2 = rs.Clone
    
    Call RSclose
    
    Set DataCombo2.RowSource = rs_c2
    DataCombo2.ListField = "Name"
    DataCombo2.Refresh

    '��--------------------------------------------------------
    
SQLERR:
    Exit Sub
End Sub '-------------------------------------------------------
    
Private Sub Combo1_LostFocus() '���������ݒ��eOBJ�ݒ�
        
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
        
End Sub '-----------------------------------------------
     

'���eTextbox�Ɋւ�鋓��--------------------------------
        
    '��Text1�Ɋւ�鋓��----------------------------

Private Sub text1_GotFocus()
    
    If Text1.Text = " Keyword1" Then
        Text1.Text = ""
    End If
    
End Sub

    '��---------------------------------------------


    '��Text2�Ɋւ�鋓��----------------------------

Private Sub Text2_Gotfocus()
    
    If Text2.Text = " Keyword2" Then
        Text2.Text = ""
    End If

End Sub
    
    '��---------------------------------------------
'��------------------------------------------------------
