VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Join Menu"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   450
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
   LinkTopic       =   "Form5"
   ScaleHeight     =   9870
   ScaleWidth      =   6390
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   34
      Top             =   6360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   15724527
      HighLight       =   0
      Appearance      =   0
   End
   Begin VB.ComboBox Combo2 
      Height          =   390
      Left            =   2280
      TabIndex        =   32
      Text            =   "LEFT"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      Left            =   360
      TabIndex        =   31
      Text            =   "INNER JOIN"
      Top             =   5280
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo DataCombo12 
      Height          =   390
      Left            =   3600
      TabIndex        =   12
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   688
      _Version        =   393216
      Text            =   "Table Name"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Join Key 2"
      Height          =   3375
      Left            =   3480
      TabIndex        =   24
      Top             =   1440
      Width           =   2775
      Begin MSDataListLib.DataCombo DataCombo7 
         Height          =   390
         Left            =   240
         TabIndex        =   13
         Tag             =   "dcombo"
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
         TabIndex        =   14
         Tag             =   "dcombo"
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
         TabIndex        =   15
         Tag             =   "dcombo"
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
         TabIndex        =   16
         Tag             =   "dcombo"
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
         TabIndex        =   17
         Tag             =   "dcombo"
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
      Begin VB.ComboBox Combo11 
         Height          =   390
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo12 
         Height          =   390
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo13 
         Height          =   390
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo14 
         Height          =   390
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo15 
         Height          =   390
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reseach (F5)"
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Join Key 1"
      Height          =   3375
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   2775
      Begin VB.ComboBox Combo10 
         Height          =   390
         Left            =   240
         TabIndex        =   11
         Tag             =   "syukei"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo9 
         Height          =   390
         Left            =   240
         TabIndex        =   10
         Tag             =   "syukei"
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo8 
         Height          =   390
         Left            =   240
         TabIndex        =   9
         Tag             =   "syukei"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo6 
         Height          =   390
         Left            =   240
         TabIndex        =   7
         Tag             =   "syukei"
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo7 
         Height          =   390
         Left            =   240
         TabIndex        =   8
         Tag             =   "syukei"
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   390
         Left            =   240
         TabIndex        =   1
         Tag             =   "dcombo"
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
         TabIndex        =   2
         Tag             =   "dcombo"
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
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   390
         Left            =   240
         TabIndex        =   3
         Tag             =   "dcombo"
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
      Begin MSDataListLib.DataCombo DataCombo5 
         Height          =   390
         Left            =   240
         TabIndex        =   4
         Tag             =   "dcombo"
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
      Begin MSDataListLib.DataCombo DataCombo6 
         Height          =   390
         Left            =   240
         TabIndex        =   5
         Tag             =   "dcombo"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SavingTable"
      Height          =   1455
      Left            =   3120
      TabIndex        =   21
      Top             =   8280
      Width           =   3135
      Begin VB.CommandButton Command3 
         Caption         =   "ClearTable"
         Height          =   495
         Left            =   1560
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "TableCall"
         Height          =   495
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   405
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   840
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
   Begin VB.Frame Frame4 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      Caption         =   "Join Type"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   33
      Top             =   4920
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1575
      Left            =   3240
      TabIndex        =   35
      Top             =   6360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   15724527
      HighLight       =   0
      Appearance      =   0
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Column"
      Height          =   2055
      Left            =   120
      TabIndex        =   36
      Top             =   6000
      Width           =   6135
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   2640
      X2              =   4440
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   2520
      X2              =   4320
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   2640
      X2              =   4440
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   2520
      X2              =   4320
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   2640
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "Join Menu"
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
      TabIndex        =   30
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_activate()
On Error GoTo SQLERR

    Form1.MousePointer = 11

    Call CNclose
    Call RSclose
    
'���e�I�u�W�F�N�g���[�h�ݒ�--------------------------

    '��FormActive���eOBJ���Z�b�g--------------
    jointflg = "0"
    DataCombo1.Text = "Table Name"
    DataCombo12.Text = "Table Name"
    
    With Combo1
        .Clear: .Text = "INNER JOIN"
    End With
    With Combo2
        .Clear: .Text = "LEFT"
    End With

    With MSFlexGrid1
        .ColWidth(0) = 3000: .Rows = 1: .Clear
    End With
    With MSFlexGrid2
        .ColWidth(0) = 3000: .Rows = 1: .Clear
    End With
    
    'DBcombo1�`11---------------
    For Each ctl In Me.Controls
    If (ctl.Tag = "dcombo") Then
        ctl.Text = "Column Name"
    End If
    Next ctl
    '---------------------------
    '��----------------------------------------
    
    '��Joint�󋵊m�F�E�Ή�---------------------
    If tableXnm <> "" Then
        Label3.Caption = tableXnm
    Else
        Label3.Caption = "Saving Nothing"
    End If
    
    If Combo6.Visible = True Then
        For Each ctl In Me.Controls
            If (ctl.Tag = "syukei") Then
                ctl.Visible = False
            End If
        Next ctl
    End If
    '��----------------------------------------
    
    '��Combobox�ݒ�---------------------------------------------
    With Combo1
        .AddItem ("INNER JOIN"): .AddItem ("OUTER JOIN")
    End With

    With Combo2
        .AddItem ("LEFT"): .AddItem ("RIGHT"): .AddItem ("FULL")
    End With
    '��---------------------------------------------------------

    '��DataCombo1�Ƀe�[�u���ꗗ��\������-----------------------
    cn.Open cnstr
     
    ssql = "SELECT name FROM sysobjects WHERE xtype = 'u'"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c1 = rs.Clone
    
    Call RSclose
        
    Set DataCombo1.RowSource = rs_c1
    DataCombo1.ListField = "Name"
    DataCombo1.Refresh
    Set DataCombo12.RowSource = rs_c1
    DataCombo12.ListField = "Name"
    DataCombo12.Refresh
    '��---------------------------------------------------------
'��--------------------------------------------------
        
    Form1.MousePointer = 0

SQLERR:
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Call RSclose: Call CNclose
    ssql = "": pSql = ""
End Sub


Private Sub DataCombo1_Lostfocus()
On Error GoTo SQLERR
    
    Call RSclose
     
'���ǂݍ��񂾃e�[�u���̍��ږ��ꗗ��\��------------------------

    '��DataCombo2�`6�ɃJ�����ꗗ��\������-------------
    
    ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo1.Text & "')"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c2 = rs.Clone
    Call RSclose

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

    '��jointflg����-------------------------------------
    
    If jointflg = "0" Then 'joint����ĂȂ����ǂ���
        ssql = "select * from " & DataCombo1.Text
    Else
        ssql = tableX
    End If
    
    '��-------------------------------------------------
    
    '��Grid1�ɃJ�����ꗗ��\������----------------------
    
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    With MSFlexGrid1
        .ColWidth(0) = 3000: .Rows = 1: .Clear
    End With
    
    For i = rs.Fields.Count - 1 To 0 Step -1
        MSFlexGrid1.Rows = rs.Fields.Count - 1
        MSFlexGrid1.AddItem rs.Fields(i).Name, MSFlexGrid1.Row
        MSFlexGrid1.Refresh
    Next i
    
    MSFlexGrid1.ColWidth(0) = 5000
    Call RSclose
    
    '�����ږ����f��rs2��W�J�s�v�̉\���A�v�C��--------
'��----------------------------------------------------------------

SQLERR:
    Exit Sub
End Sub

Private Sub DataCombo12_Lostfocus()
On Error GoTo SQLERR
    
    Call RSclose
     
'���ǂݍ��񂾃e�[�u���̍��ږ��ꗗ��\��------------------------

    '��DataCombo2�`6�ɃJ�����ꗗ��\������-------------
    
    ssql = "SELECT name FROM syscolumns WHERE id = object_id('" & DataCombo12.Text & "')"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c7 = rs.Clone
    Call RSclose

    Set DataCombo7.RowSource = rs_c7: DataCombo7.ListField = "Name":
        DataCombo7.Refresh
    Set DataCombo8.RowSource = rs_c7: DataCombo8.ListField = "Name"
        DataCombo8.Refresh
    Set DataCombo9.RowSource = rs_c7: DataCombo9.ListField = "Name"
        DataCombo9.Refresh
    Set DataCombo10.RowSource = rs_c7: DataCombo10.ListField = "Name"
        DataCombo10.Refresh
    Set DataCombo11.RowSource = rs_c7: DataCombo11.ListField = "Name"
        DataCombo11.Refresh

'    '��jointflg����-------------------------------------
'
'    If jointflg = "0" Then 'joint����ĂȂ����ǂ���
        ssql = "select * from " & DataCombo12.Text
'    Else
'        ssql = tableX
'    End If
'
'    '��-------------------------------------------------
    
    '��Grid2�ɃJ�����ꗗ��\������----------------------
    
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    With MSFlexGrid2
        .ColWidth(0) = 3000: .Rows = 1: .Clear
    End With
    
    For i = rs.Fields.Count - 1 To 0 Step -1
        MSFlexGrid2.Rows = rs.Fields.Count - 1
        MSFlexGrid2.AddItem rs.Fields(i).Name, MSFlexGrid2.Row
        MSFlexGrid2.Refresh
    Next i
    
    MSFlexGrid2.ColWidth(0) = 5000
    Call RSclose
    
    '�����ږ����f��rs2��W�J�s�v�̉\���A�v�C��--------
'��----------------------------------------------------------------
         
SQLERR:
    Exit Sub
End Sub


Private Sub Combo1_Click() 'Join��ޑI����OUTER�t�������\��

    If Combo1.Text = "OUTER JOIN" Then
        Combo2.Visible = True
        Else: Combo2.Visible = False
    End If
    
End Sub '---------------------------------------------------


Private Sub MSFlexGrid1_Click() '�I���s�ɑ΂���C�x���g------------------------
On Error GoTo SQLERR

'���I���s�𔻒�A������ݒ�/�ݒ�ς݂Ȃ�߂�----------------

    Wr = MSFlexGrid1.RowSel '�I���s
    
    MSFlexGrid1.Row = Wr '�N���b�N�ʒuRow
    MSFlexGrid1.Col = 0 '�N���b�N�ʒuCol
    If MSFlexGrid1.CellFontBold = False Then
        MSFlexGrid1.CellFontBold = True
    Else '���ɐݒ�ς݂̏ꍇ
        MSFlexGrid1.CellFontBold = False
    End If
    
'��---------------------------------------------------------

SQLERR:
    Exit Sub
End Sub

Private Sub MSFlexGrid2_Click()
On Error GoTo SQLERR

'���I���s�𔻒�A������ݒ�/�ݒ�ς݂Ȃ�߂�----------------

    Wr = MSFlexGrid2.RowSel '�I���s
    
    MSFlexGrid2.Row = Wr '�N���b�N�ʒuRow
    MSFlexGrid2.Col = 0 '�N���b�N�ʒuCol
    If MSFlexGrid2.CellFontBold = False Then
        MSFlexGrid2.CellFontBold = True
    Else '���ɐݒ�ς݂̏ꍇ
        MSFlexGrid2.CellFontBold = False
    End If
    
'��---------------------------------------------------------
    
SQLERR:
    Exit Sub
End Sub


Private Sub Command1_Click()
On Error GoTo SQLERR

'���d���m�F---------------------------------------------------------
'   �ǋL�\�聁������������������������������������������������������
'��-----------------------------------------------------------------
    
'���e�[�u���E�J�����EJOIN�����ݒ�-------------------------------------------------
    
    '��Joint���s���Ă���ꍇ�ADatacombo1�ɃG�C���A�X���i�[
    
    If jointflg = "1" Then
        DataCombo1.Text = tableXnm
    End If
    
    '��----------------------------------------------------
    
    
    '��GRID1�̑I�����ڔ���Eselcol1�Ɋi�[------------------
    
    For i = 0 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 0
    If MSFlexGrid1.CellFontBold = True Then
        If selcol1 = "" Then
            selcol1 = DataCombo1.Text & "." & MSFlexGrid1.Text
        Else
            selcol1 = selcol1 & "," & DataCombo1.Text & "." & MSFlexGrid1.Text
        End If
    End If
    Next i
    
    '��----------------------------------------------------
    
    
    '��GRID2�̑I�����ڔ���Eselcol2�Ɋi�[------------------
    
    For i = 0 To MSFlexGrid2.Rows - 1
    MSFlexGrid2.Row = i
    MSFlexGrid2.Col = 0
    If MSFlexGrid2.CellFontBold = True Then
        If selcol2 = "" Then
            selcol2 = DataCombo12.Text & "." & MSFlexGrid2.Text
        Else
            selcol2 = selcol2 & "," & DataCombo12.Text & "." & MSFlexGrid2.Text
        End If
    End If
    Next i
    
    '��----------------------------------------------------
    
    '��selcol1,2�̑I��L������-----------------------------
    
    If selcol1 = "" Then
        MsgBox "�\�����ڂ��I������Ă��܂���B"
        selcol1 = "": selcol2 = ""
        Exit Sub
    ElseIf selcol2 = "" Then
        MsgBox "�\�����ڂ��I������Ă��܂���B"
        selcol1 = "": selcol2 = ""
        Exit Sub
    Else '���Ȃ��Ȃ��selcol�ɑo���i�[
        selcol = selcol1 & "," & selcol2
        selcol1 = "": selcol2 = ""
    End If
    
    '��----------------------------------------------------
'��---------------------------------------------------------------

'��sSql�쐬����---------------------------------------------------

    '��Joint����Ă��邩�ۂ�����---------------------------
    'Joint����Ă���Ȃ�e�[�u��������TableX�i�[
    '�e�[�u���������ɃG�C���A�X���i�[
    
    If jointflg = "1" Then
        DataCombo1.Text = tableXstr
        jointbl = tableXnm
        
    'Joint���ꂽ�̂�Join�����ۂ�
        tmp = InStr(1, tableX, "JOIN")
        If tmp = 0 Then
            cSql1 = "SELECT " & selcol & " FROM " & DataCombo1.Text
        Else
            cSql1 = "SELECT " & selcol & " FROM " & tableXstr
        End If
    Else
        jointbl = DataCombo1.Text
        cSql1 = "SELECT " & selcol & " FROM " & DataCombo1.Text
    End If
    
    '���킩��h���v�C��-------------------------------------
    
    '��Joint����AText6���\�����̏ꍇ-----------------------
    
    If Text6.Visible = True Then
        DataCombo2.Text = Combo6.Text
        DataCombo3.Text = Combo7.Text
        DataCombo4.Text = Combo8.Text
        DataCombo5.Text = Combo9.Text
        DataCombo6.Text = Combo10.Text
    End If
    
    '��-----------------------------------------------------
    
    '��������ނɂ�蕪��-----------------------------------
    
    If Combo1.Text = "INNER JOIN" Then
            cSql1 = cSql1 & " " & Combo1.Text & " " & DataCombo12.Text & " ON (" _
            & jointbl & "." & DataCombo2.Text & " = " & DataCombo12.Text & "." & DataCombo7.Text
            
    ElseIf Combo1.Text = "OUTER JOIN" Then
            cSql1 = cSql1 & " " & Combo2.Text & " " & Combo1.Text & " " & DataCombo12.Text & " ON (" _
            & jointbl & "." & DataCombo2.Text & " = " & DataCombo12.Text & "." & DataCombo7.Text
    Else
            MsgBox "JOIN�������w�肵�Ă��������B(INNER or OUTER)"
            Exit Sub
    End If
    '��------------------------------------------------------
                    
    '��2�`5��Join�L�[�����邩�ǂ���--------------------------
    If DataCombo3.Text <> "" And DataCombo3.Text <> "Column Name" Or DataCombo8.Text <> "" And DataCombo8.Text <> "Column Name" Then
        cSql1 = cSql1 & " AND " & jointbl & "." & DataCombo3.Text & " = " & DataCombo12.Text & "." & DataCombo8.Text
    End If
    If DataCombo4.Text <> "" And DataCombo4.Text <> "Column Name" Or DataCombo9.Text <> "" And DataCombo9.Text <> "Column Name" Then
        cSql1 = cSql1 & " AND " & jointbl & "." & DataCombo4.Text & " = " & DataCombo12.Text & "." & DataCombo9.Text
    End If
    If DataCombo5.Text <> "" And DataCombo5.Text <> "Column Name" Or DataCombo10.Text <> "" And DataCombo10.Text <> "Column Name" Then
        cSql1 = cSql1 & " AND " & jointbl & "." & DataCombo5.Text & " = " & DataCombo12.Text & "." & DataCombo10.Text
    End If
    If DataCombo6.Text <> "" And DataCombo6.Text <> "Column Name" Or DataCombo11.Text <> "" And DataCombo11.Text <> "Column Name" Then
        cSql1 = cSql1 & " AND " & jointbl & "." & DataCombo6.Text & " = " & DataCombo12.Text & "." & DataCombo11.Text
    End If
    '��------------------------------------------------------
    
    cSql1 = cSql1 & ")" 'JOINjoint�Ή�
    ssql = cSql1
    
    '��------------------------------------------------------
'��-----------------------------------------------------------------

    Call RSclose

    '��UI����쐬����SQL��EFORM�y��TDBGrid�֓n��-------------
    
    pSql = ssql '�󂯓n���p
    jointflg = "0"
    EForm.Text1.Text = ssql
    Nowform = "form1"
    EForm.Show (1)

    '��------------------------------------------------------

SQLERR:
    Exit Sub
End Sub


Private Sub Command2_Click() 'TableX Call
On Error GoTo SQLERR

    '��TableX �L������-------------------------------
    If tableX = "" Then
        MsgBox ("�ێ�����Ă���e�[�u��������܂���")
        Exit Sub
    End If
    '��----------------------------------------------
    
    Call TableCall

'���e�I�u�W�F�N�g�ݒ�-------------------------------------
    
    'obj���Z�b�g--------
    With MSFlexGrid1
        .ColWidth(0) = 3000: .Rows = 1: .Clear
    End With
'    MSFlexGrid2.Clear������
    Combo6.Clear: Combo7.Clear: Combo8.Clear: Combo9.Clear: Combo10.Clear:
'    Combo11.Clear: Combo12.Clear: Combo13.Clear: Combo14.Clear: Combo15.Clear:������
    '----------------
    
    Call RSclose
    Call CNclose

    '��datacombo�ݒ�---------------------------------��������

    cn.Open cnstr
    ssql = tableX
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c1 = rs.Clone

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
    '��-----------------------------------------------

    '��grid1�Ƀf�[�^��\������------------------------
    For i = rs.Fields.Count - 1 To 0 Step -1
        With MSFlexGrid1
            .Rows = rs.Fields.Count - 1
            .AddItem rs.Fields(i).Name, MSFlexGrid1.Row
            .Refresh
        End With
    Next i
    MSFlexGrid1.ColWidth(0) = 5000
    Call RSclose
    '��-----------------------------------------------


    '��datacombo12�Ƀf�[�^��\������------------------����������
    ssql = "SELECT name FROM sysobjects WHERE xtype = 'u'"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
        Set rs_c7 = rs.Clone
    Set DataCombo12.RowSource = rs_c7
        DataCombo12.ListField = "Name"
        DataCombo12.Refresh
    Call RSclose
    '��-----------------------------------------------
    
    '��combo1�`11�̕\���ݒ�---------------------------
    For Each ctl In Me.Controls
        If (ctl.Tag = "syukei") Then
            ctl.Visible = True
            ctl.Text = "Column Name"
        End If
    Next ctl
    '��-----------------------------------------------
'��---------------------------------------------------------
        
    Text6.Text = tableXnm
    Text6.Visible = True

    tableXstr = "(" & tableX & ") as " & tableXnm
    Combo6.SetFocus
    
SQLERR:
    Exit Sub
End Sub


Private Sub Command3_Click() 'TableX clear
    
    '���ǂݍ��݃e�[�u������--
    If jointflg = "0" Then
        MsgBox "�ǂݍ��ݍσe�[�u��������܂���"
        Exit Sub
    End If
    '��----------------------
    
    
'���e�I�u�W�F�N�g���[�h�ݒ�--------------------------

    '����ԏ�����(Joint����)-----------------
    
    jointflg = "0"
    Text6.Visible = False
    tableXstr = tableX

    Call CNclose
    Call RSclose
    
    '��---------------------------------------

    '��FormActive���eOBJ���Z�b�g--------------
    
    DataCombo1.Text = "Table Name"
    
    With MSFlexGrid1
        .ColWidth(0) = 3000: .Rows = 1: .Clear
    End With
    
    'DBcombo1�`11---------------
    For Each ctl In Me.Controls
    If (ctl.Tag = "dcombo") Then
        ctl.Text = "Column Name"
    End If
    Next ctl
    '---------------------------
    '��----------------------------------------
    
    '��Joint�󋵊m�F�E�Ή�---------------------
    If tableXnm <> "" Then
        Label3.Caption = tableXnm
    Else
        Label3.Caption = "Saving Nothing"
    End If
    
    If Combo6.Visible = True Then
        For Each ctl In Me.Controls
            If (ctl.Tag = "syukei") Then
                ctl.Visible = False
            End If
        Next ctl
    End If
    '��----------------------------------------
    

    '��DataCombo1�Ƀe�[�u���ꗗ��\������-----------------------
    cn.Open cnstr
     
    ssql = "SELECT name FROM sysobjects WHERE xtype = 'u'"
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs_c1 = rs.Clone
    
    Call RSclose
        
    Set DataCombo1.RowSource = rs_c1
    DataCombo1.ListField = "Name"
    DataCombo1.Refresh
    Set DataCombo12.RowSource = rs_c1
    DataCombo12.ListField = "Name"
    DataCombo12.Refresh
    '��---------------------------------------------------------
'��--------------------------------------------------
    
End Sub





