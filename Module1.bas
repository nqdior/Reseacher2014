Attribute VB_Name = "SRModule"
 Option Explicit
 Public cn As New ADODB.Connection
 Public cna As New ADODB.Connection 'access
 Public rs As New ADODB.Recordset
 Public rsa As New ADODB.Recordset 'access
 Public rs_c0 As New ADODB.Recordset 'form0�ɂĎg�p
 Public rs_c1 As New ADODB.Recordset 'datacombo1
 Public rs_c2 As New ADODB.Recordset 'datacombo2
 Public rs_c3 As New ADODB.Recordset 'form2�N�����Ɏg�p
 Public rs_c4 As New ADODB.Recordset 'form2dataedit�Ɏg�p
 Public rs_c5 As New ADODB.Recordset '�\�[�g�p
 Public rs_c6 As New ADODB.Recordset
 Public rs_c7 As New ADODB.Recordset 'join�p
 Public ctl As Control '�I�u�W�F�N�g�ꊇ��\��
 Public oControl As Control
 Public ssql As String '�d�v
 Public ssql1 As String '���[�g�p
 Public pSql As String '�d�p
 Public cSql1 As String '�R���{�{�b�N�X�p
 Public Nowform As String 'form2close���g�p
 Public i As Integer '�e���[�v��Integer�i�[
 Public selDB As String 'select DB�i�[
 Public cnstr As String 'con string�i�[
 Public cnstra As String 'con string access
 Public instance As String 'PC instance�i�[
 Public Lid As String 'LoginID�i�[
 Public Lpass As String 'LoginPass�i�[
 Public x As Integer
 
 '��joint�Ή��ǉ���-------
 Public tmp As String
 Public tableX As String 'joint table'X'�i�[
 Public tableXnm As String 'joint table���i�[
 Public tableXstr As String 'jointtable+@
 Public settable As String
 Public jointflg As String
 Public jointbl As String
 '��20141121---------------
 
 '��Joinjoint�Ή��ǉ���----
 Public Wr As Integer
 Public selcol1 As String
 Public selcol2 As String
 Public selcol As String
 Public retcode As String
 '��-----------------------
 
 Public Asql As String
 Public j As Integer
 Public k As Integer
 Public strcriteria As String


'���������珈���J�n------------------------------------
Public Sub Main()
    StartingForm.Show
End Sub
'��----------------------------------------------------


'��cn��Ԋm�F�y�ѕ��鏈��----------------------------
Public Function CNclose()
    If cn.State <> 0 Then
        cn.Close
    End If
    ssql = ""
End Function
'��----------------------------------------------------


'��rs��Ԋm�F�y�ѕ��鏈��----------------------------
Public Function RSclose()
    If rs.State <> 0 Then
        rs.Close
    End If
End Function
'------------------------------------------------------


'���e�[�u��X��ǂݍ���------------------------------------
Public Function TableCall()
           
    '��TableX��Orderby�؎�---------------------------
    tmp = InStr(1, tableX, "order")
    If tmp <> 0 Then
        tableX = Left(tableX, tmp - 2)
    End If
    '��----------------------------------------------
    
    jointflg = "1"
'��-------------------------------------------------------

End Function


'null�l��Empty�l�ɕς���
Public Function nullEmpty(inVal As Variant) As Variant
    If IsNull(inVal) Then
        nullEmpty = Empty
    Else
        nullEmpty = inVal
    End If
End Function


'Form�N���O�ɃA�N�V����
Public Sub Form_Initialize()
    MsgBox ("The form is loading")
End Sub


Public Sub CConvert()

    cnstra = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SQLRes.mdb;Persist Security Info=False"
    cna.Open cnstra
    Asql = "SELECT * FROM ColumnName"
    rsa.Open Asql, cna, adOpenStatic, adLockOptimistic, adCmdText
    
    Dim strcriteria As String


    For i = 0 To TDBGrid1.Columns.Count - 1
        strcriteria = "ColumnName = " & "'" & TDBGrid1.Columns(i).Name & "'"
        rsa.Find strcriteria, 0, adSearchForward
        For x = 0 To rsa.EOF
            If TDBGrid1.Columns(i).Name = rsa!ColumnName Then
                TDBGrid1.Columns(i).Caption = rsa!columnJPN
            End If
            rsa.MoveNext
        Next x
    Next i

End Sub

