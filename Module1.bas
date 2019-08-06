Attribute VB_Name = "SRModule"
 Option Explicit
 Public cn As New ADODB.Connection
 Public cna As New ADODB.Connection 'access
 Public rs As New ADODB.Recordset
 Public rsa As New ADODB.Recordset 'access
 Public rs_c0 As New ADODB.Recordset 'form0にて使用
 Public rs_c1 As New ADODB.Recordset 'datacombo1
 Public rs_c2 As New ADODB.Recordset 'datacombo2
 Public rs_c3 As New ADODB.Recordset 'form2起動時に使用
 Public rs_c4 As New ADODB.Recordset 'form2dataeditに使用
 Public rs_c5 As New ADODB.Recordset 'ソート用
 Public rs_c6 As New ADODB.Recordset
 Public rs_c7 As New ADODB.Recordset 'join用
 Public ctl As Control 'オブジェクト一括非表示
 Public oControl As Control
 Public ssql As String '重要
 Public ssql1 As String '帳票使用
 Public pSql As String '重用
 Public cSql1 As String 'コンボボックス用
 Public Nowform As String 'form2close時使用
 Public i As Integer '各ループ時Integer格納
 Public selDB As String 'select DB格納
 Public cnstr As String 'con string格納
 Public cnstra As String 'con string access
 Public instance As String 'PC instance格納
 Public Lid As String 'LoginID格納
 Public Lpass As String 'LoginPass格納
 Public x As Integer
 
 '▼joint対応追加分-------
 Public tmp As String
 Public tableX As String 'joint table'X'格納
 Public tableXnm As String 'joint table名格納
 Public tableXstr As String 'jointtable+@
 Public settable As String
 Public jointflg As String
 Public jointbl As String
 '▲20141121---------------
 
 '▼Joinjoint対応追加分----
 Public Wr As Integer
 Public selcol1 As String
 Public selcol2 As String
 Public selcol As String
 Public retcode As String
 '▲-----------------------
 
 Public Asql As String
 Public j As Integer
 Public k As Integer
 Public strcriteria As String


'▼ここから処理開始------------------------------------
Public Sub Main()
    StartingForm.Show
End Sub
'▲----------------------------------------------------


'▼cn状態確認及び閉じる処理----------------------------
Public Function CNclose()
    If cn.State <> 0 Then
        cn.Close
    End If
    ssql = ""
End Function
'▲----------------------------------------------------


'▼rs状態確認及び閉じる処理----------------------------
Public Function RSclose()
    If rs.State <> 0 Then
        rs.Close
    End If
End Function
'------------------------------------------------------


'■テーブルXを読み込む------------------------------------
Public Function TableCall()
           
    '▼TableX内Orderby切捨---------------------------
    tmp = InStr(1, tableX, "order")
    If tmp <> 0 Then
        tableX = Left(tableX, tmp - 2)
    End If
    '▲----------------------------------------------
    
    jointflg = "1"
'■-------------------------------------------------------

End Function


'null値をEmpty値に変える
Public Function nullEmpty(inVal As Variant) As Variant
    If IsNull(inVal) Then
        nullEmpty = Empty
    Else
        nullEmpty = inVal
    End If
End Function


'Form起動前にアクション
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

