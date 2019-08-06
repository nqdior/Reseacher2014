Attribute VB_Name = "SPS0000M"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim PASS_ENTR As String
Public SistemINI As String '2009/03 nohara add クライアントサーバーの設定

'ヘルプボタン用
    Public helpPageNo As Integer
    Public mPageNo(11) As Long 'マニュアルページ
    Public strHelpURL As String
    Public strHelpPDF As String
'optKANRI色管理用
    Public optindex As Integer '前回クリックしたIndex
    Public optIndex_move As Integer '前回moveしたIndex
    

Public Sub Main()
    Dim MSDE As String '2009/03 nohara add クライアントサーバーの設定
                       '0：MSDEがｲﾝｽﾄｰﾙされている，1：MSDEがｲﾝｽﾄｰﾙされていない
    
    If App.PrevInstance Then '二重起動チェック
        'MsgBox ("既に起動されています")
        End
    End If
    
    
    Screen.MousePointer = vbHourglass
    Openingf.Show
    DoEvents
    
    '2009/03 nohara add クライアントサーバーの設定 -start///////////////////////////////////////////////////////////////////////////////////
    'INIファイルよりクライアントorサーバー
    SistemINI = IIf(getIniFileInfo("SPOS.INI", "SPOS", "CLIENT_SERVER") = "", "SERVER", getIniFileInfo("SPOS.INI", "SPOS", "CLIENT_SERVER"))
    
    '↓2003/2/10
    'INIファイルより　0：MSDEがｲﾝｽﾄｰﾙされている，1：MSDEがｲﾝｽﾄｰﾙされていない
    MSDE = IIf(getIniFileInfo("SPOS.INI", "SPOS", "MSDE") = "", "0", getIniFileInfo("SPOS.INI", "SPOS", "MSDE"))
    '2009/03 nohara add クライアントサーバーの設定 -end////////////////////////////////////////////////////////////////////////////////////
    
    If MSDE = "0" Then '0：MSDEがｲﾝｽﾄｰﾙされている'2009/03 nohara add クライアントサーバーの設定
        '----ローカルMSDE起動------------------------------
        If OpenLocalMDB() = False Then
            'ローカルＭＤＢへ接続できないとき
            End
        End If
'SQLServer2008Exp対応のためコメント-----------------------------------------------------------------
'        If Not StartSQL7("(local)" _
'                , GetLocalParameter(DB_CON_INFO, "LOGIN_ID") _
'                , GetLocalParameter(DB_CON_INFO, "LOGIN_PASSWORD") _
'                , getParamaterValFromConnectString( _
'                    GetLocalParameter(DB_CON_INFO, "CONNECTION_STRING"), "Initial Catalog") _
'                ) Then
'            MsgBox "データベースが起動できません。", vbOKOnly, "メニュー管理システム"
'            End
'        End If
'SQLServer2008Exp対応のためコメント-----------------------------------------------------------------
        Call DDB_Restore("(local)" _
                , GetLocalParameter(DB_CON_INFO, "LOGIN_ID") _
                , GetLocalParameter(DB_CON_INFO, "LOGIN_PASSWORD") _
                , getParamaterValFromConnectString( _
                    GetLocalParameter(DB_CON_INFO, "CONNECTION_STRING"), "Initial Catalog"))
        Call CloseLocalMDB
        '---------------------------------------------------
    End If
    
    If OpenLocalMDB() = False Then
        'ローカルＭＤＢへ接続できないとき
        End
    End If
'    Sleep 5000
    Screen.MousePointer = vbDefault
    
    If Not ConnectSPSDATA() Then
        MsgBox "データベースへ接続できません。", vbOKOnly, "食堂POSシステム"
        End
    End If
    Unload Openingf
    
    PASS_ENTR = GetEnvironVal("0010", "0001") 'パスワード設定
    If PASS_ENTR = "1" Then
        PASSENTR.Show
    Else
        SPS0000F.Show
    End If
    
End Sub
'SQLServer2008EXP 対応時追加 Add Start-----------------
'DDBからSPSDATA作成
Function DDB_Restore(svNM As String, loginID As String, loginPW As String, dbName As String) As Boolean

    Dim ErrorString As String
    Dim i As Integer, j As Integer, k As Integer
    Dim clsINI As New clsIniFile
    Dim arSectionNM() As String
    Dim arSectionString() As String
    Dim ar() As String
    Dim DatabaseNothingFLG As Boolean 'True:ない False:ある
    Dim strSqlDBPath As String 'SQLServerのDB保存場所 C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA
    
'SetupDateがIncompleteであれば、DB作成を行う　Win7版
    If "Incomplete" = getIniFileInfo(SPOS_INI, "SETUP", "SetupDate") Then
        'セットアップ未完了時
        clsINI.FileName = App.Path & "\" & SPOS_INI 'INIファイル設定
        clsINI.EmnumSection arSectionNM
        For j = 0 To UBound(arSectionNM)
            clsINI.Section = arSectionNM(j)
            clsINI.EnumSectionString arSectionString
            For k = 0 To UBound(arSectionString)
                ar = Split(arSectionString(k), "=")
                If UBound(ar) > 0 Then
                    If Repl_Str(ar(1), ".\" _
                                , App.Path & IIf(Right(App.Path, 1) = "\", "", "\")) > 0 Then

                        'INIファイル内容書き換え（『.\』 ⇒『App.Path』）
                        Call clsINI.SetString(ar(0), ar(1))

                    End If
                End If
            Next k
        Next j

        'SQLServerのDB保存場所をINIファイルより取得
        strSqlDBPath = getIniFileInfo(SPOS_INI, "SETUP", "SQLServerDBPath")
        strSqlDBPath = IIf(strSqlDBPath = "", "C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA", strSqlDBPath)
        
        'データベースが無い時はデフォルトデータをリストアする。
'        Call executeShell("SPSRESTP.exe " & App.Path & "\DDB", True)
'DbRestor.exe SPSDATA,D:\Spospro2011\DDB,C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA,sa,5473nsk0036
                                                        
            Call executeShell("DbRestor.exe " & dbName & "," & App.Path & IIf(Right(App.Path, 1) = "\", "DDB", "\DDB") & "," _
                                                        & strSqlDBPath & "," & loginID & "," & loginPW, True)

        Call writeIniFileInfo(SPOS_INI, "SETUP", "SetupDate", Format(Date, "YYYY/MM/DD"))
    End If

End Function
'SQLServer2008EXP 対応時追加 Add End-------------------

''SQLServer2008EXP 対応のためコメント---------------------------------------------------------------------------
 Function StartSQL7(svNM As String, loginID As String, loginPW As String, dbName As String) As Boolean

''This function starts SQL Server or MSDE, and tries to connect to it
'
'    Dim ErrorString As String
'    Dim oSvroot As Object
'    Dim i As Integer, j As Integer, k As Integer
'    Dim clsINI As New clsIniFile
'    Dim arSectionNM() As String
'    Dim arSectionString() As String
'    Dim ar() As String
'    Dim DatabaseNothingFLG As Boolean 'True:ない False:ある
'
'
'    Set oSvroot = CreateObject("SQLDMO.SQLServer")
'
'    On Error GoTo StartError
'
'    'Set the time out fairly high.
'    'Note this value is in seconds.
'    oSvroot.LoginTimeout = 60
'
'    ' Start the service logon would look like this.
'    oSvroot.Start True, svNM, loginID, loginPW
'
'    ' This just starts the service
'    ' oSvr.Start False, "(local)"
'
'ExitFunc:
'    StartSQL7 = oSvroot.VerifyConnection
'
'    If "Incomplete" = getIniFileInfo(SPOS_INI, "SETUP", "SetupDate") Then
'        'セットアップ未完了時
'        clsINI.FileName = App.Path & "\" & SPOS_INI 'INIファイル設定
'        clsINI.EmnumSection arSectionNM
'        For j = 0 To UBound(arSectionNM)
'            clsINI.Section = arSectionNM(j)
'            clsINI.EnumSectionString arSectionString
'            For k = 0 To UBound(arSectionString)
'                ar = Split(arSectionString(k), "=")
'                If UBound(ar) > 0 Then
'                    If Repl_Str(ar(1), ".\" _
'                                , App.Path & IIf(Right(App.Path, 1) = "\", "", "\")) > 0 Then
'
'                        'INIファイル内容書き換え（『.\』 ⇒『App.Path』）
'                        Call clsINI.SetString(ar(0), ar(1))
'
'                    End If
'                End If
'            Next k
'        Next j
'
'        DatabaseNothingFLG = True
'        For i = 1 To oSvroot.Databases.Count
'            If oSvroot.Databases(i).Name = dbName Then
'                DatabaseNothingFLG = False
'                Exit For
'            End If
'        Next i
'        If DatabaseNothingFLG = True Then
'            'データベースが無い時はデフォルトデータをリストアする。
'                Call executeShell("DbRestor.exe " & dbName & "," & App.Path & IIf(Right(App.Path, 1) = "\", "DDB", "\DDB") & "," _
'                                                        & oSvroot.Databases(1).PrimaryFilePath, True)
'        End If
'
'        Call writeIniFileInfo(SPOS_INI, "SETUP", "SetupDate", Format(Date, "YYYY/MM/DD"))
'    End If
'
'
'    If StartSQL7 = False Then
'       ErrorString = "Could not start or connect to SQL Server"
'        MsgBox ErrorString
'    End If
'
'
'    Set clsINI = Nothing
'
'    Exit Function
'
'StartError:
'    ' This error happens if service is already running
'    If Err.Number = -2147023840 Or Err.Number = 1056 Or Err.Number = 440 Then
'       ' Use this to logon
'        oSvroot.Connect svNM, loginID, loginPW
'       ' Otherwise
'       Resume Next
'    Else
'    StartSQL7 = True
''        ErrorString = "SQL Server start failed.  " + Chr(13) + "SQL-DMO error:  " _
''            + Str(Err.Number) + " " + Err.Description
''       MsgBox ErrorString
''       StartSQL7 = False
'    End If
'
''SQLServer2008EXP 対応のためコメント---------------------------------------------------------------------------
End Function

'ＡＤＯ接続文字列から指定されたパラメータの値を返します。
Function getParamaterValFromConnectString(pConnectString As String, paraName As String) As String
    Dim ar() As String
    Dim i As Integer

    ar = Split(pConnectString, ";")

    getParamaterValFromConnectString = ""
    For i = 0 To UBound(ar)
        If InStr(1, ar(i), paraName & "=") > 0 Then
            getParamaterValFromConnectString = Mid(ar(i), Len(paraName & "=") + 1)
            Exit For
        End If
    Next i

End Function

'文字列置換
Function Repl_Str(sSrcStr As String, sFndStr As String, sRepStr As String) As Long

    Dim sTmpBefore As String    '置換対象文字列以前部分
    Dim sTmpAfter As String     '置換対象文字列以降部分
    Dim lDelStrPos As Long      '置換対象文字列ポジション
    Dim lSrcStart As Long       '検索開始ポジション
    Dim lReplCount As Long      '置換個数カウンタ

    '置換文字の個数をイニシャルクリア！
    lReplCount = 0

    '検索開始位置を設定。
    lSrcStart = 1

    '置換対象文字列の位置を検索！
    lDelStrPos = InStr(lSrcStart, sSrcStr, sFndStr)

    Do Until lDelStrPos = 0
        '置換対象文字列が見つかった位置から前の部分を抽出。
        sTmpBefore = Left(sSrcStr, lDelStrPos - 1)

        '置換対象文字列が見つかった位置から後ろの部分を抽出。
        sTmpAfter = Right(sSrcStr, (Len(sSrcStr) - (lDelStrPos + Len(sFndStr) - 1)))

        '抽出した文字列の間に置換後の文字列を挟むと、置換したことになる。
        sSrcStr = sTmpBefore & sRepStr & sTmpAfter

        '検索開始位置を設定し直す(置換後の文字列を検索しないため)。
        lSrcStart = Len(sTmpBefore) + Len(sRepStr) + 1

        '置換文字の個数をカウントアップ！
        lReplCount = lReplCount + 1

        '置換対象文字列の位置を検索！
        lDelStrPos = InStr(lSrcStart, sSrcStr, sFndStr)

    Loop

    '置換文字の個数を戻す。
    Repl_Str = lReplCount

End Function
