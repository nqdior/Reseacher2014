Attribute VB_Name = "SPS0000M"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim PASS_ENTR As String
Public SistemINI As String '2009/03 nohara add �N���C�A���g�T�[�o�[�̐ݒ�

'�w���v�{�^���p
    Public helpPageNo As Integer
    Public mPageNo(11) As Long '�}�j���A���y�[�W
    Public strHelpURL As String
    Public strHelpPDF As String
'optKANRI�F�Ǘ��p
    Public optindex As Integer '�O��N���b�N����Index
    Public optIndex_move As Integer '�O��move����Index
    

Public Sub Main()
    Dim MSDE As String '2009/03 nohara add �N���C�A���g�T�[�o�[�̐ݒ�
                       '0�FMSDE���ݽİق���Ă���C1�FMSDE���ݽİق���Ă��Ȃ�
    
    If App.PrevInstance Then '��d�N���`�F�b�N
        'MsgBox ("���ɋN������Ă��܂�")
        End
    End If
    
    
    Screen.MousePointer = vbHourglass
    Openingf.Show
    DoEvents
    
    '2009/03 nohara add �N���C�A���g�T�[�o�[�̐ݒ� -start///////////////////////////////////////////////////////////////////////////////////
    'INI�t�@�C�����N���C�A���gor�T�[�o�[
    SistemINI = IIf(getIniFileInfo("SPOS.INI", "SPOS", "CLIENT_SERVER") = "", "SERVER", getIniFileInfo("SPOS.INI", "SPOS", "CLIENT_SERVER"))
    
    '��2003/2/10
    'INI�t�@�C�����@0�FMSDE���ݽİق���Ă���C1�FMSDE���ݽİق���Ă��Ȃ�
    MSDE = IIf(getIniFileInfo("SPOS.INI", "SPOS", "MSDE") = "", "0", getIniFileInfo("SPOS.INI", "SPOS", "MSDE"))
    '2009/03 nohara add �N���C�A���g�T�[�o�[�̐ݒ� -end////////////////////////////////////////////////////////////////////////////////////
    
    If MSDE = "0" Then '0�FMSDE���ݽİق���Ă���'2009/03 nohara add �N���C�A���g�T�[�o�[�̐ݒ�
        '----���[�J��MSDE�N��------------------------------
        If OpenLocalMDB() = False Then
            '���[�J���l�c�a�֐ڑ��ł��Ȃ��Ƃ�
            End
        End If
'SQLServer2008Exp�Ή��̂��߃R�����g-----------------------------------------------------------------
'        If Not StartSQL7("(local)" _
'                , GetLocalParameter(DB_CON_INFO, "LOGIN_ID") _
'                , GetLocalParameter(DB_CON_INFO, "LOGIN_PASSWORD") _
'                , getParamaterValFromConnectString( _
'                    GetLocalParameter(DB_CON_INFO, "CONNECTION_STRING"), "Initial Catalog") _
'                ) Then
'            MsgBox "�f�[�^�x�[�X���N���ł��܂���B", vbOKOnly, "���j���[�Ǘ��V�X�e��"
'            End
'        End If
'SQLServer2008Exp�Ή��̂��߃R�����g-----------------------------------------------------------------
        Call DDB_Restore("(local)" _
                , GetLocalParameter(DB_CON_INFO, "LOGIN_ID") _
                , GetLocalParameter(DB_CON_INFO, "LOGIN_PASSWORD") _
                , getParamaterValFromConnectString( _
                    GetLocalParameter(DB_CON_INFO, "CONNECTION_STRING"), "Initial Catalog"))
        Call CloseLocalMDB
        '---------------------------------------------------
    End If
    
    If OpenLocalMDB() = False Then
        '���[�J���l�c�a�֐ڑ��ł��Ȃ��Ƃ�
        End
    End If
'    Sleep 5000
    Screen.MousePointer = vbDefault
    
    If Not ConnectSPSDATA() Then
        MsgBox "�f�[�^�x�[�X�֐ڑ��ł��܂���B", vbOKOnly, "�H��POS�V�X�e��"
        End
    End If
    Unload Openingf
    
    PASS_ENTR = GetEnvironVal("0010", "0001") '�p�X���[�h�ݒ�
    If PASS_ENTR = "1" Then
        PASSENTR.Show
    Else
        SPS0000F.Show
    End If
    
End Sub
'SQLServer2008EXP �Ή����ǉ� Add Start-----------------
'DDB����SPSDATA�쐬
Function DDB_Restore(svNM As String, loginID As String, loginPW As String, dbName As String) As Boolean

    Dim ErrorString As String
    Dim i As Integer, j As Integer, k As Integer
    Dim clsINI As New clsIniFile
    Dim arSectionNM() As String
    Dim arSectionString() As String
    Dim ar() As String
    Dim DatabaseNothingFLG As Boolean 'True:�Ȃ� False:����
    Dim strSqlDBPath As String 'SQLServer��DB�ۑ��ꏊ C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA
    
'SetupDate��Incomplete�ł���΁ADB�쐬���s���@Win7��
    If "Incomplete" = getIniFileInfo(SPOS_INI, "SETUP", "SetupDate") Then
        '�Z�b�g�A�b�v��������
        clsINI.FileName = App.Path & "\" & SPOS_INI 'INI�t�@�C���ݒ�
        clsINI.EmnumSection arSectionNM
        For j = 0 To UBound(arSectionNM)
            clsINI.Section = arSectionNM(j)
            clsINI.EnumSectionString arSectionString
            For k = 0 To UBound(arSectionString)
                ar = Split(arSectionString(k), "=")
                If UBound(ar) > 0 Then
                    If Repl_Str(ar(1), ".\" _
                                , App.Path & IIf(Right(App.Path, 1) = "\", "", "\")) > 0 Then

                        'INI�t�@�C�����e���������i�w.\�x �ˁwApp.Path�x�j
                        Call clsINI.SetString(ar(0), ar(1))

                    End If
                End If
            Next k
        Next j

        'SQLServer��DB�ۑ��ꏊ��INI�t�@�C�����擾
        strSqlDBPath = getIniFileInfo(SPOS_INI, "SETUP", "SQLServerDBPath")
        strSqlDBPath = IIf(strSqlDBPath = "", "C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA", strSqlDBPath)
        
        '�f�[�^�x�[�X���������̓f�t�H���g�f�[�^�����X�g�A����B
'        Call executeShell("SPSRESTP.exe " & App.Path & "\DDB", True)
'DbRestor.exe SPSDATA,D:\Spospro2011\DDB,C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA,sa,5473nsk0036
                                                        
            Call executeShell("DbRestor.exe " & dbName & "," & App.Path & IIf(Right(App.Path, 1) = "\", "DDB", "\DDB") & "," _
                                                        & strSqlDBPath & "," & loginID & "," & loginPW, True)

        Call writeIniFileInfo(SPOS_INI, "SETUP", "SetupDate", Format(Date, "YYYY/MM/DD"))
    End If

End Function
'SQLServer2008EXP �Ή����ǉ� Add End-------------------

''SQLServer2008EXP �Ή��̂��߃R�����g---------------------------------------------------------------------------
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
'    Dim DatabaseNothingFLG As Boolean 'True:�Ȃ� False:����
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
'        '�Z�b�g�A�b�v��������
'        clsINI.FileName = App.Path & "\" & SPOS_INI 'INI�t�@�C���ݒ�
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
'                        'INI�t�@�C�����e���������i�w.\�x �ˁwApp.Path�x�j
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
'            '�f�[�^�x�[�X���������̓f�t�H���g�f�[�^�����X�g�A����B
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
''SQLServer2008EXP �Ή��̂��߃R�����g---------------------------------------------------------------------------
End Function

'�`�c�n�ڑ������񂩂�w�肳�ꂽ�p�����[�^�̒l��Ԃ��܂��B
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

'������u��
Function Repl_Str(sSrcStr As String, sFndStr As String, sRepStr As String) As Long

    Dim sTmpBefore As String    '�u���Ώە�����ȑO����
    Dim sTmpAfter As String     '�u���Ώە�����ȍ~����
    Dim lDelStrPos As Long      '�u���Ώە�����|�W�V����
    Dim lSrcStart As Long       '�����J�n�|�W�V����
    Dim lReplCount As Long      '�u�����J�E���^

    '�u�������̌����C�j�V�����N���A�I
    lReplCount = 0

    '�����J�n�ʒu��ݒ�B
    lSrcStart = 1

    '�u���Ώە�����̈ʒu�������I
    lDelStrPos = InStr(lSrcStart, sSrcStr, sFndStr)

    Do Until lDelStrPos = 0
        '�u���Ώە����񂪌��������ʒu����O�̕����𒊏o�B
        sTmpBefore = Left(sSrcStr, lDelStrPos - 1)

        '�u���Ώە����񂪌��������ʒu������̕����𒊏o�B
        sTmpAfter = Right(sSrcStr, (Len(sSrcStr) - (lDelStrPos + Len(sFndStr) - 1)))

        '���o����������̊Ԃɒu����̕���������ނƁA�u���������ƂɂȂ�B
        sSrcStr = sTmpBefore & sRepStr & sTmpAfter

        '�����J�n�ʒu��ݒ肵����(�u����̕�������������Ȃ�����)�B
        lSrcStart = Len(sTmpBefore) + Len(sRepStr) + 1

        '�u�������̌����J�E���g�A�b�v�I
        lReplCount = lReplCount + 1

        '�u���Ώە�����̈ʒu�������I
        lDelStrPos = InStr(lSrcStart, sSrcStr, sFndStr)

    Loop

    '�u�������̌���߂��B
    Repl_Str = lReplCount

End Function
