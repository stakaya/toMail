Attribute VB_Name = "SubMain"
'
'LhaZip.bas  Ver 1.0
'
'�@�\�FLha�`��
'�K�v��DLL:UnLHA32.dll
' https://micco.mars.jp/mysoft/unlha32.htm
'
'====================================================================
'1.Lha�`���ň��k����B
'
'�֐��FLzhPack(lngWnd , strPreFile , strDir , [strSwitch] , [strOption])
'
'����   lngWnd      :�ΏۂƂȂ�E�C���h�E�n���h��(��FForm1.hWnd)
'       strPreFile  :���k�������t�@�C���̃p�X(��: c:\My Documents\test.txt)
'       strDir      :���k�t�@�C���̏o�͐�(��: C:\My Documents\)
'       strSwitch   :�R�}���h���C���ɓn���X�C�b�`(�ȗ���:�f�t�H���g�ł�"u"  ��L�Q��)
'       strOption   :�R�}���h���C���ɓn���I�v�V����(�ȗ���:�f�t�H���g�ł�"-a1 -r2 -x1 -l1 -jp -o2"�@��L�Q��)
'
'
'�߂�l�F����G0�@�@���s�G1
'====================================================================
'
'
'
'====================================================================


Option Explicit

'API��錾
Private Declare Function Unlha Lib "UnLHA32.dll" _
(ByVal hWnd As Long, ByVal szCmdLine As String, ByVal szOutput As String, ByVal dwSize As Long) As Long
                                                
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
(ByVal a As Long, ByVal b As String, c As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal a As Long, ByVal b As String, ByVal _
c As Long, d As Long, e As Any, f As Long) As Long

Private Declare Function RegCloseKey _
Lib "advapi32.dll" (ByVal a As Long) As Long
                                                
Public listNum As Long

Sub Main()
    Dim t() As Variant
    Dim i As Long
    Dim path As String
    

    t = GetCommandLine(500)
    
    listNum = UBound(t)
    For i = 1 To listNum
        zip.List.List(i - 1) = t(i)
    Next i
        
    path = GetSetting("lzh", "out", "path", "c:\temp")
    
    If 0 <> listNum Then
        For i = 0 To listNum - 1
            Call LzhPack(zip.hWnd, zip.List.List(i), path)
        Next i
    Else
        Call MsgBox("�Ώۃt�@�C�����h���b�O���Ă��������B", vbOKOnly + vbExclamation, "���[��")
    End If
    End
End Sub
Function GetCommandLine(Optional MaxArgs)
   ' �ϐ���錾���܂��B
   Dim c, CmdLine, CmdLnLen, InArg, i, NumArgs
   ' MaxArgs ���񋟂���邩�ǂ����𒲂ׂ܂��B
   If IsMissing(MaxArgs) Then MaxArgs = 10
   ' ���݂̃T�C�Y�̔z��ɂ��܂��B
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   ' �R�}���h ���C���̈������擾���܂��B
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   ' �����ɃR�}���h ���C���̈������擾���܂��B
   For i = 1 To CmdLnLen
      c = Mid(CmdLine, i, 1)
      ' �X�y�[�X�܂��̓^�u�𒲂ׂ܂��B
      If (c <> " " And c <> vbTab) Then
         ' �X�y�[�X�܂��̓^�u�̂�����ł�����܂���B
         ' ���Ɉ����̒��ł͂Ȃ����ǂ����𒲂ׂ܂��B
         If Not InArg Then
         ' �V�����������n�܂�܂��B
         ' �������������Ȃ����𒲂ׂ܂��B
            If NumArgs = MaxArgs Then Exit For
               NumArgs = NumArgs + 1
               InArg = True
            End If
         ' ���݂̈����ɕ�����ǉ����܂��B
         ArgArray(NumArgs) = ArgArray(NumArgs) & c
      Else
         ' �X�y�[�X�܂��̓^�u�������܂����B
         ' InArg �t���O�� False ��ݒ肵�܂��B
         InArg = False
      End If
   Next i
   ' ���������ׂĊi�[�ł���悤�ɔz��̃T�C�Y��ύX���܂��B
   ReDim Preserve ArgArray(NumArgs)
   ' �֐����ɔz���Ԃ��܂��B
   GetCommandLine = ArgArray()
End Function
                                        

'
'Lzh�`���ň��k���s��
'
Public Function LzhPack(ByVal lngWnd As Long, _
                        ByVal strPreFile As String, _
                        ByVal strDir As String, _
                        Optional ByVal strSwitch As String = "u", _
                        Optional ByVal strOptions As String = "-a1 -r2 -x1 -l1 -jp -o2") As Long


    'lngWnd     :�ΏۂƂȂ�E�C���h�E�n���h��(��FForm1.hWnd)
    'strPreFile :���k�������t�@�C���̃p�X(��: c:\My Documents\test.txt)
    'strDir     :���k�t�@�C���̏o�͐�(��: C:\My Documents\)
    'strSwitch  :�R�}���h���C���Ƃ��ēn���X�C�b�`(�ȗ���:�f�t�H���g�ł�"u"  ��L�Q��)
    'strOption  :�R�}���h���C���ɓn���I�v�V����(�ȗ���:�f�t�H���g�ł�"-a1 -r2 -x1 -l1 -jp -o2"�@��L�Q��)
    
    
    Dim strLzhFile As String    '���k��̃t�@�C���̃p�X
    
    Dim strCommandLine As String
    Dim strBuffer As String * 1024
    
    Dim lngResult As Long   '�߂�l���i�[
    
    Dim pvHnd  As Long                      '�������
    Dim pvType As Long                      '����
    Dim pvStr  As String * 1024             '�ޯ̧
    Dim pvRet  As Long                      '�߂�l
    Dim mail   As String
    Dim shel   As String
    Dim flag   As Integer
    Dim i      As Integer
    
    '�h���C�u�����k���悤�Ƃ����Ƃ��̓G���[��Ԃ��B
    If Len(strPreFile) < 4 Then
        LzhPack = 1
        Call MsgBox("�h���C�u�͈��k�ł��܂���BNo.1", vbOKOnly, "LzhPack�G���[")
        Exit Function
    End If
    
    '���k���悤�Ƃ��Ă���t�@�C�������݂��Ȃ��Ƃ��̓G���[��Ԃ��B
    If Len(Dir$(strPreFile, 31)) = 0 Then
        LzhPack = 1
        Call MsgBox("�t�@�C����������܂���BNo.2", vbOKOnly, "LzhPack�G���[")
        Exit Function
    End If
    
    '���k��̃t�@�C���̃p�X���擾
    strLzhFile = AddDirSep(strDir) & PackName(strPreFile, ".lzh")
    
    '�R�}���h���C��
   If Dir(strPreFile, vbNormal) = Empty Then
       strCommandLine = strSwitch & " " & _
                     strOptions & " " & _
                     AddQuotesToFN(strLzhFile) & " " & strPreFile
   Else
       strCommandLine = strSwitch & " " & _
                     strOptions & " " & _
                     AddQuotesToFN(strLzhFile) & " " & _
                     AddQuotesToFN(AddDirSep(GetPath(strPreFile))) & " " & _
                     AddQuotesToFN(Dir$(strPreFile, 31)) & " "
   End If

            
    '���k�̎��s:���탊�^�[��=0
    lngResult = Unlha(lngWnd, strCommandLine, strBuffer, Len(strBuffer))
    'Call MsgBox(strCommandLine)
    
    '�l��Ԃ�
    If lngResult <> 0 Then
        Call MsgBox("���k�Ɏ��s���܂����BNo.3", vbOKOnly + vbExclamation, "���k���s")
        LzhPack = 1
        Exit Function
    Else
        LzhPack = 0
    End If
    
    If RegOpenKey(&H80000000, "mailto\shell\open\command", pvHnd) = 0 Then
        If RegQueryValueEx(pvHnd, "", 0, pvType, ByVal pvStr, 1024) = 0 Then
                                            '������̒��o
            mail = Left(pvStr, InStr(pvStr, Chr(0)) - 1)
        End If
    End If
    pvRet = RegCloseKey(pvHnd)              'ڼ޽�ظ۰��

    For i = 1 To Len(mail)
        If Mid(mail, i, 1) <> Chr(34) Then
            flag = 1
        End If
        If Mid(mail, i, 1) <> Chr(34) And flag <> 0 Then
            shel = shel & Mid(mail, i, 1)
        End If
        If Mid(mail, i, 1) = Chr(34) And flag = 1 Then
            Exit For
        End If
    Next i

     Call Shell(shel & " " & AddQuotesToFN(strLzhFile), vbNormalFocus)



End Function


'
'�p�X���ɃX�y�[�X��J���}���܂܂�Ă���ꍇ�ɂ��̃p�X���ɓ�d���p����t�����ĕԂ��B
'
Private Function AddQuotesToFN(ByVal strFileName As String) As String

    If InStr(strFileName, " ") Or InStr(strFileName, ".") Then
        AddQuotesToFN = """" & strFileName & """"
    Else
        AddQuotesToFN = strFileName
    End If
    
End Function

'
'�p�X�̖����ɋ�؂�L���i���j���Ȃ��Ƃ��ɂ́i���j������
'
Private Function AddDirSep(strPathName As String) As String

    If Right$(strPathName, 1) = "\" Then
        AddDirSep = RTrim$(strPathName)
    Else
        AddDirSep = RTrim$(strPathName) & "\"
    End If
    
End Function


'���k�O�̃t�@�C���̏ꏊ���擾
Private Function GetPath(ByVal strPathName As String) As String

    '�p�X���̍Ōオ"\"�Ȃ�͂���
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If
    
    'strPathName���h���C�u�Ȃ炻�̂܂ܕԂ�
    If Right$(strPathName, 1) = ":" Then
        
        GetPath = strPathName
        
        Exit Function
        
    End If
    
    
    '�߂�l���i�[

    GetPath = Left$(strPathName, InStrRev(strPathName, "\") - 1)

End Function


'���k��̃t�@�C�������쐬
Private Function PackName(ByVal strPathName As String, ByVal strExtName As String) As String

    Dim strName As String   '�t�@�C�����܂��̓f�B���N�g�������i�[
    
    
    '�p�X���̍Ōオ"\"�Ȃ�͂���
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If
    
    '�t�@�C�����܂��̓f�B���N�g�������i�[
    strName = Dir$(strPathName, 31)
    
    
    '�����AstrName��"."���܂܂�ĂȂ���΂��̂܂�".zip"�����ĕԂ�
    If InStrRev(strName, ".") = 0 Then
        
        PackName = strName & strExtName
    
    Else
        
        PackName = Left$(strName, InStrRev(strName, ".") - 1) & strExtName
    
    End If
        
End Function


