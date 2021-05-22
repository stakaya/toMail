Attribute VB_Name = "SubMain"
'
'LhaZip.bas  Ver 1.0
'
'機能：Lha形式
'必要なDLL:UnLHA32.dll
' https://micco.mars.jp/mysoft/unlha32.htm
'
'====================================================================
'1.Lha形式で圧縮する。
'
'関数：LzhPack(lngWnd , strPreFile , strDir , [strSwitch] , [strOption])
'
'引数   lngWnd      :対象となるウインドウハンドル(例：Form1.hWnd)
'       strPreFile  :圧縮したいファイルのパス(例: c:\My Documents\test.txt)
'       strDir      :圧縮ファイルの出力先(例: C:\My Documents\)
'       strSwitch   :コマンドラインに渡すスイッチ(省略可:デフォルトでは"u"  上記参照)
'       strOption   :コマンドラインに渡すオプション(省略可:デフォルトでは"-a1 -r2 -x1 -l1 -jp -o2"　上記参照)
'
'
'戻り値：正常；0　　失敗；1
'====================================================================
'
'
'
'====================================================================


Option Explicit

'APIを宣言
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
        Call MsgBox("対象ファイルをドラッグしてください。", vbOKOnly + vbExclamation, "メール")
    End If
    End
End Sub
Function GetCommandLine(Optional MaxArgs)
   ' 変数を宣言します。
   Dim c, CmdLine, CmdLnLen, InArg, i, NumArgs
   ' MaxArgs が提供されるかどうかを調べます。
   If IsMissing(MaxArgs) Then MaxArgs = 10
   ' 現在のサイズの配列にします。
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   ' コマンド ラインの引数を取得します。
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   ' 同時にコマンド ラインの引数を取得します。
   For i = 1 To CmdLnLen
      c = Mid(CmdLine, i, 1)
      ' スペースまたはタブを調べます。
      If (c <> " " And c <> vbTab) Then
         ' スペースまたはタブのいずれでもありません。
         ' 既に引数の中ではないかどうかを調べます。
         If Not InArg Then
         ' 新しい引数が始まります。
         ' 引数が多すぎないかを調べます。
            If NumArgs = MaxArgs Then Exit For
               NumArgs = NumArgs + 1
               InArg = True
            End If
         ' 現在の引数に文字を追加します。
         ArgArray(NumArgs) = ArgArray(NumArgs) & c
      Else
         ' スペースまたはタブを見つけました。
         ' InArg フラグに False を設定します。
         InArg = False
      End If
   Next i
   ' 引数がすべて格納できるように配列のサイズを変更します。
   ReDim Preserve ArgArray(NumArgs)
   ' 関数名に配列を返します。
   GetCommandLine = ArgArray()
End Function
                                        

'
'Lzh形式で圧縮を行う
'
Public Function LzhPack(ByVal lngWnd As Long, _
                        ByVal strPreFile As String, _
                        ByVal strDir As String, _
                        Optional ByVal strSwitch As String = "u", _
                        Optional ByVal strOptions As String = "-a1 -r2 -x1 -l1 -jp -o2") As Long


    'lngWnd     :対象となるウインドウハンドル(例：Form1.hWnd)
    'strPreFile :圧縮したいファイルのパス(例: c:\My Documents\test.txt)
    'strDir     :圧縮ファイルの出力先(例: C:\My Documents\)
    'strSwitch  :コマンドラインとして渡すスイッチ(省略可:デフォルトでは"u"  上記参照)
    'strOption  :コマンドラインに渡すオプション(省略可:デフォルトでは"-a1 -r2 -x1 -l1 -jp -o2"　上記参照)
    
    
    Dim strLzhFile As String    '圧縮後のファイルのパス
    
    Dim strCommandLine As String
    Dim strBuffer As String * 1024
    
    Dim lngResult As Long   '戻り値を格納
    
    Dim pvHnd  As Long                      'ｷｰﾊﾝﾄﾞﾙ
    Dim pvType As Long                      'ﾀｲﾌﾟ
    Dim pvStr  As String * 1024             'ﾊﾞｯﾌｧ
    Dim pvRet  As Long                      '戻り値
    Dim mail   As String
    Dim shel   As String
    Dim flag   As Integer
    Dim i      As Integer
    
    'ドライブを圧縮しようとしたときはエラーを返す。
    If Len(strPreFile) < 4 Then
        LzhPack = 1
        Call MsgBox("ドライブは圧縮できません。No.1", vbOKOnly, "LzhPackエラー")
        Exit Function
    End If
    
    '圧縮しようとしているファイルが存在しないときはエラーを返す。
    If Len(Dir$(strPreFile, 31)) = 0 Then
        LzhPack = 1
        Call MsgBox("ファイルが見つかりません。No.2", vbOKOnly, "LzhPackエラー")
        Exit Function
    End If
    
    '圧縮後のファイルのパスを取得
    strLzhFile = AddDirSep(strDir) & PackName(strPreFile, ".lzh")
    
    'コマンドライン
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

            
    '圧縮の実行:正常リターン=0
    lngResult = Unlha(lngWnd, strCommandLine, strBuffer, Len(strBuffer))
    'Call MsgBox(strCommandLine)
    
    '値を返す
    If lngResult <> 0 Then
        Call MsgBox("圧縮に失敗しました。No.3", vbOKOnly + vbExclamation, "圧縮失敗")
        LzhPack = 1
        Exit Function
    Else
        LzhPack = 0
    End If
    
    If RegOpenKey(&H80000000, "mailto\shell\open\command", pvHnd) = 0 Then
        If RegQueryValueEx(pvHnd, "", 0, pvType, ByVal pvStr, 1024) = 0 Then
                                            '文字列の抽出
            mail = Left(pvStr, InStr(pvStr, Chr(0)) - 1)
        End If
    End If
    pvRet = RegCloseKey(pvHnd)              'ﾚｼﾞｽﾄﾘｸﾛｰｽﾞ

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
'パス名にスペースやカンマが含まれている場合にこのパス名に二重引用符を付加して返す。
'
Private Function AddQuotesToFN(ByVal strFileName As String) As String

    If InStr(strFileName, " ") Or InStr(strFileName, ".") Then
        AddQuotesToFN = """" & strFileName & """"
    Else
        AddQuotesToFN = strFileName
    End If
    
End Function

'
'パスの末尾に区切り記号（￥）がないときには（￥）をつける
'
Private Function AddDirSep(strPathName As String) As String

    If Right$(strPathName, 1) = "\" Then
        AddDirSep = RTrim$(strPathName)
    Else
        AddDirSep = RTrim$(strPathName) & "\"
    End If
    
End Function


'圧縮前のファイルの場所を取得
Private Function GetPath(ByVal strPathName As String) As String

    'パス名の最後が"\"ならはずす
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If
    
    'strPathNameがドライブならそのまま返す
    If Right$(strPathName, 1) = ":" Then
        
        GetPath = strPathName
        
        Exit Function
        
    End If
    
    
    '戻り値を格納

    GetPath = Left$(strPathName, InStrRev(strPathName, "\") - 1)

End Function


'圧縮後のファイル名を作成
Private Function PackName(ByVal strPathName As String, ByVal strExtName As String) As String

    Dim strName As String   'ファイル名またはディレクトリ名を格納
    
    
    'パス名の最後が"\"ならはずす
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If
    
    'ファイル名またはディレクトリ名を格納
    strName = Dir$(strPathName, 31)
    
    
    'もし、strNameに"."が含まれてなければそのまま".zip"をつけて返す
    If InStrRev(strName, ".") = 0 Then
        
        PackName = strName & strExtName
    
    Else
        
        PackName = Left$(strName, InStrRev(strName, ".") - 1) & strExtName
    
    End If
        
End Function


