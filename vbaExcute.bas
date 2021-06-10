Attribute VB_Name = "模?1"
Sub creatFileList()
Attribute creatFileList.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' 宏1 宏
'
' 快捷?: Ctrl+b
Dim sht As Worksheet
Dim LastColumn As Long
Dim fso As FileSystemObject
Set fso = New FileSystemObject

file_name = "D:\リリース\release.txt"
Set ts = fso.OpenTextFile(file_name, ForWriting, True)

Set sht = ThisWorkbook.Worksheets("Sheet1")
  LastRow = sht.UsedRange.Rows(sht.UsedRange.Rows.Count).Row
  For i = 1 To LastRow
    txt_value = sht.Cells(i, 1)
    ts.WriteLine (txt_value & Chr(9) & "Modified") ' 最後に改行を付けて書き込み
  Next i
    ts.Close ' ファイルを閉じる
    ' 後始末
    Set ts = Nothing
    Set fso = Nothing
End Sub

Sub cmd_exec()
    Dim sh  As New IWshRuntimeLibrary.WshShell  '// WshShellクラスオブジェクト
    Dim ex  As WshExec                          '// Execメソッド戻り値
    Dim sArCmd(2)                               '// 実行コマンド配列
    Dim sCmd                                    '// 実行コマンド
    Dim i                                       '// ループカウンタ
    
    Set WSH = CreateObject("WScript.Shell")
    
    sCmdFile = "D:\リリース\sample.cmd;"
  
    '// 実行する順にコマンドを配列に格納
    sArCmd(0) = "cd /d D:\リリース\"
    sArCmd(1) = sCmdFile
    
    '// コマンドを[ & ]で連結
    sCmd = sArCmd(0) & " & " & sArCmd(1)
        
    '// コマンド実行
    Set ex = sh.Exec("cmd.exe /c " & sCmd)
    
    '// コマンド失敗時
    If (ex.Status = WshFailed) Then
        '// 処理を抜ける
        Exit Sub
    End If
    
    '// コマンド実行中は待ち
    Do While (ex.Status = WshRunning)
        DoEvents
    Loop
    
    
End Sub

