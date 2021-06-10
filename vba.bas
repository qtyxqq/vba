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

Sub cmdExec()
    Dim sh  As New IWshRuntimeLibrary.WshShell  '// WshShellクラスオブジェクト
    Dim ex  As WshExec                          '// Execメソッド戻り値
    Dim sArCmd(2)                               '// 実行コマンド配列
    Dim sCmd                                    '// 実行コマンド
    Dim i                                       '// ループカウンタ
    
    Set WSH = CreateObject("WScript.Shell")
    
    sCmdFile = "D:\リリース\sample_mutb.cmd;"
  
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
    
    If InStr(sCmdFile, "mutb") > 0 Then
        Call fileSplit
    End If
    
End Sub

Sub fileSplit()
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' インスタンス化
    Dim txt As String
    Dim ts As TextStream
    
    Folder = "C:\Users\qtyxqq\OneDrive\ドキュメント\GitHub\vba\"
    
    Set ts = fso.OpenTextFile(Folder & "release_result.txt")
    bl_ctx = "#!/bin/bash" & Chr(10)
    client_ctx = "#!/bin/bash" & Chr(10)
    fc_ctx = "#!/bin/bash" & Chr(10)
    ' 全てのデータを取得
    Do Until ts.AtEndOfStream ' 最後の行を取得するまで
        txt = ts.ReadLine ' 1 行ずつ取得
        If InStr(txt, "client start") Then
            client_flg = True
            fc_flg = False
            bl_flg = False
        ElseIf InStr(txt, "bl start") Then
            client_flg = False
            bl_flg = True
            fc_flg = False
        ElseIf InStr(txt, "fc start") Then
            client_flg = False
            fc_flg = True
            bl_flg = False
        End If
        If client_flg = True Then
              client_ctx = client_ctx & txt & Chr(10)
              write_client = True
        ElseIf bl_flg = True Then
              bl_ctx = bl_ctx & txt & Chr(10)
              write_bl = True
        ElseIf fc_flg = True Then
              fc_ctx = fc_ctx & txt & Chr(10)
              write_fc = True
        End If
    Loop
    ts.Close ' ファイルを閉じる
    

    Set fso = New FileSystemObject
    
    bl_name = Folder & "bl_result.txt"
    client_name = Folder & "client_result.txt"
    fc_name = Folder & "fc_result.txt"
    If write_bl = True Then
        Set bl = fso.OpenTextFile(bl_name, ForWriting, True)
        bl.WriteLine (bl_ctx) '
        bl.Close ' ファイルを閉じる
        Set bl = Nothing
    End If
    
    If write_client = True Then
        Set client = fso.OpenTextFile(client_name, ForWriting, True)
        client.WriteLine (client_ctx) '
        client.Close ' ファイルを閉じる
        Set client = Nothing
    End If
    
    If write_fc = True Then
        Set fc = fso.OpenTextFile(fc_name, ForWriting, True)
        fc.WriteLine (fc_ctx) '
        fc.Close ' ファイルを閉じる
        Set fc = Nothing
    End If
    

    ' 後始末
    Set fso = Nothing
    

End Sub


