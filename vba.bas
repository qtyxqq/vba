Attribute VB_Name = "��?1"
Sub creatFileList()
Attribute creatFileList.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' �G1 �G
'
    ' ����?: Ctrl+b
    Dim sht As Worksheet
    Dim LastColumn As Long
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    file_name = "D:\�����[�X\release.txt"
    Set ts = fso.OpenTextFile(file_name, ForWriting, True)
    
    Set sht = ThisWorkbook.Worksheets("Sheet1")
    LastRow = sht.UsedRange.Rows(sht.UsedRange.Rows.Count).Row
    For i = 1 To LastRow
      txt_value = sht.Cells(i, 1)
      ts.WriteLine (txt_value & Chr(9) & "Modified") ' �Ō�ɉ��s��t���ď�������
    Next i
      ts.Close ' �t�@�C�������
      ' ��n��
      Set ts = Nothing
      Set fso = Nothing
End Sub

Sub cmdExec()
    Dim sh  As New IWshRuntimeLibrary.WshShell  '// WshShell�N���X�I�u�W�F�N�g
    Dim ex  As WshExec                          '// Exec���\�b�h�߂�l
    Dim sArCmd(2)                               '// ���s�R�}���h�z��
    Dim sCmd                                    '// ���s�R�}���h
    Dim i                                       '// ���[�v�J�E���^
    
    Set WSH = CreateObject("WScript.Shell")
    
    sCmdFile = "D:\�����[�X\sample_mutb.cmd;"
  
    '// ���s���鏇�ɃR�}���h��z��Ɋi�[
    sArCmd(0) = "cd /d D:\�����[�X\"
    sArCmd(1) = sCmdFile
    
    '// �R�}���h��[ & ]�ŘA��
    sCmd = sArCmd(0) & " & " & sArCmd(1)
        
    '// �R�}���h���s
    Set ex = sh.Exec("cmd.exe /c " & sCmd)
    
    '// �R�}���h���s��
    If (ex.Status = WshFailed) Then
        '// �����𔲂���
        Exit Sub
    End If
    
    '// �R�}���h���s���͑҂�
    Do While (ex.Status = WshRunning)
        DoEvents
    Loop
    
    If InStr(sCmdFile, "mutb") > 0 Then
        Call fileSplit
    End If
    
End Sub

Sub fileSplit()
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' �C���X�^���X��
    Dim txt As String
    Dim ts As TextStream
    
    Folder = "C:\Users\qtyxqq\OneDrive\�h�L�������g\GitHub\vba\"
    
    Set ts = fso.OpenTextFile(Folder & "release_result.txt")
    bl_ctx = "#!/bin/bash" & Chr(10)
    client_ctx = "#!/bin/bash" & Chr(10)
    fc_ctx = "#!/bin/bash" & Chr(10)
    ' �S�Ẵf�[�^���擾
    Do Until ts.AtEndOfStream ' �Ō�̍s���擾����܂�
        txt = ts.ReadLine ' 1 �s���擾
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
    ts.Close ' �t�@�C�������
    

    Set fso = New FileSystemObject
    
    bl_name = Folder & "bl_result.txt"
    client_name = Folder & "client_result.txt"
    fc_name = Folder & "fc_result.txt"
    If write_bl = True Then
        Set bl = fso.OpenTextFile(bl_name, ForWriting, True)
        bl.WriteLine (bl_ctx) '
        bl.Close ' �t�@�C�������
        Set bl = Nothing
    End If
    
    If write_client = True Then
        Set client = fso.OpenTextFile(client_name, ForWriting, True)
        client.WriteLine (client_ctx) '
        client.Close ' �t�@�C�������
        Set client = Nothing
    End If
    
    If write_fc = True Then
        Set fc = fso.OpenTextFile(fc_name, ForWriting, True)
        fc.WriteLine (fc_ctx) '
        fc.Close ' �t�@�C�������
        Set fc = Nothing
    End If
    

    ' ��n��
    Set fso = Nothing
    

End Sub


