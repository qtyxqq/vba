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

Sub cmd_exec()
    Dim sh  As New IWshRuntimeLibrary.WshShell  '// WshShell�N���X�I�u�W�F�N�g
    Dim ex  As WshExec                          '// Exec���\�b�h�߂�l
    Dim sArCmd(2)                               '// ���s�R�}���h�z��
    Dim sCmd                                    '// ���s�R�}���h
    Dim i                                       '// ���[�v�J�E���^
    
    Set WSH = CreateObject("WScript.Shell")
    
    sCmdFile = "D:\�����[�X\sample.cmd;"
  
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
    
    
End Sub

