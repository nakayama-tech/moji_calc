Option Explicit

Dim file
Dim objFileSys
Dim objReadStream
Dim strLine
Dim lngCnt

if WScript.Arguments.Count = 0 then
  WScript.echo("�t�@�C�������̃c�[���Ƀh���b�O���h���b�v���Ă�������")
  WScript.Quit(-1)
end if

' �t�@�C�����擾
file = Wscript.Arguments(0)

' FS�I�u�W�F�N�g
Set objFileSys = CreateObject("Scripting.FileSystemObject")

' �t�@�C���I�[�v��
Set objReadStream  = objFileSys.OpenTextFile(file, 1)

Do Until objReadStream.AtEndOfStream = True
  '1 �s�ǂݍ���
  strLine = objReadStream.ReadLine
  lngCnt = lngCnt + Len(strLine)
Loop

WScript.echo "�������F" & lngCnt

objReadStream.Close

Set objFileSys = Nothing 
