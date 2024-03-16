Option Explicit

Dim file
Dim objFileSys
Dim objReadStream
Dim strLine
Dim lngCnt

if WScript.Arguments.Count = 0 then
  WScript.echo("ファイルをこのツールにドラッグ＆ドロップしてください")
  WScript.Quit(-1)
end if

' ファイル名取得
file = Wscript.Arguments(0)

' FSオブジェクト
Set objFileSys = CreateObject("Scripting.FileSystemObject")

' ファイルオープン
Set objReadStream  = objFileSys.OpenTextFile(file, 1)

Do Until objReadStream.AtEndOfStream = True
  '1 行読み込み
  strLine = objReadStream.ReadLine
  lngCnt = lngCnt + Len(strLine)
Loop

WScript.echo "文字数：" & lngCnt

objReadStream.Close

Set objFileSys = Nothing 
