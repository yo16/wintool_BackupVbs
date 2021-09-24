' backupツール
' version 10.0
'
' 更新履歴
' v5.0 2005/11/10 引数を与えた場合、その引数のフォルダを
'                 カレントフォルダとして起動する機能を追加。
' v6.0 2006/01/27 BackUpフォルダの下に、日付ごとのフォルダを
'                 作成する機能を追加。
' v7.0 2006/02/21 OSのログインユーザ名を記録するファイルを作成。
' v7.1 2006/02/27 ログインユーザ名は、ファイル名自体に書くよう変更。
' v8.0 2007/11/07 圧縮に時間がかかるため、Topのフォルダが圧縮されていたら
'                 圧縮処理を走らせないよう変更。
' v9.0 2008/05/14 圧縮処理のタイミングを変更し、対象を日付フォルダだけに変更。
' v10.0 2015/12/22 フォルダも圧縮できるフラグを追加。

Option Explicit


'-------------------
' 初期設定
'-------------------
' バックアップフォルダ名
Dim backupDirName
backupDirName = "_Backup"

' 確認メッセージを出すフラグ
Dim messageFlag
messageFlag = false

' 日付ごとのフォルダを作るフラグ
Dim makeDateFolderFlag
makeDateFolderFlag = true

' ログインユーザ情報のファイル名（ヘッダ）
Dim loginUserFileNameHeader
loginUserFileNameHeader = "_user_"

' フォルダも圧縮するフラグ
Dim compressDir
compressDir = true



'-------------------
' 起動確認メッセージ
'-------------------
if ( messageFlag ) Then
	Dim YNmodori
	YNmodori = MsgBox("バックアップを取ってもいいですか？",4,"フォルダごとバックアップ")
	If (YNmodori <> 6) Then
		WScript.Quit
	End If
End If



Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

'-------------------
' 引数があった場合、第１引数をカレントフォルダとする
'-------------------
Dim objArgs
Dim CurrentFolderName
Set objArgs = WScript.Arguments
If ( objArgs.Count > 0 ) Then
	CurrentFolderName = objArgs(0)
	' フォルダ存在確認
	If ( objFS.FolderExists(CurrentFolderName) = 0 ) Then
		MsgBox "指定されたフォルダ["&CurrentFolderName&"]が存在しません", vbOKOnly, "backup"
		WScript.Quit
	End If
	' 末尾に\がなかったら追加する
	If ( Right(CurrentFolderName, 1) <> "\" ) Then
		CurrentFolderName = CurrentFolderName & "\"
	End If
Else
	CurrentFolderName = ".\"
End If



'-------------------
' 作成するバックアップフォルダ名
'-------------------
Dim motoBKFolderName
motoBKFolderName = CurrentFolderName & backupDirName


'***BackUpフォルダを置くフォルダを作成***
If (objFS.FolderExists(motoBKFolderName) = 0) Then
	objFS.CreateFolder motoBKFolderName
End If


'***BackUp用フォルダ名を作る***'
Dim fYear,fMonth,fDay,fHour,fMinute,fSecond
fYear   = Year(Now)
fMonth  = Month(Now)
fDay    = Day(Now)
fHour   = Hour(Now)
fMinute = Minute(Now)
fSecond = Second(Now)
If (fYear < 10) Then fYear = "0" & fYear
If (fMonth < 10) Then fMonth = "0" & fMonth
If (fDay < 10) Then fDay = "0" & fDay
If (fHour < 10) Then fHour = "0" & fHour
If (fMinute < 10) Then fMinute = "0" & fMinute
If (fSecond < 10) Then fSecond = "0" & fSecond

'***日付ごとのフォルダを作る***
Dim dateFolderName
If (makeDateFolderFlag) Then
	dateFolderName = Right(fYear,2) & "-" & fMonth & "-" & fDay
	dateFolderName = motoBKFolderName & "\" & dateFolderName
	
	' なかったら作る
	If (objFS.FolderExists(dateFolderName) = 0) Then
		objFS.CreateFolder dateFolderName
	End If
	
	' パスを変えておく
	motoBKFolderName = dateFolderName
	
End If


'日付・時間ごとのフォルダを作る
Dim bkFolderName
bkFolderName = Right(fYear,2) & "-" & fMonth & "-" & fDay & "-" & fHour & fMinute & "-" & fSecond
bkFolderName = motoBKFolderName & "\" & bkFolderName
If (objFS.FolderExists(bkFolderName) = 0) Then
	objFS.CreateFolder bkFolderName
Else
	MsgBox "フォルダ["&bkFolderName&"]は既に存在します！！"
	WScript.Quit
End If


'***ファイルが置かれているフォルダ内のFilesオブジェクトを取得
Dim objFolder,colFiles
Set objFolder = objFS.GetFolder( CurrentFolderName )
Set colFiles = objFolder.Files

'***ファイルをバックアップフォルダへコピー***
Dim x
For Each x in colFiles
	If (x.Name <> WScript.ScriptName) Then
		x.Copy bkFolderName&"\"&x.Name
	End If
Next

'***フォルダオブジェクトを取得
Dim colDirs
Set colDirs = objFolder.SubFolders
'***フォルダをバックアップフォルダへコピー***
For Each x in colDirs
	If (x.Name <> backupDirName) Then
		objFS.CopyFolder _
			CurrentFolderName & x.Name, _
			bkFolderName&"\"&x.Name, _
			True
	End If
Next


'***ログインユーザ名を取得してファイルを作る***
Dim WSHShell,WSHEnv,strList,strEnv
Set WSHShell = WScript.CreateObject("WScript.Shell")
Set WSHEnv = WshShell.Environment("PROCESS")
Dim userName
userName = WSHEnv.Item("USERNAME")
' ファイル名
Dim userFilePath, loginUserFileName
loginUserFileName = loginUserFileNameHeader & userName
userFilePath = bkFolderName&"\"&loginUserFileName
Dim objTS
Dim overWrite
overWrite = True
' 存在確認（既にある場合は黙ってなにもしない）
If (objFS.FileExists(userFilePath) = 0) Then
	Set objTS = objFS.CreateTextFile(userFilePath,overWrite)
	objTS.WriteLine userName
End If


'***フォルダを圧縮する***
compressFolder bkFolderName



'MsgBox "バックアップフォルダ["&bkFolderName&"]を作りました！！",,Now




Sub compressFolder(folderName)
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")

	Dim objFS,objFolder
	Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFS.GetFolder(folderName)
	
	Dim nCompressed
	nCompressed = 2048
	If (objFolder.Attributes and nCompressed) Then
'		msgbox "nocompress"
	Else
'		msgbox "COMPRESS!"&objFolder.path
		WshShell.Run "cmd /c COMPACT /C /S:"""&objFolder.path&"""",0,1
	End If
End Sub


