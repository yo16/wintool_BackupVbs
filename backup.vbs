' backup�c�[��
' version 10.0
'
' �X�V����
' v5.0 2005/11/10 ������^�����ꍇ�A���̈����̃t�H���_��
'                 �J�����g�t�H���_�Ƃ��ċN������@�\��ǉ��B
' v6.0 2006/01/27 BackUp�t�H���_�̉��ɁA���t���Ƃ̃t�H���_��
'                 �쐬����@�\��ǉ��B
' v7.0 2006/02/21 OS�̃��O�C�����[�U�����L�^����t�@�C�����쐬�B
' v7.1 2006/02/27 ���O�C�����[�U���́A�t�@�C�������̂ɏ����悤�ύX�B
' v8.0 2007/11/07 ���k�Ɏ��Ԃ������邽�߁ATop�̃t�H���_�����k����Ă�����
'                 ���k�����𑖂点�Ȃ��悤�ύX�B
' v9.0 2008/05/14 ���k�����̃^�C�~���O��ύX���A�Ώۂ���t�t�H���_�����ɕύX�B
' v10.0 2015/12/22 �t�H���_�����k�ł���t���O��ǉ��B

Option Explicit


'-------------------
' �����ݒ�
'-------------------
' �o�b�N�A�b�v�t�H���_��
Dim backupDirName
backupDirName = "_Backup"

' �m�F���b�Z�[�W���o���t���O
Dim messageFlag
messageFlag = false

' ���t���Ƃ̃t�H���_�����t���O
Dim makeDateFolderFlag
makeDateFolderFlag = true

' ���O�C�����[�U���̃t�@�C�����i�w�b�_�j
Dim loginUserFileNameHeader
loginUserFileNameHeader = "_user_"

' �t�H���_�����k����t���O
Dim compressDir
compressDir = true



'-------------------
' �N���m�F���b�Z�[�W
'-------------------
if ( messageFlag ) Then
	Dim YNmodori
	YNmodori = MsgBox("�o�b�N�A�b�v������Ă������ł����H",4,"�t�H���_���ƃo�b�N�A�b�v")
	If (YNmodori <> 6) Then
		WScript.Quit
	End If
End If



Dim objFS
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")

'-------------------
' �������������ꍇ�A��P�������J�����g�t�H���_�Ƃ���
'-------------------
Dim objArgs
Dim CurrentFolderName
Set objArgs = WScript.Arguments
If ( objArgs.Count > 0 ) Then
	CurrentFolderName = objArgs(0)
	' �t�H���_���݊m�F
	If ( objFS.FolderExists(CurrentFolderName) = 0 ) Then
		MsgBox "�w�肳�ꂽ�t�H���_["&CurrentFolderName&"]�����݂��܂���", vbOKOnly, "backup"
		WScript.Quit
	End If
	' ������\���Ȃ�������ǉ�����
	If ( Right(CurrentFolderName, 1) <> "\" ) Then
		CurrentFolderName = CurrentFolderName & "\"
	End If
Else
	CurrentFolderName = ".\"
End If



'-------------------
' �쐬����o�b�N�A�b�v�t�H���_��
'-------------------
Dim motoBKFolderName
motoBKFolderName = CurrentFolderName & backupDirName


'***BackUp�t�H���_��u���t�H���_���쐬***
If (objFS.FolderExists(motoBKFolderName) = 0) Then
	objFS.CreateFolder motoBKFolderName
End If


'***BackUp�p�t�H���_�������***'
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

'***���t���Ƃ̃t�H���_�����***
Dim dateFolderName
If (makeDateFolderFlag) Then
	dateFolderName = Right(fYear,2) & "-" & fMonth & "-" & fDay
	dateFolderName = motoBKFolderName & "\" & dateFolderName
	
	' �Ȃ���������
	If (objFS.FolderExists(dateFolderName) = 0) Then
		objFS.CreateFolder dateFolderName
	End If
	
	' �p�X��ς��Ă���
	motoBKFolderName = dateFolderName
	
End If


'���t�E���Ԃ��Ƃ̃t�H���_�����
Dim bkFolderName
bkFolderName = Right(fYear,2) & "-" & fMonth & "-" & fDay & "-" & fHour & fMinute & "-" & fSecond
bkFolderName = motoBKFolderName & "\" & bkFolderName
If (objFS.FolderExists(bkFolderName) = 0) Then
	objFS.CreateFolder bkFolderName
Else
	MsgBox "�t�H���_["&bkFolderName&"]�͊��ɑ��݂��܂��I�I"
	WScript.Quit
End If


'***�t�@�C�����u����Ă���t�H���_����Files�I�u�W�F�N�g���擾
Dim objFolder,colFiles
Set objFolder = objFS.GetFolder( CurrentFolderName )
Set colFiles = objFolder.Files

'***�t�@�C�����o�b�N�A�b�v�t�H���_�փR�s�[***
Dim x
For Each x in colFiles
	If (x.Name <> WScript.ScriptName) Then
		x.Copy bkFolderName&"\"&x.Name
	End If
Next

'***�t�H���_�I�u�W�F�N�g���擾
Dim colDirs
Set colDirs = objFolder.SubFolders
'***�t�H���_���o�b�N�A�b�v�t�H���_�փR�s�[***
For Each x in colDirs
	If (x.Name <> backupDirName) Then
		objFS.CopyFolder _
			CurrentFolderName & x.Name, _
			bkFolderName&"\"&x.Name, _
			True
	End If
Next


'***���O�C�����[�U�����擾���ăt�@�C�������***
Dim WSHShell,WSHEnv,strList,strEnv
Set WSHShell = WScript.CreateObject("WScript.Shell")
Set WSHEnv = WshShell.Environment("PROCESS")
Dim userName
userName = WSHEnv.Item("USERNAME")
' �t�@�C����
Dim userFilePath, loginUserFileName
loginUserFileName = loginUserFileNameHeader & userName
userFilePath = bkFolderName&"\"&loginUserFileName
Dim objTS
Dim overWrite
overWrite = True
' ���݊m�F�i���ɂ���ꍇ�͖ق��ĂȂɂ����Ȃ��j
If (objFS.FileExists(userFilePath) = 0) Then
	Set objTS = objFS.CreateTextFile(userFilePath,overWrite)
	objTS.WriteLine userName
End If


'***�t�H���_�����k����***
compressFolder bkFolderName



'MsgBox "�o�b�N�A�b�v�t�H���_["&bkFolderName&"]�����܂����I�I",,Now




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


