Option Explicit

'====================================================================================================
'hi, mark ganson here (latest revision: 7/20/2015)
'reach me at mwganson@hotmail.com with comments (put md5.vbs in subject line please)
'source of this file: http://mwganson.freeyellow.com/md5/md5.vbs
'all of this stuff i found on the web, sources noted in the code where applicable
'i integrated it all into this script with very few modifications of my own

'this is a .vbs script for computing and displaying md5 hashes (along with sha1, sha256, sha384, sha512, and ripemd160)
'it uses a simple drag and drop interface (or you can use the send to context menu option -- read on)
'to wit: drag and drop your file(s) onto the MD5.vbs file and it will do the rest

'*should* work on all windows computers

'not the fastest way to do md5 hashes maybe, but still has its convenience advantages

'here's how to put this into your send to folder so you can right-click the file(s)
'you want to create md5 hashes for and send them to this script
'right-click on the md5.vbs script and choose create shortcut
'this produces a file named md5 shortuct.lnk or something similar
'rename it to md5.lnk (optional)

'paste this text into windows explorer address bar: %APPDATA%\Microsoft\Windows\SendTo
'you will now be in the send to folder -- simply move the shortcut md5.lnk file to this folder
'source of this information: http://www.howtogeek.com/howto/windows-vista/customize-the-windows-vista-send-to-menu/


'ENTRY POINT for this script
'users can double-click the script, drop files onto the script, or use the send to context menu option.  dropped files will be filenames as arguments
'parses command line arguments to get filename(s)
'following is something i modified from  the file copying script found at
'https://social.technet.microsoft.com/forums/scriptcenter/en-US/e032bc2e-31c4-4306-9f43-b5202d23e9d7/copy-files-using-drag-n-drop-on-vbs-file

Dim doMd5, doSha1, doSha256, doSha384, doSha512, doRipemd160
'set these values according to your needs
'setting unnecessary ones to False will greatly speed up the process
doMd5 = False
doSha1 = False
doSha256 = False
doSha384 = False
doSha512 = False
doRipemd160 = False

Dim objFile,objFolder,objFSO 
Dim Arg, strText 
Dim PDFPath

strText ="" 
Set objFSO = CreateObject("Scripting.FileSystemObject") 

' > 0 arguments means files were dropped onto the md5.vbs script (or sent to via send to option)
 
If WScript.Arguments.Count > 0 Then 
	Dim OlApp
	Dim Eml
	Dim File
	Set OlApp = CreateObject("Outlook.Application")

    For Each Arg in Wscript.Arguments 
        Arg =  Trim(Arg) 
		If InStr(Arg,".") Then 
			'' Assume a File
			File = Arg
			PDFPath = Left(File, InStr(UCase(File), "\MSG\") - 1) + "\PDF\"
			Set Eml = OlApp.CreateItemFromTemplate(File)
			Download(Eml)
		End If 
    Next
 
 End If



'improved md5 function uses .net md5 component (formerly had code here to do the hashing in the script)
'source: https://github.com/falconws/wshLib/blob/master/lib/MD5.vbs

Function md5(filename)
	Dim MSXML, EL, MD5Obj

    Set MD5Obj = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
	MD5Obj.ComputeHash_2(readBinaryFile(filename))

	Set MSXML = CreateObject("MSXML2.DOMDocument")
	Set EL = MSXML.CreateElement("tmp")
	EL.DataType = "bin.hex"
	EL.NodeTypedValue = MD5Obj.Hash
	md5 = EL.Text
End Function

Function sha1(filename)
	Dim MSXML, EL
	Set SHA1 = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
	SHA1.ComputeHash_2(readBinaryFile(filename))

	Set MSXML = CreateObject("MSXML2.DOMDocument")
	Set EL = MSXML.CreateElement("tmp")
	EL.DataType = "bin.hex"
	EL.NodeTypedValue = SHA1.Hash
	sha1 = EL.Text
End Function

Function sha256(filename)
	Dim MSXML, EL
	Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
	SHA256.ComputeHash_2(readBinaryFile(filename))

	Set MSXML = CreateObject("MSXML2.DOMDocument")
	Set EL = MSXML.CreateElement("tmp")
	EL.DataType = "bin.hex"
	EL.NodeTypedValue = SHA256.Hash
	sha256 = EL.Text
End Function

Function sha384(filename)
	Dim MSXML, EL
	Set SHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")
	SHA384.ComputeHash_2(readBinaryFile(filename))

	Set MSXML = CreateObject("MSXML2.DOMDocument")
	Set EL = MSXML.CreateElement("tmp")
	EL.DataType = "bin.hex"
	EL.NodeTypedValue = SHA384.Hash
	sha384 = EL.Text
End Function


Function sha512(filename)
	Dim MSXML, EL
	Set SHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
	SHA512.ComputeHash_2(readBinaryFile(filename))

	Set MSXML = CreateObject("MSXML2.DOMDocument")
	Set EL = MSXML.CreateElement("tmp")
	EL.DataType = "bin.hex"
	EL.NodeTypedValue = SHA512.Hash
	sha512 = EL.Text
End Function


Function ripemd160(filename)
	Dim MSXML, EL
	Set ripemd160 = CreateObject("System.Security.Cryptography.ripemd160Managed")
	ripemd160.ComputeHash_2(readBinaryFile(filename))

	Set MSXML = CreateObject("MSXML2.DOMDocument")
	Set EL = MSXML.CreateElement("tmp")
	EL.DataType = "bin.hex"
	EL.NodeTypedValue = ripemd160.Hash
	ripemd160 = EL.Text
End Function

Function readBinaryFile(filename)
	Const adTypeBinary = 1
	Dim objStream
	Set objStream = CreateObject("ADODB.Stream")
	objStream.Type = adTypeBinary
	objStream.Open
	If filename <> "" Then objStream.LoadFromFile filename End If 'slight modification here to prevent error msg if no file selected
	readBinaryFile = objStream.Read
	objStream.Close
	Set objStream = Nothing
End Function

'====================================================================================================

Sub Download(objEml)
	Dim Attch
	Dim objFso  
	Set objFso= CreateObject("Scripting.FileSystemObject") 
	
	Dim myFormat
	Dim t
	Dim temp
	Dim milliseconds
	Dim sb
	
	For Each Attch In objEml.Attachments
		myFormat = "yyyyMMddhhmmss"
		t = Timer
		temp = Int(t)
		milliseconds = Right("00" & Int((t-temp) * 1000),3)

		' formatting the date 
		Set sb = createobject("System.Text.StringBuilder")
		sb.AppendFormat "{0:" & myFormat & "}" & milliseconds, Now

		' passing the result
		' WScript.Echo sb.ToString()


		' Attch.SaveAsFile PDFPath & Attch.FileName
		If UCase(objFSO.GetExtensionName(Attch)) = "PDF" Then
			'Attch.SaveAsFile PDFPath & sb.ToString() & ".PDF"

			Attch.SaveAsFile PDFPath & Attch.FileName
			'Dim strmd5 
			'strmd5 = md5(PDFPath & sb.ToString() & ".PDF")
			'objFso.DeleteFile PDFPath & sb.ToString() & ".PDF"
			'If objFso.FileExists(PDFPath & strmd5 & ".PDF") Then
			'	objFso.DeleteFile PDFPath & strmd5 & ".PDF"
			'End if
			'Attch.SaveAsFile PDFPath & strmd5 & ".PDF"
		End If
	Next
End Sub

'this function we use only if the user simply double-clicked on the md5.vbs script file
'instead of dragging and dropping or using the send to interface
'argument count will be 0
'source: http://todayguesswhat.blogspot.com/2012/08/windows-7-replacement-for.html

dim shell, defaultLocalDir, objWMIService, colItems, objItem, ex

Set shell = CreateObject( "WScript.Shell" )
defaultLocalDir = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop"
Set shell = Nothing
