<!--#include file = "private.asp"-->

<%
'######################################
' eWebEditor v3.00 - Advanced online web based WYSIWYG HTML editor.
' Copyright (c) 2003-2004 eWebSoft.com
'
' For further information go to http://www.ewebsoft.com/
' This copyright notice MUST stay intact for use.
'######################################
%>

<%

Dim sStyleID, sUploadDir, sCurrDir, sDir

sPosition = sPosition & "上传文件管理"

Call Header()
Call Content()
Call Footer()


Sub Content()
	If IsObjInstalled("Scripting.FileSystemObject") = False Then
		Response.Write "此功能要求服务器支持文件系统对象（FSO），而你当前的服务器不支持！"
		Exit Sub
	End If

	Call InitParam()

	Select Case sAction
	Case "DELALL"
		Call DoDelAll()
	Case "DEL"
		Call DoDel()
	Case "DELFOLDER"
		Call DoDelFolder()
	End Select

	Call ShowList()
End Sub

Sub ShowList()
	Response.Write "<table border=0 cellspacing=1 align=center class=navi>" & _
		"<form action='?' method=post name=queryform>" & _
		"<tr><th>" & sPosition & "</th></tr>" & _
		"<tr><td align=right><b>选择样式目录：</b><select name='id' size=1 onchange=""location.href='?id='+this.options[this.selectedIndex].value"">" & InitSelect(sStyleID, "选择...") & "</select></td></tr>" & _
		"</form></table><br>"
	
	If sCurrDir = "" Then Exit Sub

	Response.Write "<table border=0 cellspacing=1 class=list align=center>" & _
		"<form action='?id=" & sStyleID & "&dir=" & sDir & "&action=del' method=post name=myform>" & _
		"<tr align=center>" & _
			"<th width='10%'>类型</th>" & _
			"<th width='40%'>文件地址</th>" & _
			"<th width='10%'>大小</th>" & _
			"<th width='15%'>最后访问</th>" & _
			"<th width='15%'>上传日期</th>" & _
			"<th width='10%'>删除</th>" & _
		"</tr>"

	Dim sCurrPage, nCurrPage, nFileNum, nPageNum, nPageSize
	sCurrPage = Trim(Request("page"))
	nPageSize = 20
	If sCurrpage = "" Or Not IsNumeric(sCurrPage) Then
		nCurrPage = 1
	Else
		nCurrPage = CLng(sCurrPage)
	End If

	Dim oFSO, oUploadFolder, oUploadFiles, oUploadFile, sFileName

	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	Set oUploadFolder = oFSO.GetFolder(Server.MapPath(sCurrDir))
	If Err.Number>0 Then
		Response.Write "<tr><td colspan=6>无效的目录！</td></tr></table>"
		Exit Sub
	End If

	
	If sDir <> "" Then
		Response.Write "<tr align=center>" & _
			"<td><img border=0 src='../sysimage/file/folderback.gif'></td>" & _
			"<td align=left colspan=5><a href=""?id=" & sStyleID & "&dir="
		If InstrRev(sDir, "/") > 1 Then
			Response.Write Left(sDir, InstrRev(sDir, "/") - 1)
		End If
		Response.Write """>返回上一级目录</a></td></tr>"
	End If

	Dim oSubFolder
	For Each oSubFolder In oUploadFolder.SubFolders
		Response.Write "<tr align=center>" & _
			"<td><img border=0 src='../sysimage/file/folder.gif'></td>" & _
			"<td align=left colspan=4><a href=""?id=" & sStyleID & "&dir="
		If sDir <> "" Then
			Response.Write sDir & "/"
		End If
		Response.Write oSubFolder.Name & """>" & oSubFolder.Name & "</a></td>" & _
			"<td><a href='?id=" & sStyleID & "&dir=" & sDir & "&action=delfolder&foldername=" & oSubFolder.Name & "'>删除</a></td></tr>"
	Next

	
	Set oUploadFiles = oUploadFolder.Files

	nFileNum = oUploadFiles.Count
	nPageNum = Int(nFileNum / nPageSize)
	If nFileNum Mod nPageSize > 0 Then
		nPageNum = nPageNum+1
	End If
	If nCurrPage > nPageNum Then
		nCurrPage = 1
	end If

	Dim i
	i = 0
	For Each oUploadFile In oUploadFiles
		i = i + 1
		If i > (nCurrPage - 1) * nPageSize And i <= nCurrPage * nPageSize Then
			sFileName = oUploadFile.Name
			Response.Write "<tr align=center>" & _
				"<td>" & FileName2Pic(sFileName) & "</td>" & _
				"<td align=left><a href=""" & sCurrDir & sFileName & """ target=_blank>" & sFileName & "</a></td>" & _
				"<td>" & oUploadFile.size & " B </td>" & _
				"<td>" & oUploadFile.datelastaccessed & "</td>" & _
				"<td>" & oUploadFile.datecreated & "</td>" & _
				"<td><input type=checkbox name=delfilename value=""" & sFileName & """></td></tr>"
		Elseif i > nCurrPage * nPageSize Then
			Exit For
		End If
	Next
	Set oUploadFolder = Nothing
	Set oUploadFiles = Nothing

	If nFileNum <= 0 Then
		Response.Write "<tr><td colspan=6>指定目录下现在还没有文件！</td></tr>"
	End If
	

	If nFileNum > 0 Then
		Response.Write "<tr><td colspan=6><table border=0 cellpadding=3 cellspacing=0 width='100%'><tr><td>"
		If nCurrPage > 1 Then
			Response.Write "<a href='?id=" & sStyleID & "&dir=" & sDir & "&page=1'>首页</a>&nbsp;&nbsp;<a href='?id=" & sStyleID & "&dir=" & sDir & "&page="& nCurrPage - 1 & "'>上一页</a>&nbsp;&nbsp;"
		Else
			Response.Write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
		End If
		If nCurrPage < i / nPageSize Then
			Response.Write "<a href='?id=" & sStyleID & "&dir=" & sDir & "&page=" & nCurrPage + 1 & "'>下一页</a>&nbsp;&nbsp;<a href='?id=" & sStyleID & "&dir=" & sDir & "&page=" & nPageNum & "'>尾页</a>"
		Else
			Response.Write "下一页&nbsp;&nbsp;尾页"
		End If
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;共<b>" & nFileNum & "</b>个&nbsp;&nbsp;页次:<b><span class=highlight2>" & nCurrPage & "</span>/" & nPageNum & "</b>&nbsp;&nbsp;<b>" & nPageSize & "</b>个文件/页"
		Response.Write "</td><td align=right><input type=submit name=b value=' 删除选定的文件 '> <input type=button name=b1 value=' 清空所有文件 ' onclick=""javascript:if (confirm('你确定要清空所有文件吗？')) {location.href='?id=" & sStyleID & "&dir=" & sDir & "&action=delall';}""></td></tr></table></td></tr>"
	End If

	Response.Write "</form></table>"


End Sub

Sub DoDel()
	On Error Resume Next
	Dim sFileName, oFSO, sMapFileName
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	For Each sFileName In Request.Form("delfilename")
		sMapFileName = Server.MapPath(sCurrDir & sFileName)
		If oFSO.FileExists(sMapFileName) Then
			oFSO.DeleteFile(sMapFileName)
		End If
	Next
	Set oFSO = Nothing
End Sub

Sub DoDelAll()
	On Error Resume Next
	Dim sFileName, oFSO, sMapFileName, oFolder, oFiles, oFile
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set oFolder = oFSO.GetFolder(Server.MapPath(sCurrDir))
	Set oFiles = oFolder.Files
	For Each oFile In oFiles
		sFileName = oFile.Name
		sMapFileName = Server.MapPath(sCurrDir & sFileName)
		If oFSO.FileExists(sMapFileName) Then
			oFSO.DeleteFile(sMapFileName)
		End If
	Next
	Set oFile = Nothing
	Set oFolder = Nothing
	Set oFSO = Nothing
End Sub

Sub DoDelFolder()
	On Error Resume Next
	Dim sFolderName, oFSO, sMapFolderName
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	sFolderName = Trim(Request("foldername"))
	sMapFolderName = Server.Mappath(sCurrDir & sFolderName)
	If oFSO.FolderExists(sMapFolderName) = True Then
		oFSO.DeleteFolder(sMapFolderName)
	End If
	Set oFSO = Nothing
End Sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function FileName2Pic(sFileName)
	Dim sExt, sPicName
	sExt = UCase(Mid(sFileName, InstrRev(sFileName, ".")+1))
	Select Case sExt
	Case "TXT"
		sPicName = "txt.gif"
	Case "CHM", "HLP"
		sPicName = "hlp.gif"
	Case "DOC"
		sPicName = "doc.gif"
	Case "PDF"
		sPicName = "pdf.gif"
	Case "MDB"
		sPicName = "mdb.gif"
	Case "GIF"
		sPicName = "gif.gif"
	Case "JPG"
		sPicName = "jpg.gif"
	Case "BMP"
		sPicName = "bmp.gif"
	Case "PNG"
		sPicName = "pic.gif"
	Case "ASP", "JSP", "JS", "PHP", "PHP3", "ASPX"
		sPicName = "code.gif"
	Case "HTM", "HTML", "SHTML"
		sPicName = "htm.gif"
	Case "ZIP"
		sPicName = "zip.gif"
	Case "RAR"
		sPicName = "rar.gif"
	Case "EXE"
		sPicName = "exe.gif"
	Case "AVI"
		sPicName = "avi.gif"
	Case "MPG", "MPEG", "ASF"
		sPicName = "mp.gif"
	Case "RA", "RM"
		sPicName = "rm.gif"
	Case "MP3"
		sPicName = "mp3.gif"
	Case "MID", "MIDI"
		sPicName = "mid.gif"
	Case "WAV"
		sPicName = "audio.gif"
	Case "XLS"
		sPicName = "xls.gif"
	Case "PPT", "PPS"
		sPicName = "ppt.gif"
	Case "SWF"
		sPicName = "swf.gif"
	Case Else
		sPicName = "unknow.gif"
	End Select
	FileName2Pic = "<img border=0 src='../sysimage/file/" & sPicName & "'>"
End Function

Function InitSelect(v_InitValue, s_AllName)
	Dim i, aTemp
	InitSelect = ""
	If s_AllName <> "" Then
		InitSelect = InitSelect & "<option value=''>" & s_AllName & "</option>"
	End If
	For i = 1 To Ubound(aStyle)
		aTemp = Split(aStyle(i), "|||")
		InitSelect = InitSelect & "<option value='" & i & "'"
		If CStr(i) = CStr(v_InitValue) Then
			InitSelect = InitSelect & " selected"
		End If
		InitSelect = InitSelect & ">样式：" & inHTML(aTemp(0)) & "---目录：" & inHTML(aTemp(3)) & "</option>"
	Next
End Function

Function InitParam()
	Dim i
	sStyleID = Trim(Request("id"))
	sUploadDir = ""
	If IsNumeric(sStyleID) = True Then
		If Clng(sStyleID) <= Ubound(aStyle) Then
			sUploadDir = Split(aStyle(Clng(sStyleID)), "|||")(3)
		End If
	End If
	If sUploadDir = "" Then
		sStyleID = ""
	Else
		sUploadDir = Replace(sUploadDir, "\", "/")
		If Right(sUploadDir, 1) <> "/" Then
			sUploadDir = sUploadDir & "/"
		End If
		If Left(sUploadDir, 1) <> "/" Then
			sUploadDir = "../" & sUploadDir
		End If
	End If
	sCurrDir = sUploadDir

	sDir = Trim(Request("dir"))
	If sDir <> "" Then
		If CheckValidDir(Server.Mappath(sUploadDir & sDir)) = True Then
			sCurrDir = sUploadDir & sDir & "/"
		Else
			sDir = ""
		End If
	End If
End Function

Function CheckValidDir(s_Dir)
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	CheckValidDir = oFSO.FolderExists(s_Dir)
	Set oFSO = Nothing	
End Function
%>