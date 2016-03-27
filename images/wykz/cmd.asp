<!--#include file="checklogin.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<%
dim act,frompath,page
dim folder_name,file_name,fso,realpath,backurl
act=lcase(trim(request.form("act")))&""
path=trim(Request.Form("path"))&""
page=trim(request.form("page"))&""

if path="" then
	path=trim(Request.QueryString("path"))&""
end if
if act="" then
	act=lcase(trim(request.QueryString("act")))
end if
if page="" then
	page=lcase(trim(request.QueryString("page")))
end if

if right(path,1)<>"/" then
	path=path&"/"
end if
if left(path,1)<>"/" then
	path="/"&path
end if
if InvalidChar(path,0) then
	call showmsg("����ָ����·�����зǷ��ַ���","")
	response.end
end if
Call CheckPath(PathCanModify,path)
realpath=Server.MapPath(path)

select case act
	case "copy","move"
		Call CopyMove()
	case "del","rd","md","copycon","renfile","renfolder"
		Call Commands()
	case "deltree"
		Call Deltree()
	case else
		call showmsg("������ʧ����������ȷ������","back:off")
		response.end
end select

Sub CopyMove()
	CheckSubmit()
	frompath=trim(request.form("frompath"))
	if InvalidChar(frompath,0) then
		call showmsg("����ָ����·�����зǷ��ַ���","")
		response.end
	end if
	Call CheckPath(PathCanModify,frompath)
	session(SessionPrefix&"folders")=""
	session(SessionPrefix&"files")=""
	for each folder_name in request.form("folders")
		session(SessionPrefix&"folders")=session(SessionPrefix&"folders")&"||"&folder_name
	next
	if left(session(SessionPrefix&"folders"),2)="||" then
		session(SessionPrefix&"folders")=right(session(SessionPrefix&"folders"),len(session(SessionPrefix&"folders"))-2)
	end if
	for each file_name in request.form("files")
		session(SessionPrefix&"files")=session(SessionPrefix&"files")&"||"&file_name
	next
	if left(session(SessionPrefix&"files"),2)="||" then
		session(SessionPrefix&"files")=right(session(SessionPrefix&"files"),len(session(SessionPrefix&"files"))-2)
	end if
	if session(SessionPrefix&"folders")="" and session(SessionPrefix&"files")="" then
		session(SessionPrefix&"act")=""
		session(SessionPrefix&"frompath")=""
		call showmsg("����û��ѡ���κ��ļ����ļ��У�","")
	else
		session(SessionPrefix&"act")=act
		session(SessionPrefix&"frompath")=server.mappath(frompath)
		response.redirect "explorer.asp?path="&Server.URLEncode(path)&"&page="&page
	end if
End Sub

Sub Commands()
		frompath=Server.URLEncode(trim(request("currentpath")))
		set fso=server.createobject("scripting.filesystemobject")
		backurl="explorer.asp?path="&frompath&"&page="&page
		if act="del" then
			if fso.FileExists(realpath) then
				fso.DeleteFile(realpath)
				if fso.FileExists(realpath) then
					set fso=nothing
					call showmsg("����ɾ���ļ�ʧ�ܣ���ȷ��������Ӧ��Ȩ�ޡ�","back:off")
					response.End
				end if
				set fso=nothing
				'response.Redirect(backurl)
			else
				set fso=nothing
				call showmsg("������Ҫɾ�����ļ� "&getselfname(realpath,"\")&" �����ڣ�","back:off")
				response.end
			end if
		elseif act="rd" then
			if trim(path)="/" then
				call showmsg("�����޷�ɾ����վ��Ŀ¼��","back:off")
				response.End
			end if
			if fso.FolderExists(realpath) then
				fso.DeleteFolder(realpath)
				if fso.FolderExists(realpath) then
					set fso=nothing
					call showmsg("����ɾ���ļ���ʧ�ܣ���ȷ��������Ӧ��Ȩ�ޡ�","back:off")
					response.End
				end if
				set fso=nothing
				'response.Redirect(backurl)
			else
				set fso=nothing
				call showmsg("������Ҫɾ�����ļ��� "&getselfname(realpath,"\")&" �����ڣ�","back:off")
				response.end
			end if
		elseif act="md" then
			if fso.FolderExists(realpath) then
				set fso=nothing
				call showmsg("������Ҫ�½���Ŀ¼ "&getselfname(realpath,"\")&" �Ѿ����ڣ�","back:off")
				response.end
			end if
			fso.CreateFolder(realpath)
			if not fso.FolderExists(realpath) then
				set fso=nothing
				call showmsg("�����½�Ŀ¼ʧ�ܣ���ȷ��������Ӧ��Ȩ�ޡ�","back:off")
				response.end
			end if
			set fso=nothing
			'response.Redirect(backurl)
		elseif act="copycon" then
			if fso.FileExists(realpath) then
				set fso=nothing
				call showmsg("������Ҫ�½����ļ� "&getselfname(realpath,"\")&" �Ѿ����ڣ�","back:off")
				response.end
			end if
			fso.CreateTextFile(realpath)
			if not fso.FileExists(realpath) then
				set fso=nothing
				call showmsg("�����½��ļ�ʧ�ܣ���ȷ��������Ӧ��Ȩ�ޡ�","back:off")
				response.end
			end if
			set fso=nothing
			'response.Redirect(backurl)
		elseif act="renfile" or act="renfolder" then
			dim oldname,newname,addchar
			oldname=trim(request("oldname"))
			newname=trim(request("newname"))
			if act="renfolder" then
				addchar="��"
			end if
			if oldname="" or newname="" then
				set fso=nothing
				call showmsg("���󣺾��ļ�"&addchar&"�������ļ�"&addchar&"�������벻��Ϊ�գ�","back:off")
				response.end
			elseif InvalidChar(oldname,1) then
				set fso=nothing
				call showmsg("���󣺾��ļ�"&addchar&"�����Ƿ��ַ���","back:off")
				response.end
			elseif InvalidChar(newname,1) then
				set fso=nothing
				call showmsg("�������ļ�"&addchar&"�����Ƿ��ַ���","back:off")
				response.end
			end if
			oldname=Server.MapPath(path&oldname)
			newname=Server.MapPath(path&newname)
			if act="renfile" then
				if fso.FileExists(oldname) then
					if not fso.FileExists(newname) then 
						fso.GetFile(oldname).name=trim(request("newname"))
						if not fso.FileExists(newname) then
							set fso=nothing
							call showmsg("���󣺸����ļ�"&addchar&"��ʧ�ܣ���ȷ��������Ӧ��Ȩ�ޡ�","back:off")
							response.end
						end if
						set fso=nothing
						'response.Redirect(backurl)
					else
						set fso=nothing
						call showmsg("ʧ�ܣ����ļ�"&addchar&"�� "&trim(request("newname"))&" ���ļ�"&addchar&"�Ѿ����ڣ�","back:off")
						response.end
					end if
				else
					set fso=nothing
					call showmsg("ʧ�ܣ�Ҫ���ĵ��ļ�"&addchar&" "&trim(request("oldname"))&" �����ڣ�","back:off")
					response.end
				end if
			elseif act="renfolder" then
				if fso.FolderExists(oldname) then
					if not fso.FolderExists(newname) then 
						fso.GetFolder(oldname).name=trim(request("newname"))
						if not fso.FolderExists(newname) then
							set fso=nothing
							call showmsg("���󣺸����ļ�"&addchar&"��ʧ�ܣ���ȷ��������Ӧ��Ȩ�ޡ�","back:off")
							response.end
						end if
						set fso=nothing
						'response.Redirect(backurl)
					else
						set fso=nothing
						call showmsg("ʧ�ܣ����ļ�"&addchar&"�� "&trim(request("newname"))&" ���ļ�"&addchar&"�Ѿ����ڣ�","back:off")
						response.end
					end if
				else
					set fso=nothing
					call showmsg("ʧ�ܣ�Ҫ���ĵ��ļ�"&addchar&" "&trim(request("oldname"))&" �����ڣ�","back:off")
					response.end
				end if
			end if
		end if
End Sub

Sub Deltree()
		CheckSubmit()
		dim folders,files,foldername,filename,tempname
		frompath=trim(request.form("frompath"))
		if InvalidChar(frompath,0) then
			call showmsg("����ָ����·�����зǷ��ַ���","")
			response.end
		end if
		Call CheckPath(PathCanModify,frompath)
		frompath=Server.MapPath(frompath)
		backurl="explorer.asp?path="&Server.URLEncode(path)&"&page="&page
		
		if lcase(trim(request.form("act")))="deltree" then
			set folders=request.form("folders")
			set files=request.form("files")
			set fso=server.createobject("scripting.filesystemobject")
			if folders.count>0 then
				for each foldername in folders
					tempname=frompath&"\"&foldername
					if fso.FolderExists(tempname) then
						fso.DeleteFolder tempname,true
					end if
				next
			end if
			if files.count>0 then
				for each filename in files
					tempname=frompath&"\"&filename
					if fso.FileExists(tempname) then
						fso.DeleteFile tempname,true
					end if
				next
			end if
			set folders=nothing
			set files=nothing
			set fso=nothing
		end if
		'response.Redirect(backurl)
End Sub

if trim(backurl)&""<>"" then
	backurl=replace(backurl,"\","\\")
	backurl=replace(backurl,chr(34),"\"&chr(34))
end if
%>
<script language="javascript">
alert("ִ�в���������");
try{
	parent.location.reload();
}catch(e){};
this.location.href="about:blank";
</script>
</head>
<body></body>
</html>