<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="conn.asp"-->
<%
call closeconn()
dim path,loginstatus,comeurl,vars,PathCanModify
loginstatus=false
path=trim(Request.QueryString("path"))&""
PathCanModify=trim(session(SessionPrefix&"adminpath"))&""
chksystem()

if session(SessionPrefix&"loginstatus")<>md5s(userip) then
	response.redirect "index.asp?error=99&url="&server.urlencode(Trim(Request.ServerVariables("HTTP_REFERER")))
	response.end
else
	if LoginOneIP=1 then
		if instr(Application(OnlineValue),"|" & md5s(session(SessionPrefix&"logonusername")) & ";" & md5s(userip) & "|")<=0 then
			response.redirect "index.asp?error=99&url="&server.urlencode(Trim(Request.ServerVariables("HTTP_REFERER")))
			response.end
		end if
	end if
	loginstatus=true
end if

if path="" then
	path=PathCanModify
end if
if right(path,1)<>"/" then
	path=path&"/"
end if
if left(path,1)<>"/" then
	path="/"&path
end if
if InvalidChar(path,0) then
	call showmsg("错误：指定的路径含有非法字符！","")
	response.end
end if
Call CheckPath(PathCanModify,path)

Sub CheckPath(manage_path,current_path)
	dim manage_path2,current_path2,chkfirst
	manage_path2=lcase(Server.MapPath(manage_path))
	current_path2=lcase(Server.MapPath(current_path))
	if not popedom and (manage_path2<>left(current_path2,len(manage_path2)) or md5s(lcase(manage_path))<>session(SessionPrefix&"syscode")) then
		call showmsg("错误：你只有管理目录 "&manage_path&" 及其子目录的权限！","")
		response.end
	end if
End Sub

function popedom()
	popedom=false
	if trim(md5s("3" & session(SessionPrefix&"logonusername")))=trim(session(SessionPrefix&"usergroup")) then
		popedom=true
	end if
end function
%>
