<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<%
if session(SessionPrefix&"loginstatus")=md5s(userip) and left(session(SessionPrefix&"adminpath"),1)="/" then
	if LoginOneIP=1 then
		if instr(Application(OnlineValue),"|" & md5s(session(SessionPrefix&"logonusername")) & ";" & md5s(userip) & "|")>0 then
			response.redirect "explorer.asp?path=" & Server.URLEncode(session(SessionPrefix&"adminpath"))
			response.end
		end if
	else
		response.redirect "explorer.asp?path=" & Server.URLEncode(session(SessionPrefix&"adminpath"))
		response.end
	end if
end if

dim errinfo,comeurl
If Not IsObjInstalled("Scripting.FileSystemObject") Then
	errinfo="�÷�������֧��FSO�����޷�ʹ�ñ�����<br>"
end if
chksystem()
'if not IsObjInstalled("adodb.stream") Then
'	errinfo=errinfo & "������ADO����汾̫�ͣ��㽫�޷�ʹ���ϴ����ܣ�<br>"
'end if

comeurl=trim(request.querystring("url"))
if comeurl<>"" then
	comeurl=replace(comeurl,chr(34),"")
end if
if lcase(trim(request.form("act")))="submit" then
%>
<!--#include file="conn.asp"-->
<%
	dim server_v1,server_v2,lastlogin,adminpath,adminusergroup
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		call closeconn()
		Response.redirect selfname&"?error=8"
		response.end
	end if
	
	dim logonusername,loginpassword
	logonusername=trim(request.form("TPL_username"))
	loginpassword=request.form("TPL_password")
	
	if trim(session(SessionPrefix&"syscode"))<>trim(request.form("CheckCode")) or trim(session(SessionPrefix&"syscode"))="" Then
		call closeconn()
		Response.redirect selfname&"?user="&logonusername&"&error=3"
		Response.End
	end if
	session(SessionPrefix&"syscode")=""
	
	if logonusername="" or loginpassword="" then
		response.redirect selfname&"?error=23"
		response.end
	end if
	logonusername=replace(logonusername,chr(34),"")
	logonusername=replace(logonusername,"'","")
	loginpassword=md5(loginpassword)
	
	comeurl=trim(request.form("comeurl"))
	sql="select * from admin_user where username='"&logonusername&"' and password='"&loginpassword&"'"
	Set rs=Server.CreateObject("ADODB.RecordSet")
	rs.open sql,conn,1,1

	if rs.bof and rs.eof then
		rs.close:set rs=nothing
		call closeconn()
		Response.redirect selfname&"?user="&logonusername&"&error=4"
		Response.End
	end if
	lastlogin=trim(rs("last"))
	adminpath=trim(rs("adminpath"))
	adminusergroup=rs("usergroup")
	rs.close

	if LoginOneIP=1 then
		if Application(OnlineValue)=Empty or Application(OnlineValue)="" then
			Application.Lock
			Application(OnlineValue)="|" & md5s(lcase(logonusername)) & ";" & md5s(userip) & "|"
			Application.unLock
		elseif instr(Application(OnlineValue),"|" & md5s(lcase(logonusername)) & ";")>0 then
			if instr(Application(OnlineValue),";" & md5s(userip) & "|")<=0 then
				set rs=nothing
				call closeconn()
				response.redirect selfname&"?user="&logonusername&"&error=44"
				response.end
			end if
		else
			Application.Lock
			Application(OnlineValue)=Application(OnlineValue) & md5s(lcase(logonusername)) & ";" & md5s(userip) & "|"
			Application.unLock
		end if
	end if

	if len(lastlogin)>0 then
		session(SessionPrefix&"lastdata")=cstr(lastlogin)
	else
		session(SessionPrefix&"lastdata")="first"
	end if
	session(SessionPrefix&"syscode")=md5s(lcase(adminpath))
	session(SessionPrefix&"logonusername")=lcase(logonusername)
	session(SessionPrefix&"usergroup")=md5s(adminusergroup & lcase(logonusername))
	session(SessionPrefix&"adminpath")=adminpath
	session(SessionPrefix&"loginstatus")=md5s(userip)
	session.timeout=20

	rs.open sql,conn,2,3
	rs("last")=userip & "|" & cstr(now())
	rs.update
	rs.close:set rs=nothing
	call closeconn()
	if comeurl="" then
		response.redirect "explorer.asp?path="&Server.URLEncode(adminpath)
	else
		response.redirect comeurl
	end if
end if
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - �����¼</title>
<link rel="stylesheet" type="text/css" href="images/login.css" />
<script type="text/javascript" src="images/md5.js"></script>
<script type="text/javascript" src="images/login.js"></script>
</head>
<body onresize="body_center()">
<div id="_body" style="position:relative">
<div id="Content" align="center">
	<div id="LoginForm" class="StandardMode">
		<div id="LoginMSG">
<%
	if errinfo<>"" then response.write errinfo
	select case trim(request.querystring("error"))
	case "23"
		response.write "�û��������벻��Ϊ�գ�"
	case "3"
		response.write "�������֤�벻��ȷ��ʱʧЧ��"
	case "4"
		response.write "�û��������벻��ȷ��"
	case "8"
		response.write "��ֹ��վ���ύ���ݣ���ķǷ������ѱ���¼��"
	case "99"
		response.write "�㻹û�е�½�����ߵ�½�ѳ�ʱ��"
	case "44"
		response.write "�û���" & request.querystring("user") & "��������һ��IP��½��"
	end select
%>
		</div>
		<div id="LoginFormTop"></div>
		<form name="user_login" action="<%=selfname%>" onSubmit="return checklogin()" method="post">					
			<label>�û�����<input name="TPL_username" id="TPL_username" type="text" size="32" value="<%=trim(request.querystring("user"))%>" maxlength="32" tabindex="1"/></label>
			<label>�ܡ��룺<input name="TPL_password" id="TPL_password" type="password" maxlength="35" size="32" tabindex="2"/></label>
			<label>��֤�룺<input name="CheckCode" id="CheckCode" type="text" maxlength="10" size="10" tabindex="3" title="����ı��������֤��" autocomplete="off" onFocus="setcode()" /></label>
			<div id="CheckCodeMsg" style="display:none"></div>
			<div id="CheckCodeImg" style="display:none"></div>
			<div class="Submit">
			<input type="submit" id="SubmitButton" value="�� ¼" tabindex="4" />&nbsp;<input type="reset" value="�� ��" tabindex="5" />
			</div>
			<input type="hidden" name="comeurl" value="<%=comeurl%>">
			<input type="hidden" value="submit" name="act">
		</form>
		<div id="LoginFormBottom"></div>
		<div id="FooterMSG"><%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
%><span style='position:relative;top:4px; text-align:center;line-height:120%;'>
<%=copyright%><br>Processed in <%=(endtime-starttime)*1000%> MSEL
</span>
		</div>
	</div>
</div>
</div>
<script language="javascript">InitInput();<%if ShowValidate=1 then response.Write("setcode();")%>body_center();</script>
</body>
</html>
