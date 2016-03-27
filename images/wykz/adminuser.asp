<!--#include file="checklogin.asp"-->
<%
if not popedom then
	call showmsg("你没有权限进入此页面！","index.asp")
	response.end
end if

set conn=server.createobject("adodb.connection")
call conndate()
dim username,act,isdel

username=trim(request.querystring("username"))
act=lcase(trim(request.querystring("act")))
isdel=lcase(trim(request.querystring("isdel")))
chksystem()
if act="deluser" and username<>"" and isdel="yes" then
	username=lcase(replace(username,"'",""))
	if session(SessionPrefix&"logonusername")=username then
		call closeconn()
		call showmsg("",selfname&"?editinfo=" & Server.URLEncode("不能删除自己！"))
		response.end
	end if
	sql="delete from admin_user where username='"&username&"'"
	conn.execute(sql)
	call closeconn()
	call showmsg("",selfname&"?editinfo=" & Server.URLEncode("删除用户“"&username&"”成功！"))
	response.end
end if
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - 用户管理</title>
<link rel="stylesheet" type="text/css" href="images/admin.css" />
<script language="javascript">
<!--
function deluser(username)
{
	var yourname="<%=session(SessionPrefix&"logonusername")%>";
	if (username!=yourname)
	{
		if (window.confirm("你真的要删除用户“"+username+"”吗？"))
			window.location.href="<%=selfname%>?act=deluser&username="+username+"&isdel=yes";
	}else{
		alert("晕，杀自己不要命了吗！？");
	}
}
//-->
</script>
</head>
<body>
<div id="bodyposition">
	<div id="titler"><a href="adduser.asp">添加用户</a></div>
	<div id="content">
	  <div id="title"><%=SystemName%> - 用户管理</div>
	  <div id="main">
		<div class="peruser">
			<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr align="center" style="color:#990000; font-weight:bold;">
				<td width="16%" class="puborder">用户名</td>
				<td width="15%" class="puborder">管理级别</td>
				<td width="54%" class="puborder">管理目录</td>
				<td width="15%">操作</td>
			  </tr>
			</table>
		</div>
<%
dim tempname,uname,uusergroup
sql="select * from admin_user order by username asc"
set rs=Server.CreateObject("ADODB.RecordSet")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.Write "<div class='peruser'><font color='#ff0000'><b>还没有任何用户！请添加。</b></font></div>"
else
	do while not rs.eof
		tempname=lcase(rs("username"))
		uname=htmlcode(tempname)
		if tempname=session(SessionPrefix&"logonusername") then
			uname="<b>"&uname&"</b>"
		end if
		if rs("usergroup")=3 then
			uusergroup="<font color='#ff0000'>管理员</font>"
		else
			uusergroup="普通用户"
		end if
%>
		<div class="peruser">
			<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td width="16%" class="puborder"><%=uname%></td>
				<td width="15%" class="puborder"><%=uusergroup%></td>
				<td width="54%" class="puborder" align="left">&nbsp;
				<a target="_blank" href="explorer.asp?path=<%=Server.URLEncode(rs("adminpath"))%>"><%=htmlcode(Server.MapPath(rs("adminpath")))%></a>
				</td>
				<td width="15%" align="center"><a href="javascript:deluser('<%=tempname%>')">删除</a>&nbsp;&nbsp;&nbsp;<a href="edituser.asp?username=<%=tempname%>">编辑</a></td>
			  </tr>
			</table>
		</div>
<%
		rs.movenext
	loop
end if
rs.close
set rs=nothing
call closeconn()
if trim(request("editinfo"))<>"" then
	response.Write "<div class='peruser'><span style='position:relative; top:6px; color:#ff0000;'>"&trim(request("editinfo"))&"</span></div>"
end if
%>
	  </div>
	  <div id="footer"><br><%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
%><span style='position:relative;top:4px; text-align:center;line-height:120%;'>
<%=copyright%><br>Processed in <%=(endtime-starttime)*1000%> MSEL
</span><br><br></div>
	</div>
</div>
</body>
</html>