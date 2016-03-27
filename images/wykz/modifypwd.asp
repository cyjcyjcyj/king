<!--#include file="checklogin.asp"-->
<%
set conn=server.createobject("adodb.connection")
call conndate()

dim username,password,password2,oldpassword
username=session(SessionPrefix&"logonusername")
chksystem()
if request.form("edituser")="yes" then
	password=trim(request.form("password"))
	password2=trim(request.form("password2"))
	oldpassword=trim(request.Form("oldpassword"))

	if username="" or instr(username,"'")>0 or instr(username,chr(34))>0 then
		call closeconn()
		call showmsg("当前用户名非法变更！","")
		response.end
	end if

	if password="" or oldpassword="" then
		call closeconn()
		call showmsg("请输入旧密码和新密码！","")
		response.end
	end if
	if password<>password2 then
		call closeconn()
		call showmsg("你两次输入的新密码不一致，请重新输入！","")
		response.end
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin_user where username='"&username&"' and password='"&md5(oldpassword)&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		rs.close
		set rs=nothing
		call closeconn()
		call showmsg("抱歉，用户“"&username&"”不存在或者旧密码不正确！","")
		response.end
	else
		if rs("usergroup")=3 and not popedom then
			rs.close
			set rs=nothing
			call closeconn()
			call showmsg("抱歉，用户“"&username&"”是管理员，你无权修改！","")
			response.end
		end if
		rs("password")=md5(password)
		rs.update
		rs.close
		set rs=nothing
		call closeconn()
		call showmsg("密码修改成功，请记住新密码。",selfname)
	end if
	response.end
end if
call closeconn()
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - 修改用户密码</title>
<link rel="stylesheet" type="text/css" href="images/admin.css" />
<script type="text/javascript" src="images/md5.js"></script>
<script type="text/javascript" src="images/admin.js"></script>
</head>
<body>
<div id="bodyposition">
	<div id="content">
	  <div id="title"><%=SystemName%> - 修改用户密码</div>
	  <div id="main" style="border-top:1px #000099 solid;">
		<table width="100%" border="0" cellspacing="5" cellpadding="0">
		 <form method="post" action="<%=selfname%>?username=<%=username%>" onsubmit="return formcheck(this,1)">
		 <input type="hidden" name="edituser" value="yes">
		  <tr><td colspan="2" height="5"></td></tr>
		  <tr>
			<td align="right" width="33%">旧 密 码：</td>
			<td align="left" width="67%"><input type="password" name="oldpassword" class="txtinput">&nbsp;请输入旧密码</td>
		  </tr>
		  <tr>
			<td align="right">新 密 码：</td>
			<td align="left"><input type="password" name="password" class="txtinput">&nbsp;请输入新密码</td>
		  </tr>
		  <tr>
			<td align="right">确认密码：</td>
			<td align="left"><input type="password" name="password2" class="txtinput">&nbsp;确认新密码</td>
		  </tr>
		  <tr>
			<td colspan="2" height="40">
			<input type="submit" value="确定" class="button">&nbsp;&nbsp;<input type="reset" value="重置" class="button">&nbsp;&nbsp;<input type="button" value="关闭" class="button" onClick="window.close()">
			</td>
		  </tr>
		 </form>
		</table>
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