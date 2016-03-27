<!--#include file="checklogin.asp"-->
<%
if not popedom then
	call showmsg("你没有权限进入此页面！","index.asp")
	response.end
end if

set conn=server.createobject("adodb.connection")
call conndate()

dim username,password,password2,usergroup,adminpath
username=lcase(trim(request.form("username")))
chksystem()
if request.form("adduser")="yes" then
	password=trim(request.form("password"))
	password2=trim(request.form("password2"))
	usergroup=trim(request.form("usergroup"))
	if usergroup<>"" and isnumeric(usergroup) then
		if cstr(usergroup)<>"1" and cstr(usergroup)<>"3" then
			usergroup=1
		end if
	else
		usergroup=1
	end if
	adminpath=trim(request.form("adminpath"))
	if left(adminpath,1)<>"/" then
		adminpath="/"
	end if
	if right(adminpath,1)<>"/" then
		adminpath=adminpath&"/"
	end if
	if InvalidChar(adminpath,0) then
		call closeconn()
		call showmsg("目录路径不得含有以下非法字符：\n\\ : * ? < > | \"& chr(34),"")
		response.end
	end if
	if username="" or instr(username,"'")>0 or instr(username,chr(34))>0 then
		call closeconn()
		call showmsg("用户名不能为空而且不含引号，请重新填写！","")
		response.end
	end if
	if password="" then
		call closeconn()
		call showmsg("密码不能为空！","")
		response.end
	end if
	if password<>password2 then
		call closeconn()
		call showmsg("你两次输入的用户密码不一致，请重新输入！","")
		response.end
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin_user where username='"&username&"'"
	rs.open sql,conn,1,3
	if not (rs.bof and rs.eof) then
		rs.close
		set rs=nothing
		call closeconn()
		call showmsg("对不起，用户“"&username&"”已经存在，请使用其它用户名！","")
		response.end
	else
		rs.addnew
		rs("username")=username
		rs("password")=md5(password)
		rs("usergroup")=usergroup
		rs("adminpath")=adminpath
		rs.update
		rs.close
		set rs=nothing
		call closeconn()
		call showmsg("","adminuser.asp?editinfo="&server.URLEncode("添加用户“" & username & "”成功！"))
	end if
	response.end
end if
call closeconn()
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - 添加用户</title>
<link rel="stylesheet" type="text/css" href="images/admin.css" />
<script type="text/javascript" src="images/md5.js"></script>
<script type="text/javascript" src="images/admin.js"></script>
</head>
<body>
<div id="bodyposition">
	<div id="content">
	  <div id="title"><%=SystemName%> - 添加用户</div>
	  <div id="main" style="border-top:1px #000099 solid;">
		<table width="100%" border="0" cellspacing="5" cellpadding="0">
		 <form method="post" action="<%=selfname%>" onSubmit="return formcheck(this,1)">
		 <input type="hidden" name="adduser" value="yes">
		  <tr><td colspan="2" height="5"></td></tr>
		  <tr>
			<td align="right" width="33%">用 户 名：</td>
			<td align="left" width="67%"><input type="text" name="username" value="<%=username%>" class="txtinput"></td>
		  </tr>
		  <tr>
			<td align="right">密&nbsp;&nbsp;&nbsp; 码：</td>
			<td align="left"><input type="password" name="password" class="txtinput"></td>
		  </tr>
		  <tr>
			<td align="right">确认密码：</td>
			<td align="left"><input type="password" name="password2" class="txtinput"></td>
		  </tr>
		  <tr>
			<td align="right">用户级别：</td>
			<td align="left">
				<select name="usergroup" onChange="javascript:if(this.form.usergroup.options[this.form.usergroup.options.selectedIndex].value==3) this.form.adminpath.value='/';">
				<option value="1" selected>普通用户</option>
				<option value="3">管理员</option>
				</select>
			</td>
		  </tr>
		  <tr>
			<td align="right">管理目录：</td>
			<td align="left"><input type="text" name="adminpath" class="txtinput">&nbsp;<font color="#FF0000">不要使用类似“<font color="#0000FF">../</font>”的路径</font></td>
		  </tr>
		  <tr>
			<td colspan="2" height="40">
			<input type="submit" value="确定" class="button">&nbsp;&nbsp;<input type="reset" value="重置" class="button">&nbsp;&nbsp;<input type="button" value="返回" class="button" onClick="window.location.href='adminuser.asp'">
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