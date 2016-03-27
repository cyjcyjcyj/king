<!--#include file="checklogin.asp"-->
<%
if not popedom then
	call showmsg("你没有权限进入此页面！","index.asp")
	response.end
end if

dim username,password,password2,usergroup,adminpath,oldpassword,oldusername
username=trim(request.querystring("username"))
chksystem()
if username<>"" then
	username=lcase(replace(username,"'",""))
else
	call showmsg("没有指定要修改的用户！","")
	response.end
end if

set conn=server.createobject("adodb.connection")
call conndate()
set rs=server.createobject("adodb.recordset")

if request.form("edituser")="yes" then
	password=trim(request.form("password"))
	password2=trim(request.form("password2"))
	oldusername=lcase(trim(request.Form("oldusername")))
	username=lcase(trim(request.form("username")))
	usergroup=trim(request.form("usergroup"))
	adminpath=trim(request.form("adminpath"))
	oldpassword=trim(request.Form("oldpassword"))
	
	if usergroup<>"" and isnumeric(usergroup) then
		if cstr(usergroup)<>"1" and cstr(usergroup)<>"3" then
			usergroup=1
		end if
	else
		usergroup=1
	end if

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
	if oldusername="" or instr(oldusername,"'")>0 or instr(oldusername,chr(34))>0 then
		call closeconn()
		call showmsg("旧用户名非法变更！","")
		response.end
	end if

	if password<>"" or password2<>"" then
		if password<>password2 then
			call closeconn()
			call showmsg("你两次输入的新密码不一致，请重新输入！","")
			response.end
		end if
	end if

	
	sql="select * from admin_user where username='"&oldusername&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		rs.close
		set rs=nothing
		call closeconn()
		call showmsg("抱歉，用户“"&oldusername&"”不存在！","")
		response.end
	else
		if rs("usergroup")=3 then
			if rs("password")<>md5(oldpassword) then
				rs.close
				set rs=nothing
				call closeconn()
				call showmsg("抱歉，用户“"&oldusername&"”是管理员，必须输入正确的密码才可以修改！","")
				response.end
			end if
		end if
		if password<>"" then
			rs("password")=md5(password)
		end if
		if oldusername<>username then
			rs("username")=username
		end if
		rs("adminpath")=adminpath
		rs("usergroup")=usergroup
		rs.update
		rs.close
		set rs=nothing
		call closeconn()
		call showmsg("","adminuser.asp?editinfo="&server.URLEncode("修改用户“" & oldusername & "”成功！"))
	end if
	response.end
end if

dim uusergroup,upath
sql="select * from admin_user where username='"&username&"'"
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	rs.close
	set rs=nothing
	call closeconn()
	call showmsg("抱歉，用户“"&username&"”不存在！","")
	response.end
else
	uusergroup=rs("usergroup")
	upath=rs("adminpath")
end if
rs.close
set rs=nothing
call closeconn()
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - 编辑用户</title>
<link rel="stylesheet" type="text/css" href="images/admin.css" />
<script type="text/javascript" src="images/md5.js"></script>
<script type="text/javascript" src="images/admin.js"></script>
</head>
<body>
<div id="bodyposition">
	<div id="content">
	  <div id="title"><%=SystemName%> - 编辑用户</div>
	  <div id="main" style="border-top:1px #000099 solid;">
		<table width="100%" border="0" cellspacing="5" cellpadding="0">
		 <form name="edit_user" method="post" action="<%=selfname%>?username=<%=username%>" onSubmit="return formcheck(this,2)">
		 <input type="hidden" name="edituser" value="yes">
		  <tr><td colspan="2" height="5"></td></tr>
		  <tr>
			<td align="right" width="33%">用 户 名：</td>
			<td align="left" width="67%"><input type="text" name="username" value="<%=username%>" class="txtinput">&nbsp;无必要请勿随便更改用户名
			<input type="hidden" name="oldusername" value="<%=username%>">
			</td>
		  </tr>
<%if uusergroup=3 then%>
		  <tr>
			<td align="right">旧 密 码：</td>
			<td align="left"><input type="password" name="oldpassword" class="txtinput">&nbsp;<font color="#FF0000">修改管理员需要填写旧密码</font></td>
		  </tr>
<%end if%>
		  <tr>
			<td align="right">新 密 码：</td>
			<td align="left"><input type="password" name="password" class="txtinput">&nbsp;不更改密码请留空</td>
		  </tr>
		  <tr>
			<td align="right">确认密码：</td>
			<td align="left"><input type="password" name="password2" class="txtinput">&nbsp;不更改密码请留空</td>
		  </tr>
		  <tr>
			<td align="right">用户级别：</td>
			<td align="left">
				<select name="usergroup" onChange="javascript:if(this.form.usergroup.options[this.form.usergroup.options.selectedIndex].value==3) this.form.adminpath.value='/';">
				<option value="1"<%if uusergroup<>3 then response.write " selected"%>>普通用户</option>
				<option value="3"<%if uusergroup=3 then response.write " selected"%>>管理员</option>
				</select>
			</td>
		  </tr>
		  <tr>
			<td align="right">管理目录：</td>
			<td align="left"><input type="text" name="adminpath" value="<%=upath%>" class="txtinput">&nbsp;<font color="#FF0000">不要使用类似“<font color="#0000FF">../</font>”的路径</font></td>
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