<!--#include file="checklogin.asp"-->
<%
if not popedom then
	call showmsg("��û��Ȩ�޽����ҳ�棡","index.asp")
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
		call showmsg("Ŀ¼·�����ú������·Ƿ��ַ���\n\\ : * ? < > | \"& chr(34),"")
		response.end
	end if
	if username="" or instr(username,"'")>0 or instr(username,chr(34))>0 then
		call closeconn()
		call showmsg("�û�������Ϊ�ն��Ҳ������ţ���������д��","")
		response.end
	end if
	if password="" then
		call closeconn()
		call showmsg("���벻��Ϊ�գ�","")
		response.end
	end if
	if password<>password2 then
		call closeconn()
		call showmsg("������������û����벻һ�£����������룡","")
		response.end
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin_user where username='"&username&"'"
	rs.open sql,conn,1,3
	if not (rs.bof and rs.eof) then
		rs.close
		set rs=nothing
		call closeconn()
		call showmsg("�Բ����û���"&username&"���Ѿ����ڣ���ʹ�������û�����","")
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
		call showmsg("","adminuser.asp?editinfo="&server.URLEncode("�����û���" & username & "���ɹ���"))
	end if
	response.end
end if
call closeconn()
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - �����û�</title>
<link rel="stylesheet" type="text/css" href="images/admin.css" />
<script type="text/javascript" src="images/md5.js"></script>
<script type="text/javascript" src="images/admin.js"></script>
</head>
<body>
<div id="bodyposition">
	<div id="content">
	  <div id="title"><%=SystemName%> - �����û�</div>
	  <div id="main" style="border-top:1px #000099 solid;">
		<table width="100%" border="0" cellspacing="5" cellpadding="0">
		 <form method="post" action="<%=selfname%>" onSubmit="return formcheck(this,1)">
		 <input type="hidden" name="adduser" value="yes">
		  <tr><td colspan="2" height="5"></td></tr>
		  <tr>
			<td align="right" width="33%">�� �� ����</td>
			<td align="left" width="67%"><input type="text" name="username" value="<%=username%>" class="txtinput"></td>
		  </tr>
		  <tr>
			<td align="right">��&nbsp;&nbsp;&nbsp; �룺</td>
			<td align="left"><input type="password" name="password" class="txtinput"></td>
		  </tr>
		  <tr>
			<td align="right">ȷ�����룺</td>
			<td align="left"><input type="password" name="password2" class="txtinput"></td>
		  </tr>
		  <tr>
			<td align="right">�û�����</td>
			<td align="left">
				<select name="usergroup" onChange="javascript:if(this.form.usergroup.options[this.form.usergroup.options.selectedIndex].value==3) this.form.adminpath.value='/';">
				<option value="1" selected>��ͨ�û�</option>
				<option value="3">����Ա</option>
				</select>
			</td>
		  </tr>
		  <tr>
			<td align="right">����Ŀ¼��</td>
			<td align="left"><input type="text" name="adminpath" class="txtinput">&nbsp;<font color="#FF0000">��Ҫʹ�����ơ�<font color="#0000FF">../</font>����·��</font></td>
		  </tr>
		  <tr>
			<td colspan="2" height="40">
			<input type="submit" value="ȷ��" class="button">&nbsp;&nbsp;<input type="reset" value="����" class="button">&nbsp;&nbsp;<input type="button" value="����" class="button" onClick="window.location.href='adminuser.asp'">
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