<!--#include file="checklogin.asp"-->
<!--#include file="incupload.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

if not popedom then
	call showmsg("��û��Ȩ�޽����ҳ�棡","index.asp")
	response.end
end if

if lcase(trim(request("act")))="chkobj" then
	call chkupobj()
	response.end
end if

dim filecontent,obj_fso,filetemp,msg
if trim(request.Form("save"))="yes" then
	call saveconfig()
end if
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - ϵͳ����</title>
<link rel="stylesheet" type="text/css" href="images/config.css" />
</head>
<body>
<div id="bodyposition">
	<div id="content">
	  <div id="title"><%=SystemName%> - ϵͳ����</div>
	  <%if msg<>"" then response.write "<div id='msg' style='border-top:1px #000099 solid;'>"&msg&"</div>"%>
	  <div id="main" style="border-top:1px #000099 solid;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		 <form method="post" action="<%=selfname%>">
		 <input type="hidden" name="save" value="yes">
		  <tr><td colspan="2" height="5"></td></tr>
		  <tr height="30">
			<td align="right" width="22%">�ϴ��ű���ʱʱ�䣺&nbsp;</td>
			<td align="left" width="78%"><input type="text" name="Script_Timeout" value="<%=Script_Timeout%>" class="txtinput">&nbsp;��&nbsp;<font color="#999999">���鲻Ҫ���õ�̫��</font></td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">�ϴ�������ͣ�&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td>
						<select id="UploadObject" name="UploadObject" size="1" onChange="document.getElementById('chkobjpage').src='<%=selfname%>?act=chkobj&upobj='+document.getElementById('UploadObject').value">
						<option value="0"<%if UploadObject=0 then response.write " selected"%>>������ϴ���</option>
						<option value="1"<%if UploadObject=1 then response.write " selected"%>>AspUpLoad</option>
						<option value="2"<%if UploadObject=2 then response.write " selected"%>>SA-FileUp</option>
						<option value="3"<%if UploadObject=3 then response.write " selected"%>>LyfUpload</option>
						</select>
					</td>
					<td><iframe id="chkobjpage" src="about:blank" frameborder=0 scrolling=no width="300" height="15"></iframe></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">�ϴ��ļ���С��&nbsp;</td>
			<td align="left" class="tdborder"><input type="text" name="FileMaxSize" value="<%=FileMaxSize%>" class="txtinput">&nbsp;�ֽ�&nbsp;<font color="#999999">����������Ϊ0������Ա������</font></td>
		  </tr>
		  <tr height="85">
			<td align="right" class="tdborder">��ֹ�ϴ����ļ����ͣ�&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><textarea name="NotAllowFileType" cols="50" rows="5"><%=NotAllowFileType%></textarea></td>
					<td width="100%">&nbsp;<font color="#999999">�����������գ��á�|������������Է�LyfUpload�����Ч������Ա������</font></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr height="85">
			<td align="right" class="tdborder">�����ϴ����ļ����ͣ�&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><textarea name="AllowFileType" cols="50" rows="5"><%=AllowFileType%></textarea></td>
					<td width="100%">&nbsp;<font color="#999999">�����������գ��á�,���������������LyfUpload�����Ч������Ա������</font></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">ÿҳ��ʾ�ļ�����&nbsp;</td>
			<td align="left" class="tdborder"><input type="text" name="PerPageSize" value="<%=PerPageSize%>" class="txtinput"></td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">��½��ʾ��һ�ε�½ʱ�䣺&nbsp;</td>
			<td align="left" class="tdborder">
			����<input type="radio" name="showlasttime" value="1"<%if showlasttime=1 then response.write " checked"%> />&nbsp;&nbsp;
			�ر�<input type="radio" name="showlasttime" value="0"<%if showlasttime<>1 then response.write " checked"%> />
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">��IP��½ģʽ��&nbsp;</td>
			<td align="left" class="tdborder">
			����<input type="radio" name="LoginOneIP" value="1"<%if LoginOneIP=1 then response.write " checked"%> />&nbsp;&nbsp;
			�ر�<input type="radio" name="LoginOneIP" value="0"<%if LoginOneIP<>1 then response.write " checked"%> />&nbsp;<font color="#999999">��������һ���ʻ���ͬһʱ��ֻ����һ��IP��½</font>
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">��½ҳ���Զ���ʾ��֤�룺&nbsp;</td>
			<td align="left" class="tdborder">
			����<input type="radio" name="ShowValidate" value="1"<%if ShowValidate=1 then response.write " checked"%> />&nbsp;&nbsp;
			�ر�<input type="radio" name="ShowValidate" value="0"<%if ShowValidate<>1 then response.write " checked"%> />
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">ϵͳ����ǰ׺��&nbsp;</td>
			<td align="left" class="tdborder"><input type="text" name="SessionPrefix" value="<%=SessionPrefix%>" class="txtinput">&nbsp;<font color="#999999">���Ĵ������е�½�ŵ��û�����Ҫ���µ�½</font></td>
		  </tr>

		  <tr height="85">
			<td align="right" class="tdborder">����༭���ļ����ͣ�&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><textarea name="CanEditFileType" cols="50" rows="5"><%=CanEditFileType%></textarea></td>
					<td width="100%">&nbsp;<font color="#999999">�趨���Ա༭���ļ����ͣ���|������</font></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr height="85">
			<td align="right" class="tdborder">������༭���ļ����ͣ�&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><textarea name="CanNotEditFileType" cols="50" rows="5"><%=CanNotEditFileType%></textarea></td>
					<td width="100%">&nbsp;<font color="#999999">�趨���ɱ༭���ļ����ͣ���|������</font></td>
				  </tr>
				</table>
			</td>
		  </tr>

		  <tr>
			<td colspan="2" height="50" class="tdborder">
			<input type="submit" value="ȷ ��" class="button">&nbsp;&nbsp;<input type="reset" value="�� ��" class="button">&nbsp;&nbsp;<input type="button" value="�� ��" class="button" onClick="window.close()">
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
<%
sub saveconfig()
	on error resume next
	Script_Timeout=chkchar(0,trim(request.Form("Script_Timeout")),3600)
	UploadObject=chkchar(0,trim(request.Form("UploadObject")),0)
	FileMaxSize=chkchar(0,trim(request.Form("FileMaxSize")),0)
	NotAllowFileType=chkchar(1,trim(request.Form("NotAllowFileType")),"")
	AllowFileType=chkchar(2,trim(request.Form("AllowFileType")),"")
	PerPageSize=chkchar(0,trim(request.Form("PerPageSize")),40)
	showlasttime=chkchar(0,trim(request.Form("showlasttime")),1)
	LoginOneIP=chkchar(0,trim(request.Form("LoginOneIP")),1)
	ShowValidate=chkchar(0,trim(request.Form("ShowValidate")),0)
	SessionPrefix=chkchar(1,trim(request.Form("SessionPrefix")),"skymean.com")
	CanEditFileType=chkchar(1,trim(request.Form("CanEditFileType")),"")
	CanNotEditFileType=chkchar(1,trim(request.Form("CanNotEditFileType")),"")
	
	filecontent = chr(60) & "%@LANGUAGE=""VBSCRIPT"" CODEPAGE=""936""%" & chr(62) & vbcrlf
	filecontent = filecontent & chr(60) & "%" & vbcrlf
	filecontent = filecontent & "'ע�⣬��������ļ��������Ѿ������Ը����ˣ�ֻ���ڹ��������ò���" & vbcrlf
	filecontent = filecontent & "Option Explicit" & vbcrlf
	filecontent = filecontent & "Response.Buffer=True" & vbcrlf
	filecontent = filecontent & "Server.ScriptTimeOut=90" & vbcrlf
	filecontent = filecontent & "Session.CodePage=936" & vbcrlf&vbcrlf
	
	filecontent = filecontent & "dim starttime,SystemName,SystemVersion,selfname,userip,PerPageSize,CanEditFileType,CanNotEditFileType" & vbcrlf
	filecontent = filecontent & "dim showlasttime,copyright,OnlineValue,UploadObject,LoginOneIP,FileMaxSize,Script_Timeout" & vbcrlf
	filecontent = filecontent & "dim NotAllowFileType,AllowFileType,ShowValidate,SessionPrefix,endtime" & vbcrlf
	filecontent = filecontent & "starttime=timer()" & vbcrlf&vbcrlf

	filecontent = filecontent & "Script_Timeout=" & Script_Timeout & vbcrlf
	filecontent = filecontent & "UploadObject=" & UploadObject & vbcrlf
	filecontent = filecontent & "FileMaxSize=" & FileMaxSize & vbcrlf
	filecontent = filecontent & "NotAllowFileType=" & chr(34) & NotAllowFileType & chr(34) & vbcrlf
	filecontent = filecontent & "AllowFileType=" & chr(34) & AllowFileType & chr(34) & vbcrlf
	filecontent = filecontent & "CanEditFileType=" & chr(34) & CanEditFileType & chr(34) & vbcrlf
	filecontent = filecontent & "CanNotEditFileType=" & chr(34) & CanNotEditFileType & chr(34) & vbcrlf
	filecontent = filecontent & "PerPageSize=" & PerPageSize & vbcrlf
	filecontent = filecontent & "showlasttime=" & showlasttime & vbcrlf
	filecontent = filecontent & "LoginOneIP=" & LoginOneIP & vbcrlf
	filecontent = filecontent & "ShowValidate=" & ShowValidate & vbcrlf
	filecontent = filecontent & "SessionPrefix=" & chr(34) & SessionPrefix & chr(34) & vbcrlf&vbcrlf
	
	filecontent = filecontent & "copyright=""Copyright <font face='Times New Roman'>&copy;</font> 2005-2006 <a href='http://www.skymean.com' target=_blank><font color='#000080'><b>www.skymean.com</b></font></a> All Rights Reserved<br>Designed by ����""" & vbcrlf
	filecontent = filecontent & "SystemName=""���乤���������ļ�������""" & vbcrlf
	filecontent = filecontent & "SystemVersion=""v4.4""" & vbcrlf
	filecontent = filecontent & "selfname=Request.ServerVariables(""Script_Name"")" & vbcrlf
	filecontent = filecontent & "selfname=right(selfname,len(selfname)-instrrev(selfname,""/""))" & vbcrlf
	filecontent = filecontent & "userip=trim(Request.ServerVariables(""HTTP_X_FORWARDED_FOR""))" & vbcrlf
	filecontent = filecontent & "if userip="""" then" & vbcrlf
	filecontent = filecontent & "	userip=trim(Request.ServerVariables(""Remote_Addr""))" & vbcrlf
	filecontent = filecontent & "end if" & vbcrlf
	filecontent = filecontent & "OnlineValue=replace(cstr(date()),""-"","""") & ""OnlineUsers""" & vbcrlf
	filecontent = filecontent & "%" & chr(62) & vbcrlf
	filecontent = filecontent & chr(60) & "!--#include file=""dll.asp""--" & chr(62)
	
	if Err then Err.Clear
	set obj_fso=server.createobject("Scripting.FileSystemObject")
	set filetemp=obj_fso.createtextfile(server.mappath("config.asp"),true)
	filetemp.writeline(filecontent)
	filetemp.close
	set filetemp=nothing
	set obj_fso=nothing
	if Err then
		Err.Clear
		msg="<font color='#ff0000'>�����ļ�������ȷ��������֧��FSO����FSO��д�ļ�Ȩ�ޣ�</font>"
		exit sub
	end if
	msg="<font color='#ff0000'>�ɹ��޸�ϵͳ����</font>"
end sub

function chkchar(method,str,default)
	chkchar=default
	if trim(str)&""="" then
		exit function
	end if
	if method=0 then
		if isnumeric(str) then
			if str<0 then
				exit function
			end if
			chkchar=str
			exit function
		elseif isnumeric(default) then
			chkchar=default
			exit function
		end if
	elseif method=1 then
		chkchar=lcase(trim(str))
		chkchar=replace(chkchar,vbcrlf,"")
		chkchar=replace(chkchar,chr(34),"")
		chkchar=replace(chkchar," ","")
		chkchar=replace(chkchar,"��","|")
		chkchar=replace(chkchar,",","|")
		exit function
	elseif method=2 then
		chkchar=lcase(trim(str))
		chkchar=replace(chkchar,vbcrlf,"")
		chkchar=replace(chkchar,chr(34),"")
		chkchar=replace(chkchar," ","")
		chkchar=replace(chkchar,"|",",")
		chkchar=replace(chkchar,"��",",")
		exit function
	end if
end function


sub chkupobj()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<style type="text/css">
body{
	font-family: "����","Times New Roman", Times, serif;
	font-size:12px;
	color:#000000;
}
</style>
</head>
<body topmargin="0" leftmargin="0">
<div style="position:absolute; top:0px; left:0px;">
<%
	on error resume next
	dim upobj,upload
	upobj=trim(request("upobj"))
	if isnumeric(upobj) then
		upobj=cint(upobj)
	else
		response.write "��������"
		response.end
	end if
	If Err Then Err.Clear
	select case upobj
		case 1		'Aspupload ����ϴ�
			set upload=Server.CreateObject("Persits.Upload")
		case 2		'SA-FileUp ����ϴ�
			set upload=Server.CreateObject("SoftArtisans.FileUp")
		case 3		'LyfUpload ����ϴ�
			set upload=Server.CreateObject("LyfUpload.UploadFile")
		case else	'������ϴ�
			set upload=new DoteyUpload
	end select
	if Err then
		Err.Clear
		response.write "<b><font color='#ff0000'>��</font></b>&nbsp;��Ǹ����������֧��&nbsp;<font color='#ff0000'>"
		select case upobj
			case 1
				response.write "AspUpload���"
			case 2
				response.write "SA-FileUp���"
			case 3
				response.write "LyfUpload���"
			case else
				response.write "�����"
		end select
		response.write "</font>&nbsp;�ϴ�"
	else
		set upload=nothing
		response.write "<b><font color='#009900'>��</font></b>&nbsp;��ϲ��������֧��&nbsp;<font color='#009900'>"
		select case upobj
			case 1
				response.write "AspUpload���"
			case 2
				response.write "SA-FileUp���"
			case 3
				response.write "LyfUpload���"
			case else
				response.write "�����"
		end select
		response.write "</font>&nbsp;�ϴ�"
	end if
%>
</div>
</body>
</html>
<%end sub%>