<!--#include file="checklogin.asp"-->
<!--#include file="incupload.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

if not popedom then
	call showmsg("你没有权限进入此页面！","index.asp")
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
<title><%=SystemName%> - 系统配置</title>
<link rel="stylesheet" type="text/css" href="images/config.css" />
</head>
<body>
<div id="bodyposition">
	<div id="content">
	  <div id="title"><%=SystemName%> - 系统配置</div>
	  <%if msg<>"" then response.write "<div id='msg' style='border-top:1px #000099 solid;'>"&msg&"</div>"%>
	  <div id="main" style="border-top:1px #000099 solid;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		 <form method="post" action="<%=selfname%>">
		 <input type="hidden" name="save" value="yes">
		  <tr><td colspan="2" height="5"></td></tr>
		  <tr height="30">
			<td align="right" width="22%">上传脚本超时时间：&nbsp;</td>
			<td align="left" width="78%"><input type="text" name="Script_Timeout" value="<%=Script_Timeout%>" class="txtinput">&nbsp;秒&nbsp;<font color="#999999">建议不要设置得太短</font></td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">上传组件类型：&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td>
						<select id="UploadObject" name="UploadObject" size="1" onChange="document.getElementById('chkobjpage').src='<%=selfname%>?act=chkobj&upobj='+document.getElementById('UploadObject').value">
						<option value="0"<%if UploadObject=0 then response.write " selected"%>>无组件上传类</option>
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
			<td align="right" class="tdborder">上传文件大小：&nbsp;</td>
			<td align="left" class="tdborder"><input type="text" name="FileMaxSize" value="<%=FileMaxSize%>" class="txtinput">&nbsp;字节&nbsp;<font color="#999999">不限制请设为0，管理员不受限</font></td>
		  </tr>
		  <tr height="85">
			<td align="right" class="tdborder">禁止上传的文件类型：&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><textarea name="NotAllowFileType" cols="50" rows="5"><%=NotAllowFileType%></textarea></td>
					<td width="100%">&nbsp;<font color="#999999">不限制请留空，用“|”隔开；此项对非LyfUpload组件有效，管理员不受限</font></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr height="85">
			<td align="right" class="tdborder">允许上传的文件类型：&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><textarea name="AllowFileType" cols="50" rows="5"><%=AllowFileType%></textarea></td>
					<td width="100%">&nbsp;<font color="#999999">不限制请留空，用“,”隔开；此项仅对LyfUpload组件有效，管理员不受限</font></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">每页显示文件数：&nbsp;</td>
			<td align="left" class="tdborder"><input type="text" name="PerPageSize" value="<%=PerPageSize%>" class="txtinput"></td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">登陆显示上一次登陆时间：&nbsp;</td>
			<td align="left" class="tdborder">
			开启<input type="radio" name="showlasttime" value="1"<%if showlasttime=1 then response.write " checked"%> />&nbsp;&nbsp;
			关闭<input type="radio" name="showlasttime" value="0"<%if showlasttime<>1 then response.write " checked"%> />
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">单IP登陆模式：&nbsp;</td>
			<td align="left" class="tdborder">
			开启<input type="radio" name="LoginOneIP" value="1"<%if LoginOneIP=1 then response.write " checked"%> />&nbsp;&nbsp;
			关闭<input type="radio" name="LoginOneIP" value="0"<%if LoginOneIP<>1 then response.write " checked"%> />&nbsp;<font color="#999999">若开启则一个帐户在同一时刻只能在一个IP登陆</font>
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">登陆页面自动显示验证码：&nbsp;</td>
			<td align="left" class="tdborder">
			开启<input type="radio" name="ShowValidate" value="1"<%if ShowValidate=1 then response.write " checked"%> />&nbsp;&nbsp;
			关闭<input type="radio" name="ShowValidate" value="0"<%if ShowValidate<>1 then response.write " checked"%> />
			</td>
		  </tr>
		  <tr height="30">
			<td align="right" class="tdborder">系统变量前缀：&nbsp;</td>
			<td align="left" class="tdborder"><input type="text" name="SessionPrefix" value="<%=SessionPrefix%>" class="txtinput">&nbsp;<font color="#999999">更改此项所有登陆着的用户都需要重新登陆</font></td>
		  </tr>

		  <tr height="85">
			<td align="right" class="tdborder">允许编辑的文件类型：&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><textarea name="CanEditFileType" cols="50" rows="5"><%=CanEditFileType%></textarea></td>
					<td width="100%">&nbsp;<font color="#999999">设定可以编辑的文件类型，“|”隔开</font></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr height="85">
			<td align="right" class="tdborder">不允许编辑的文件类型：&nbsp;</td>
			<td align="left" class="tdborder">
				<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td><textarea name="CanNotEditFileType" cols="50" rows="5"><%=CanNotEditFileType%></textarea></td>
					<td width="100%">&nbsp;<font color="#999999">设定不可编辑的文件类型，“|”隔开</font></td>
				  </tr>
				</table>
			</td>
		  </tr>

		  <tr>
			<td colspan="2" height="50" class="tdborder">
			<input type="submit" value="确 定" class="button">&nbsp;&nbsp;<input type="reset" value="重 置" class="button">&nbsp;&nbsp;<input type="button" value="关 闭" class="button" onClick="window.close()">
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
	filecontent = filecontent & "'注意，这个配置文件的内容已经不可以更改了，只能在管理里设置参数" & vbcrlf
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
	
	filecontent = filecontent & "copyright=""Copyright <font face='Times New Roman'>&copy;</font> 2005-2006 <a href='http://www.skymean.com' target=_blank><font color='#000080'><b>www.skymean.com</b></font></a> All Rights Reserved<br>Designed by 秋忆""" & vbcrlf
	filecontent = filecontent & "SystemName=""秋忆工作室在线文件管理器""" & vbcrlf
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
		msg="<font color='#ff0000'>保存文件错误，请确保服务器支持FSO并且FSO有写文件权限！</font>"
		exit sub
	end if
	msg="<font color='#ff0000'>成功修改系统配置</font>"
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
		chkchar=replace(chkchar,"｜","|")
		chkchar=replace(chkchar,",","|")
		exit function
	elseif method=2 then
		chkchar=lcase(trim(str))
		chkchar=replace(chkchar,vbcrlf,"")
		chkchar=replace(chkchar,chr(34),"")
		chkchar=replace(chkchar," ","")
		chkchar=replace(chkchar,"|",",")
		chkchar=replace(chkchar,"，",",")
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
	font-family: "宋体","Times New Roman", Times, serif;
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
		response.write "参数错误！"
		response.end
	end if
	If Err Then Err.Clear
	select case upobj
		case 1		'Aspupload 组件上传
			set upload=Server.CreateObject("Persits.Upload")
		case 2		'SA-FileUp 组件上传
			set upload=Server.CreateObject("SoftArtisans.FileUp")
		case 3		'LyfUpload 组件上传
			set upload=Server.CreateObject("LyfUpload.UploadFile")
		case else	'无组件上传
			set upload=new DoteyUpload
	end select
	if Err then
		Err.Clear
		response.write "<b><font color='#ff0000'>×</font></b>&nbsp;抱歉，服务器不支持&nbsp;<font color='#ff0000'>"
		select case upobj
			case 1
				response.write "AspUpload组件"
			case 2
				response.write "SA-FileUp组件"
			case 3
				response.write "LyfUpload组件"
			case else
				response.write "无组件"
		end select
		response.write "</font>&nbsp;上传"
	else
		set upload=nothing
		response.write "<b><font color='#009900'>√</font></b>&nbsp;恭喜，服务器支持&nbsp;<font color='#009900'>"
		select case upobj
			case 1
				response.write "AspUpload组件"
			case 2
				response.write "SA-FileUp组件"
			case 3
				response.write "LyfUpload组件"
			case else
				response.write "无组件"
		end select
		response.write "</font>&nbsp;上传"
	end if
%>
</div>
</body>
</html>
<%end sub%>