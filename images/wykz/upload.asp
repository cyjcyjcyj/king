<!--#include file="checklogin.asp"-->
<%
dim SPid,UploadProgress,ismsie
ismsie=msie()
call aspuppid()
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - 文件上传</title>
<link rel="stylesheet" type="text/css" href="images/upload.css" />
<script language="javascript">
<!--
var uploadobj=<%=UploadObject%>;
var uploadpid="<%=SPid%>";
var IsMSIE=<%=lcase(cstr(ismsie))%>;
//-->
</script>
<script type="text/javascript" src="images/upload.js"></script>
</head>
<body>
<div id="bodyposition">
<div id="uploadadmain">
<div id="uploadr1c1"></div>
<div id="content">
	<div id="title"><%=SystemName%> - 文件上传</div>
	<div id="control">
		<div class="<%if ismsie then response.Write("contie"):else response.Write("contns"):end if%>"><br>
        <li> 需要上传的文件个数 
<%if UploadObject=3 then%>
          <input type="text" id="upcount" class="tx" value="5" readonly="true" style="width:60px">
          <input type="button" class="button" disabled="disabled" value="· 设定 ·" style="background-color:#EBEBEB"><br>（LyfUpload组件 固定设置每次最多上传5个文件）<!--无办法，水平有限目前没有好办法设计能够动态设置LyfUpload上传文件个数，你有好方法请赐教-->
<%else%>
          <input type="text" id="upcount" class="tx" value="1" onKeyDown="if(window.event.keyCode==13){SetInput();return false;}" style="width:60px">
          <input type="button" id="setinput" onClick="SetInput()" value="设 定">
<%end if%>
		<input type="hidden" id="upcount2" name="upcount2" value="1">
        </li>
		<br><%if ismsie then response.Write("<br>")%>
        <li>上传到: 
          <input type="text" id="visualpath" class="tx" style="width:360px" value="<%=path%>">
        </li>
<%if UploadObject<>3 and trim(NotAllowFileType)<>"" and not popedom then%>
		<br><%if ismsie then response.Write("<br>")%>
        <li>不可上传文件类型：<font color="#CC0000"><%=replace(NotAllowFileType,"|","、")%></font></li>
<%
else
	if trim(AllowFileType)<>"" and not popedom then
%>
		<br><%if ismsie then response.Write("<br>")%>
        <li>可上传文件类型：<font color="#CC0000"><%=replace(AllowFileType,",","、")%></font></li>
<%
	end if
end if

if FileMaxSize>0 and not popedom then
%>
		<br><%if ismsie then response.Write("<br>")%>
        <li>最大文件：<font color="#CC0000"><%=round(FileMaxSize/1048576,3)%> MB</font></li>
<%end if%>
		<br><%if ismsie then response.Write("<br>")%>
		</div>
	</div>
	<div id="upinput">
	<form name="form1" method="post" action="upprocess.asp?path=<%=Server.URLEncode(path)%>" enctype="multipart/form-data">
	<input type="hidden" name="act" value="uploadfile">
	<input type="hidden" id="realpath" name="filepath" value="">
	<div id="upid"><span id="upspan1">文件01: <input type="file" id="file1" name="file1" class="tx1"<%if not ismsie then response.Write(" size='55'")%> value=""></span></div>
	</form><%if not ismsie then response.Write("<br>")%>
	<script language="javascript">SetInput();</script>
	</div>
<%
if trim(path)<>"" then
	path=replace(path,"\","\\")
	path=replace(path,"'","\'")
	path=replace(path,chr(34),"\"&chr(34))
	path=replace(path,"+","\+")
end if
%>
	<div id="submit"><br>
        <input type="button" id="submitbutton" value="上 传" class="button" onClick="UpSubmit()">&nbsp;&nbsp;&nbsp;
        <input type="button" value="重 填" class="button" onClick="resetform('<%=path%>')">&nbsp;&nbsp;&nbsp;
		<input type="button" value="关 闭" class="button" onClick="window.close()">
	</div>
</div>
<div id="uploadr3c2"></div>
<div id="footer"><br><%
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
</body>
</html>
<%
chksystem()
sub aspuppid()
	on error resume next
	if UploadObject=1 then
		if Err then Err.Clear
		Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
		SPid = UploadProgress.CreateProgressID()
		if Err then
			Err.Clear
			response.Write "<font color='#ff0000'>警告，服务器 AspUpload组件 不支持进度上传！可能导致文件上传失败</font>"
		end if
	end if
end sub
%>