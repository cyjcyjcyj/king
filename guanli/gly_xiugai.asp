<!--#include file="lianjie.asp"-->
<!--#include file="md5.asp"------>
<%id=cint(trim(request("id")))
yhm=replace(trim(request("yhm")),"'","")
mm=replace(trim(request("mm")),"'","")
remm=replace(trim(request("remm")),"'","")
if yhm="" or mm="" then
response.Write"<script>alert('对不起，各项不能为空');history.back();</script>"
response.End()
end if
if mm<>remm then
response.Write"<script>alert('两次密码输入不一致');history.back();</script>"
response.End()
end if
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from guanliyuan where id="&id,cn,1,3
rs("yhm")=yhm
rs("mm")=md5(mm)
rs("remm")=md5(remm)
rs("tjsj")=now()
rs.update
url="gly_list.asp"
      response.write "<br><br><br><center><font size=-1>修改成功,本页1秒钟后返回.<br><br>谢谢您使用本系统</font></center>"
     response.write "<meta http-equiv=refresh content=1;url=" & url & ">"
     response.end




%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

</body>
</html>

