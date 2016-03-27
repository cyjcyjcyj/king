<!--#include file="lianjie.asp"-->
<!--#include file="md5.asp"------>
<%id=cint(trim(request("id")))
wzmc=replace(trim(request("wzmc")),"'","")
wzwz=replace(trim(request("wzwz")),"'","")
ljlogo=replace(trim(request("ljlogo")),"'","")
sfstplj=replace(trim(request("sfstplj")),"'","")

if wzmc="" or wzwz="" then
response.Write"<script>alert('对不起，各项不能为空');history.back();</script>"
response.End()
end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from yqlj where id="&id,cn,1,3
rs("wzwz")=wzwz
rs("wzmc")=wzmc
rs("ljlogo")=ljlogo
rs("jrsj")=now()
rs("sfstplj")=sfstplj
rs.update
rs.close
set rs=nothing
url="yqlj_list.asp"
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

