<!--#include file="lianjie.asp"-->
<!--#include file="md5.asp"------>
<%
wzmc=replace(trim(request("wzmc")),"'","")
wzwz=replace(trim(request("wzwz")),"'","")
ljlogo=replace(trim(request("ljlogo")),"'","")
sfstplj=replace(trim(request("sfstplj")),"'","")

if wzmc="" or wzwz="" then
response.Write"<script>alert('�Բ��𣬸����Ϊ��');history.back();</script>"
response.End()
end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from yqlj",cn,1,3
rs.addnew
rs("wzwz")=wzwz
rs("wzmc")=wzmc
rs("ljlogo")=ljlogo
rs("sfstplj")=sfstplj
rs("jrsj")=now()
rs.update
rs.close
set rs=nothing
url="yqlj_list.asp"
      response.write "<br><br><br><center><font size=-1>��ӳɹ�,��ҳ1���Ӻ󷵻�.<br><br>лл��ʹ�ñ�ϵͳ</font></center>"
     response.write "<meta http-equiv=refresh content=1;url=" & url & ">"
     response.end




%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>

<body>

</body>
</html>

