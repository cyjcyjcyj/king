<!--#include file="lianjie.asp"-->
<!--#include file="md5.asp"------>
<%id=cint(trim(request("id")))
yhm=replace(trim(request("yhm")),"'","")
mm=replace(trim(request("mm")),"'","")
remm=replace(trim(request("remm")),"'","")
if yhm="" or mm="" then
response.Write"<script>alert('�Բ��𣬸����Ϊ��');history.back();</script>"
response.End()
end if
if mm<>remm then
response.Write"<script>alert('�����������벻һ��');history.back();</script>"
response.End()
end if
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from guanliyuan where yhm='"&yhm&"'",cn,1,3
if not rs.eof then
response.Write"<script>alert('�Բ����û����ظ�');history.back();</script>"
else 
rs.addnew
rs("yhm")=yhm
rs("mm")=md5(mm)
rs("remm")=md5(remm)
rs("tjsj")=now()
rs.update
end if
response.Redirect("gly_list.asp")




%>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

