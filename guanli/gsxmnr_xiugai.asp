<!--#include file="lianjie.asp"-->
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<%id=cint(trim(request("id")))
  gsxmnr=replace(trim(request("gsxmnr")),"'","")
  engsxmnr=replace(trim(request("engsxmnr")),"'","")
  tjsj=now()

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from gsxmnr where id="&id,cn,1,3
rs("gsxmnr")=gsxmnr
rs("engsxmnr")=engsxmnr
rs("tjsj")=tjsj
rs.update
rs.close
set rs=nothing
cn.close
set cn=nothing
url="gsxmnrxiugai.asp?id="&id
  response.write "<br><br><br><center>�޸ĳɹ�,<a href="&url&">����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end





%>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

