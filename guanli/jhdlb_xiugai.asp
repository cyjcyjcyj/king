<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style><!--#include file="lianjie.asp"-->

<%id=cint(trim(request("id")))
dlb=replace(trim(request("dlb")),"'","")
endlb=replace(trim(request("endlb")),"'","")
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from dlb where id="&id,cn,1,3
rs("dlb")=dlb
rs("endlb")=endlb
rs.update
url="dlbxiugai.asp?id="&id
  
  response.write "<br><br><br><center>�޸ĳɹ�,<a href="&url&">����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end
rs.close
set rs=nothing
cn.close
set cn=nothing



%>


