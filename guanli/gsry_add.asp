<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<!--#include file="lianjie.asp"-->
<%
 tpdz=request("tpdz")
 tpsm=request("tpsm")
 entpsm=request("entpsm")
 
 leixing=request("leixing")
 tjsj=now()
 set rs=server.CreateObject("adodb.recordset")
rs.open"select * from gsry",cn,1,3
rs.addnew
rs("tpdz")=tpdz
rs("entpsm")=entpsm
rs("tpsm")=tpsm
rs("leixing")=leixing
rs("tjsj")=tjsj
rs("cptpjqsm")=request("cptpjqsm")
rs.update
rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>��ӳɹ�,<a href=gsry_list.asp?leixing="&leixing&">����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>

