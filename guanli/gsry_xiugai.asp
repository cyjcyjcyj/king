<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<!--#include file="lianjie.asp"-->
<% id=cint(trim(request("id")))
 tpdz=request("tpdz")
 tpsm=request("tpsm")
 entpsm=request("entpsm")
 leixing=request("leixing")
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from gsry where id="&id,cn,1,3
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
response.write "<br><br><br><center>�޸ĳɹ�,<a href=gsry_list.asp?leixing="&leixing&">����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>

