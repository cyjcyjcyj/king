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
response.write "<br><br><br><center>添加成功,<a href=gsry_list.asp?leixing="&leixing&">返回</a><br><br>谢谢您使用本系统</font></center>"
response.end

%>

