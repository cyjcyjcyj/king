<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<!--#include file="lianjie.asp"-->
<%
 cpmc=request("cpmc")
 cpxh=request("cpxh")
 cpjscs=request("cpjscs")
 cpgg=request("cpgg")
 sfsytj=request("sfsytj")
  cptpjqsm=request("cptpjqsm")
 tpdz=request("tpdz")
 tjsj=now()
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from cp",cn,1,3
rs.addnew
rs("cpmc")=cpmc
rs("cpxh")=cpxh
rs("cpjscs")=cpjscs
rs("cpgg")=cpgg
rs("tpdz")=tpdz
rs("sfsytj")=sfsytj
rs("cptpjqsm")=cptpjqsm
rs("tjsj")=tjsj
rs.update
rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>��ӳɹ�,<a href=cp_list.asp>����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>
