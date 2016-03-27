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
 tpdz1=request("tpdz1")
 dlb_id=request("rootid")
 xlb_id=request("classid")
 cptpjqsm=request("cptpjqsm")
 encptpjqsm=request("encptpjqsm")
 cpmc=request("cpmc")
 encpmc=request("encpmc")
 sfsytj=request("sfsytj")
 tjsj=now()
 

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from jhcp",cn,1,3
rs.addnew
rs("cptpjqsm")=cptpjqsm
rs("encptpjqsm")=encptpjqsm
rs("sfsytj")=sfsytj
rs("cpmc")=cpmc
rs("encpmc")=encpmc
rs("tjsj")=tjsj
On Error resume next
rs("dlb_id")=dlb_id
rs("xlb_id")=xlb_id
rs("tpdz")=tpdz
rs("tpdz1")=tpdz1
rs.update
rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>添加成功,<a href=dckt_list.asp>返回</a><br><br>谢谢您使用本系统</font></center>"
response.end

%>
