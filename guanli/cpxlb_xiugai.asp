<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style><!--#include file="lianjie.asp"-->

<% id=cint(trim(request("id")))
  dlb_id=request("dlb_id")
  xlb=request("xlb")
  enxlb=request("enxlb")


set rs=server.CreateObject("adodb.recordset")
rs.open"select * from xlb where id="&id,cn,1,3


rs("xlb")=xlb
rs("dlb_id")=dlb_id
rs("enxlb")=enxlb
rs.update

rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>�޸ĳɹ�,<a href=cplb_list.asp>����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>
