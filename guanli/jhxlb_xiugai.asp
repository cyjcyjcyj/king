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
if xlb="" or enxlb="" then
response.Write"<script>alert('�Բ��𣬸����Ϊ��');history.back();</script>"
response.End()
end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from jhxlb where id="&id,cn,1,3


rs("xlb")=xlb
rs("dlb_id")=dlb_id
rs("enxlb")=enxlb
rs.update

rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>�޸ĳɹ�,<a href=jhlb_list.asp>����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>
