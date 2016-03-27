<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style><!--#include file="lianjie.asp"-->

<%dlb_id=request("dlb_id")
  xlb=request("xlb")
  enxlb=request("enxlb")
if xlb="" or enxlb="" then
response.Write"<script>alert('对不起，各项不能为空');history.back();</script>"
response.End()
end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from jhxlb where xlb='"&xlb&"'",cn,1,3

rs.addnew
rs("xlb")=xlb
rs("dlb_id")=dlb_id
rs("enxlb")=enxlb
rs.update

rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>添加成功,<a href=jhlb_list.asp>返回</a><br><br>谢谢您使用本系统</font></center>"
response.end

%>
