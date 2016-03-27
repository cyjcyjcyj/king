<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style><!--#include file="lianjie.asp"-->

<%
dlb=request("dlb")
endlb=request("endlb")

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from dlb where dlb='"&dlb&"'",cn,1,3
if not rs.eof then
response.Write"<script>alert('对不起，有类别重复');history.back();</script>"
else 
rs.addnew
rs("dlb")=dlb
rs("endlb")=endlb
rs.update
end if
rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>添加成功,<a href=cplb_list.asp>返回</a><br><br>谢谢您使用本系统</font></center>"
response.end

%>
