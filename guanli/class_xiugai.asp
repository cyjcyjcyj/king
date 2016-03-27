<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style><!--#include file="lianjie.asp"-->

<%id=cint(trim(request("id")))
xxlb=request("yjfl")
enxxlb=request("enyjfl")
if xxlb="" or enxxlb="" then
response.Write"<script>alert('对不起，各项不能为空');history.back();</script>"
response.End()
end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from class where id="&id,cn,1,3

rs("enxxlb")=enxxlb
rs("xxlb")=xxlb
rs.update

rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>修改成功,<a href=class_list.asp>返回</a><br><br>谢谢您使用本系统</font></center>"
response.end

%>
