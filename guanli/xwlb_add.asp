<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style><!--#include file="lianjie.asp"-->

<%
xwlb=request("yjfl")
enxwlb=request("enyjfl")


set rs=server.CreateObject("adodb.recordset")
rs.open"select * from xwlb where xwlb='"&xwlb&"'",cn,1,3
if not rs.eof then
response.Write"<script>alert('�Բ���������ظ�');history.back();</script>"
else 
rs.addnew
rs("xwlb")=xwlb
rs("enxwlb")=enxwlb
rs.update
end if
rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>��ӳɹ�,<a href=xwlblist.asp>����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>
