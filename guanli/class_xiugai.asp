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
response.Write"<script>alert('�Բ��𣬸����Ϊ��');history.back();</script>"
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
response.write "<br><br><br><center>�޸ĳɹ�,<a href=class_list.asp>����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>
