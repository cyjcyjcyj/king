<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<!--#include file="lianjie.asp"-->
<%id=cint(trim(request("id")))
 class_id=request("class_id")
 btmc=request("btmc")
nrjs=request("nrjs")
down1=request("down1")
 down2=request("down2")
enbtmc=request("enbtmc")
 ennrjs=request("ennrjs")
tjsj=now()
 if btmc="" or nrjs="" or down1="" or down2="" or ennrjs="" or enbtmc=""  then
 response.Write"<script>alert('�Բ��𣬸����Ϊ��');history.back();</script>"
 response.End()
 end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from down where id="&id,cn,1,3
rs("class_id")=class_id
rs("btmc")=btmc
rs("nrjs")=nrjs
rs("down1")=down1
rs("down2")=down2
rs("enbtmc")=enbtmc
rs("ennrjs")=ennrjs
rs("tjsj")=tjsj
rs.update
rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>�޸ĳɹ�,<a href=down_list.asp>����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>
