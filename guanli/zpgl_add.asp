<!--#include file="conn.asp"----->
<% 
zpdx=request("zpdx")
zprs=request("zprs")
gzdd=request("gzdd")
gzdy=request("gzdy")
fbsj=request("fbsj")
yxqx=request("yxqx")
zpyq=request("zpyq")
 set rszp=server.CreateObject("adodb.recordset")
 rszp.open"select * from job ",cn,1,3
 rszp.addnew
  rszp("zpdx")=zpdx
  rszp("zprs")=zprs
  rszp("gzdd")=gzdd
  rszp("gzdy")=gzdy
  rszp("zpdx")=zpdx
  rszp("fbsj")=fbsj
  rszp("zprs")=zprs
  rszp("zpdx")=zpdx
    rszp("yxqx")=yxqx
	rszp("zpyq")=zpyq
 rszp.update
   url="zpgl.asp"
      response.write "<br><br><br><center><font size=-1>��ӳɹ�,��ҳ2���Ӻ󷵻�.<br><br>лл��ʹ�ñ�ϵͳ</font></center>"
     response.write "<meta http-equiv=refresh content=2;url=" & url & ">"
     response.end%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>

<body>

</body>
</html>
