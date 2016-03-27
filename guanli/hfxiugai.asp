<!--#include file="lianjie.asp"--->

<html>
<head><% 

 id=cint(trim(request("id")))
 nr=request("nr")
 sj=now()

  set rshf=server.CreateObject("adodb.recordset")
  rshf.open"select * from hfbook where dy_id="&id,cn,1,3
  rshf("nr")=nr
  rshf("sj")=sj
  rshf.update
url="sbcx_list.asp"
      response.write "<br><br><br><center><font size=-1>恭喜成功复.本页2秒钟后返回.<br><br>谢谢您使用本系统</font></center>"
     response.write "<meta http-equiv=refresh content=2;url=" & url & ">"
     response.end

%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

</body>
</html>
