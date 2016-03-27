<!--#include file="conn.asp"-->
<html>
<head>
<title>广州锦倍塔信息科技有限公司</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<style type="text/css">
<!--
a:link {

	text-decoration: none;
}
a:visited {
	text-decoration: none;

}
a:hover {
	text-decoration: underline;

}
a:active {
	text-decoration: none;

}
.STYLE1 {color: #18455A}
.body{width:1050px;margin:0 auto;}
.context{width:760px;height:1000px;margin:0 20px 50px 0; padding-left:100px;}
.right{width:260px;float:right;}
-->
</style></head>
<body>
<!--#include file="top.asp"--->
<div class="body">

<div class="right"><!--#include file="rightsy.asp"---></div>
<div class="context">




        <table width="678" border=0 cellspacing=0 cellpadding=0 align=center>
          <tr>

            <a href="zxly.asp?dbid=d9"><img src="images/feedback_1.gif" width="79" height="23" border="0" align="absmiddle"></a> 
            <a href="zxly_view.asp?dbid=d9">&nbsp;<img src="images/feedback_2.gif" width="79" height="23" border="0" align="absmiddle"></a>

          </tr>
          
          <tr>
            <td colspan="2">
			<%
set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from Feedback where online='1' order by Postdate desc"
rs.open sql,cn,1
if rs.eof and rs.bof then response.redirect "zxly.asp"
if pages=0 or pages="" then pages=4	
rs.pageSize = pages
allPages = rs.pageCount	
page = Request("page")	

If not isNumeric(page) then page=1

if isEmpty(page) or Cint(page) < 1 then
page = 1
elseif Cint(page) >= allPages then
page = allPages 
end if
rs.AbsolutePage = page			
	Do While Not rs.eof and pages>0
	UserName=rs("UserName")		
	pic=rs("pic")			
	face=rs("face")			
	Comments=rs("Comments")		
	Replay=rs("Replay")		
	Usermail=rs("Usermail")		
	url=rs("url")			
	I=I+1				
	%>
                <table cellSpacing="1" cellPadding="3" width="641" align="center" bgColor="#CEBAA5" border="0">
                  <tr>
                    <td width="25%" rowSpan="2" align=center vAlign="middle" bgColor="#FFFFFF">
                    <table border=0 width=94%>
                        <tr>
                          <td width="35%" align=center><img src=images/face/pic<%=pic%>.gif width="40" height="40" border=0></td>
                          <td width="65%" align=center><div align="left">&nbsp;&nbsp;<%=UserName%></div></td>
                        </tr>
                        <tr>
                          <td colspan="2"> 来自：<%=left(rs("ip"),(len(rs("ip"))-1))+"*"%><br>
                            邮件：<a href=mailto:<%=Usermail%>><img src=images/mail.gif border=0></a><br>
                            主页：<a href="<%=URL%>" target='_blank'><img src=images/home.gif border=0></a></td>
                        </tr>
                    </table>
                    </td>
                    <td width="75%" height="20" bgColor="#FFFFFF">&nbsp;[NO.<%=(cint(page)-1)*rs.pagesize+I%>] <img border=0 src=images/face/face<%=face%>.gif> 发表于：<%=rs("Postdate")%></td>
                  </tr>
                  <tr>
                    <td width='75%' height=90 vAlign="top"  bgColor="#FFFFFF" onMouseOver="bgColor='#FFffff'" onMouseOut="bgColor='#ebebeb'"><%
	if html=0 then
	response.write replace(server.htmlencode(Comments),vbCRLF,"<BR>")
	else
	response.write replace(Comments,vbCRLF,"<BR>")
	end if
	%>
                        <br>
                        <%if rs("Replay")<>"" then%>
                        <table cellSpacing="1" cellPadding="3" width="90%" align="center" bgColor="darkgray" border="0">
                          <tr>
                            <td height="26" vAlign="top" bgColor="#184963"><font color=red>客服回复： <%=Replay%></font> </td>
                          </tr>
                        </table>
                      <%end if%></td>
                  </tr>
                </table>
              <br>
                <%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
%>
            </td>
          </tr>
          <tr>
            <td height=20 colspan="2" valign=top><%
response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;总计留言"&RS.RecordCount&"条 "
if page = 1 then
response.write "<font color=darkgray>首页 前页</font>"
else
response.write "<a href="&request.ServerVariables("script_name")&"?page=1&dbid=d5>首页</a> <a href="&request.ServerVariables("script_name")&"?page="&page-1&">前页</a>"
end if
if page = allpages then
response.write "<font color=darkgray> 下页 末页</font>"
else
response.write " <a href="&request.ServerVariables("script_name")&"?page="&page+1&"&dbid=d5>下页</a> <a href="&request.ServerVariables("script_name")&"?page="&allpages&"&dbid=d5>末页</a>"
end if
response.write " 第"&page&"页 共"&allpages&"页 "
%>
            </td>
          </tr>
        </table>


</div>

</div>

<!--#include file="bottom.asp"--->

<!-- ImageReady Slices (okk1.psd) -->
<!-- End ImageReady Slices -->
</body>
</html>
