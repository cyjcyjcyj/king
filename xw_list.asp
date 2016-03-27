<!--#include file="conn.asp"--->
<% xwlb_id=request("xwlb_id")

if xwlb_id="" then
set  rsxw=server.CreateObject("adodb.recordset")
rsxw.open"select * from xw order by tjsj desc",cn,1,3
set rsxwlb=server.CreateObject("adodb.recordset")
end if
if xwlb_id<>"" then
set rsxw=server.CreateObject("adodb.recordset")
set  rsxwlb=server.CreateObject("adodb.recordset")
rsxw.open "select * from xw where xwlb_id="&xwlb_id&" order by tjsj desc",cn,1,3
rsxwlb.open"select * from xwlb where id="&xwlb_id,cn,1,3
end if

dim pages,cur_recordcounts
pages=cint(request.QueryString("pages"))
rsxw.pagesize=15
if not rsxw.eof then
 if pages=0 then
 pages=1
  else
  rsxw.move(pages-1)*rsxw.pagesize
  end if
 if pages<rsxw.pagecount then
  cur_recordcounts=rsxw.pagesize
  else
   cur_recordcounts=rsxw.recordcount-(pages-1)*rsxw.pagesize
   end if
end if%>
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
.body{width:1090px;margin:0 auto;}
.context{width:760px;min-height:800px;margin:0 20px 50px 0;}
.right{width:260px; float:right;}
.data{float:right;margin:0px 20px;}
.href{margin-left:100px;}
.toppic{text-align:center;padding-left:20px;margin-bottom:100px;}
-->
</style></head>
<body>
<!--#include file="top.asp"--->
<div class="body">
<div class="right"><!--#include file="rightsy.asp"---></div>

<div class="context">
<div class="toppic"><img src="images/toppic.png" alt=""></div>
<%for i=1 to cur_recordcounts%>
<div class="data">[<%=rsxw("tjsj")%>]</div>
<div class="href"><a href="xw_view.asp?lb=2&id=<%=rsxw("id")%>&xwlb_id=<%=xwlb_id%>&dbid=d2"><%=rsxw("xwbt")%></a></div>
<br />

<%rsxw.movenext
next%>
<br /><br /><br />
<p class="pages" align="center">
<% 
	if pages>1 then
	response.Write"<a href=xw_list.asp?pages="&pages-1&"&xwlb_id="&xwlb_id&"&xlb_id="&xlb_id&"&lb="&lb&">前一页</a>&nbsp;"
	else
	response.Write"前一页&nbsp;"
	end if
	if pages<=rsxw.pagecount-5 then
	start=pages-5
	else
	start=rsxw.pagecount-10
	end if
	if start<=0 then start=1
	for i=start to rsxw.pagecount 
	 if i=pages then
	 response.Write"&nbsp;"&i 
	 else 
	  if i<start+11 then
	  response.Write"&nbsp;<a href=xw_list.asp?pages="&i&"&xwlb_id="&xwlb_id&"&xlb_id="&xlb_id&"&lb="&lb&">"&i&"</a>"
	  else
	  exit for
	 end if
	 end if
	 next 
	 if pages<rsxw.pagecount then
	 response.Write"&nbsp;&nbsp;<a href=xw_list.asp?pages="&pages+1&"&xwlb_id="&xwlb_id&"&xlb_id="&xlb_id&"&lb="&lb&">下一页</a>"
	 else
	 response.Write"&nbsp;&nbsp;下一页"
	 end if
	 %>
</p>
 </div>  


</div>

<!--#include file="bottom.asp"--->


<!-- ImageReady Slices (okk1.psd) -->
<!-- End ImageReady Slices -->
</body>
</html>
