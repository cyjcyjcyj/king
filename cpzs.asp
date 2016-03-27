
<!--#include file="conn.asp"-->

<% 
dlb_id=request("dlb_id")
 xlb_id=request("xlb_id")
 if dlb_id="" and xlb_id="" then
set rsyy=server.CreateObject("adodb.recordset")
rsyy.open"select * from jhcp",cn,1,3
end if
 if dlb_id<>"" then
set rsyy=server.CreateObject("adodb.recordset")
rsyy.open"select * from jhcp where dlb_id="&dlb_id,cn,1,3
set rsdlbok=server.CreateObject("adodb.recordset")
rsdlbok.open"select * from dlb where id="&dlb_id,cn,1,3
end if
 if xlb_id<>"" then
set rsyy=server.CreateObject("adodb.recordset")
rsyy.open"select * from jhcp where xlb_id="&xlb_id,cn,1,3
end if
dim pages,cur_recordcounts
pages=clng(request.QueryString("pages"))
rsyy.pagesize=8
if not rsyy.eof then
 if pages=0 then
 pages=1
  else
  rsyy.move(pages-1)*rsyy.pagesize
  end if
 if pages<rsyy.pagecount then
  cur_recordcounts=rsyy.pagesize
  else
   cur_recordcounts=rsyy.recordcount-(pages-1)*rsyy.pagesize
   end if
end if
%>

<html>
<head>
<script language="javascript">
var flag=false;
function DrawImage(ImgD,w,h){
var image=new Image();
image.src=ImgD.src;
if(image.width>0 && image.height>0){
flag=true;
if(image.width/image.height>= w/h){
if(image.width>w){
ImgD.width=w;
ImgD.height=(image.height*w)/image.width;
}else{
ImgD.width=image.width;
ImgD.height=image.height;
}
//ImgD.alt=image.width+"×"+image.height;
}
else{
if(image.height>h){
ImgD.height=h;
ImgD.width=(image.width*h)/image.height;
}else{
ImgD.width=image.width;
ImgD.height=image.height;
}
//ImgD.alt=image.width+"×"+image.height;
}
}
}
</script>
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
.right{width:260px;float:right;}
.body{width:1090px;margin:0 auto;}
.context{width:760px;min-height:800px;margin:0 20px 50px 0; text-align:center;font-family:"Microsoft Yahei light"}
.toppic{text-align:center;padding-left:20px;margin-bottom:50px}
-->
</style></head>
<body>
<!--#include file="top.asp"--->
<div class="body">

<div class="right"><!--#include file="rightsy.asp"---></div>
<div class="context">
<div class="toppic"><img src="images/toppic.png" alt=""></div>
            <font size="5">
              <%if dlb_id="" then
				response.Write ">>所有产品<<"
				else
				response.Write rsdlbok("dlb")
				end if%>
            </font>
<br /><br />
<div class="picture">
            
   <%for i=1 to cur_recordcounts%>
 
   <a href="cp_view.asp?id=<%=rsyy("id")%>" target="_blank"><img style="margin:15px" src="guanli/<%=rsyy("tpdz")%>"></a>

    <%if (i mod 2)=0 then response.write "</tr><tr>"%>
    <%rsyy.movenext
			next
    %>
</div>   
<br />
<br /><br />                      
 <p class="pages" align="center">
	<% 
	if pages>1 then
	response.Write"<a href=cpzs.asp?pages="&pages-1&"&dlb_id="&dlb_id&"&xlb_id="&xlb_id&">&#21069;&#19968;&#39029;</a>&nbsp;"
	else
	response.Write"&#21069;&#19968;&#39029;&nbsp;"
	end if
	if pages<=rsyy.pagecount-5 then
	start=pages-5
	else
	start=rsyy.pagecount-10
	end if
	if start<=0 then start=1
	for i=start to rsyy.pagecount 
	 if i=pages then
	 response.Write"&nbsp;"&i 
	 else 
	  if i<start+11 then
	  response.Write"&nbsp;<a href=cpzs.asp?pages="&i&"&dlb_id="&dlb_id&"&xlb_id="&xlb_id&">"&i&"</a>"
	  else
	  exit for
	 end if
	 end if
	 next 
	 if pages<rsyy.pagecount then
	 response.Write"&nbsp;<a href=cpzs.asp?pages="&pages+1&"&dlb_id="&dlb_id&"&xlb_id="&xlb_id&">&#19979;&#19968;&#39029;</a>"
	 else
	 response.Write"&nbsp;&#19979;&#19968;&#39029;"
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
