<div align='center' id='demo' style='overflow:hidden;height:140px;width:900px;'> 
                    <!--滚动区的高度和宽度-->
<table width="799" border='0' align='center' cellpadding='0' cellspacing="0" cellspace='0'>
<tr> 
<td width="768" valign='top' id='demo1'><table width="400" border="0" cellspacing="0" cellpadding="0">
<tr> <td width="249" valign="top"><a href="gd/s1.jpg" target="_blank"><img src="gd/s1.jpg" width="179" height="91" hspace="12" border="0" /></a></td>
              <td width="249" valign="top"><a href="gd/1.jpg" target="_blank"><img src="gd/1.jpg" width="179" height="91" hspace="12" border="0" /></a></td>             
            <td width="249" valign="top"><a href="gd/2.jpg" target="_blank"><img src="gd/2.jpg" width="179" height="91" hspace="12" border="0" /></a></td>
            <td width="249" valign="top"><a href="gd/3.jpg" target="_blank"><img src="gd/3.jpg" width="179" height="91"  hspace="12" border="0"/></a></td>
            <td width="249" height="102" valign="top" bordercolor="4"><a href="gd/4.jpg" target="_blank"><img src="gd/4.jpg" width="179" height="91" hspace="12" border="0"/></a></td>
                              <!--pkvs:　图片数据循环调用开始--->
<% set rsxw=server.CreateObject("adodb.recordset")
		 rsxw.open"select * from jhcp order by tjsj desc",cn,1,3
   for i=1 to rsxw.recordcount
      %>
   
<td width="249" rowspan="2" valign="top">
<table width="207" border="0" cellpadding="0" cellspacing="0">
<tr> 
<td width="207" height="102" align="center" valign="top" background="images/bgcpok.jpg" > 
<table width="62" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="62"><a href="guanli/<%=rsxw("tpdz")%>" target="_blank"><img src="guanli/<%=rsxw("tpdz")%>"  border="0" width="179" height="91" /></a><a href="cp_view.asp?id=<%=rsxw("id")%>&dlb_id=<%=rsxw("dlb_id")%>&xlb_id=<%=rsxw("xlb_id")%>" target="_blank"></a></td>
  </tr>
</table>
<a href="cp_view.asp?id=<%=rsxw("id")%>" target="_blank"></a></td>
</tr>
<tr>
  <td height="20" align="center" valign="middle"><%=rsxw("cpmc")%></td>
</tr>
</table></td>
<% rsxw.movenext
next
rsxw.close
set rsxw=nothing%>
                              <!--pkvs:　图片数据循环调用结束--->
</tr>
<tr>
  <td width="249" align="center" valign="top">系统</td>
  <td width="249" align="center" valign="top">系统</td>
  <td width="249" align="center" valign="top">系统</td>
  <td align="center" valign="top" bordercolor="4">系统</td>
  
  <td align="center" valign="top" bordercolor="4">系统</td>
</tr>
</table></td>
<td width="18" valign=top id=demo2></td>
</tr>
</table>
</div>
<script>
var speed=1
demo2.innerHTML=demo1.innerHTML
function Marquee1(){
if(demo2.offsetWidth-demo.scrollLeft<=0)
demo.scrollLeft-=demo1.offsetWidth
else{
demo.scrollLeft++
}
}
var MyMar1=setInterval(Marquee1,speed)
demo.onmouseover=function() {clearInterval(MyMar1)}
demo.onmouseout=function() {MyMar1=setInterval(Marquee1,speed)}
</script>
