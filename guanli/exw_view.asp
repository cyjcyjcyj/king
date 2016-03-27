<!--#include file="conn.asp"-->

<% 
id=cint(trim(request("id")))
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from xw where id="&id ,cn,1,3
set rs1=server.CreateObject("adodb.recordset")
rs1.open"select * from xwlb where id="&rs("xwlb_id") ,cn,1,3



%><html>
<head><SCRIPT language=JavaScript>
var currentpos,timer;

function initialize()
{
timer=setInterval("scrollwindow()",50);
}
function sc(){
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos != document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</SCRIPT>
<title>深圳市德镒盟电子有限公司</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
a:link {
	text-decoration: none;
	color: #000000;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: underline;
	color: #FF0000;
}
a:active {
	text-decoration: none;
}
body {
	background-color: #333366;
	background-image: url(images/dEMLogo.jpg);
}
.xw1 {
	border-top: 4px solid #FFFFFF;
	border-right: 4px none #FFFFFF;
	border-bottom: 4px none #FFFFFF;
	border-left: 4px none #FFFFFF;
}
.style1 {color: #999999}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td colspan="3" align="center"><!--#include file="etop.asp"----></td>
  </tr>
  <tr>
    <td width="165" rowspan="2" align="center" valign="top"><!--#include file="exwleft.asp"----></td>
    <td width="3" rowspan="2" background="images/cpyyw_dot.gif">&nbsp;</td>
    <td width="549" align="center" valign="top"><table width="606" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="606"><TABLE width=602 border=0 align=center cellPadding=0 cellSpacing=0 class="xw1">
            <TBODY>
              <TR>
                <TD width=9 bgColor=#e7e7e7 height=25>&nbsp;</TD>
                <TD width=593 bgColor=#e7e7e7>·Position：<A 
            href="index.asp">Home</A> &gt; News</TD>
              </TR>
            </TBODY>
        </TABLE></td>
      </tr>
      <tr>
        <td height="106" valign="top"><TABLE cellSpacing=0 cellPadding=0 width=600 align=center border=0>
            <TBODY>
              <TR>
                <TD width="600"><IMG height=100 src="images/xwbt.jpg" 
            width=600 vspace=3></TD>
              </TR>
            </TBODY>
        </TABLE></td>
      </tr>
      <tr>
        <td height="61" valign="top"><table cellspacing="0">
          <tr>
            <td width="608" valign="top"><TABLE cellSpacing=0 cellPadding=0 width="95%" align=center 
                  border=0>
                <TBODY>
                  <TR>
                    <TD class=font_14_b align=middle colSpan=2 
                        height=13>&nbsp;</TD>
                  </TR>
                  <TR>
                    <TD class=news_info 
                      style="BORDER-TOP: #666666 1px solid; BORDER-BOTTOM: #666666 1px solid; color: #999999;" 
                      width="40%" height=30>双击自动滚屏</TD>
                    <TD class=news_info 
                      style="BORDER-TOP: #666666 1px solid; BORDER-BOTTOM: #666666 1px solid" 
                      align=middle width="60%"><span class="style1">发布者：管理员</span> <span class="style1">发布时间：<%=rs("tjsj")%></span> </TD>
                  </TR>
                  <TR align="center">
                    <TD colSpan=2><BR>
                        <DIV class=new_font_12_24 
                        style="FONT-SIZE: 10.5pt"><STRONG><FONT 
                        color=#5101ba><%=rs("enxwbt")%></FONT></STRONG>
                            <TABLE cellSpacing=0 cellPadding=0 width="80%" 
                          border=0>
                              <TBODY>
                                <TR>
                                  <TD><STRONG><FONT color=#5101ba><IMG height=17 
                              src="images/sanjidi1.gif" 
                              width=548></FONT></STRONG></TD>
                                </TR>
                              </TBODY>
                            </TABLE>
                            <BR>
                            <TABLE cellSpacing=0 cellPadding=0 width="548" 
                        align=center border=0>
                              <TBODY>
                                <TR>
                                  <TD height=478 align="left" valign="top" background="">
                                    <DIV align=left><%=rs("enxwnr")%></DIV></TD>
                                </TR>
                              </TBODY>
                            </TABLE>
                          
                            <DIV></DIV>
                      </DIV></TD>
                  </TR>
                  <TR align=right>
                    <TD colSpan=2>&nbsp;</TD>
                  </TR>
                </TBODY>
              </TABLE></td>
          </tr>
        </table>          
          </td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td align="center" valign="bottom"><table width="100%" height="14"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="14" background="images/qycz_03_bg.gif">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="bottom.asp"---></td>
  </tr>
</table>
</body>
</html>