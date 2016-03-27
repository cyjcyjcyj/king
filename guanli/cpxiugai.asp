<!--#include file="lianjie.asp"-->
<% id=cint(trim(request("id")))
set rscp=server.CreateObject("adodb.recordset")
rscp.open"select * from cp where id="&id,cn,1,3
on error resume next
set rsyjfl=server.CreateObject("adodb.recordset")
rsyjfl.open"select * from dlb where id="&rscp("dlb_id"),cn,1,3
set rsejfl=server.CreateObject("adodb.recordset")
 
rsejfl.open"select * from xlb where id="&rscp("xlb_id"),cn,1,3

set rs1=server.CreateObject("adodb.recordset")
set rs2=server.CreateObject("adodb.recordset")
rs1.open"select id ,dlb  from dlb ",cn,1,3
rs2.open"select id ,xlb,dlb_id from xlb",cn,1,3
%>
<title>管理页面</title><SCRIPT language = "JavaScript">
<!--//
    var onecount;
    onecount=0;
    subcat = new Array();
  <%
i=0
do while not rs2.eof

%>      
    subcat[<%=i%>] = new Array("<%=rs2("xlb")%>","<%=rs2("dlb_id")%>","<%=rs2("id")%>");

<%
rs2.movenext
i=i+1
loop
%>
    onecount=<%=i%>;

function changelocation(locationid)
    {
    document.myform.classid.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.classid.options[document.myform.classid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    } 

//-->  
    </SCRIPT>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
.style1 {	color: #DBBF90;
	font-weight: bold;
}
-->
</style></head>

<body>
<form name="myform" method="post" action="cp_xiugai.asp?id=<%=id%>">
  <table class="tableBorder" width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">产品资料</span></td>
    </tr>
    <tr>
      <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
          <tr bgcolor="#E8F1FF">
            <td width="20%" align="right">教师名称：</td>
            <td width="80%" style="PADDING-LEFT: 10px"><input name="cpmc" type="text" id="cpmc" value="<%=rscp("cpmc")%>" size="50"></td>
          </tr>
          <tr bgcolor="#E8F1FF">
            <td align="right" valign="middle">教师随笔：</td>
            <td style="PADDING-LEFT: 10px"><textarea name="cptpjqsm" style="display:none" cols="1" rows="1"><%=rscp("cptpjqsm")%></textarea>
                <IFRAME ID="eWebEditor1" SRC="edit/ewebeditor.htm?id=cptpjqsm&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME></td>
          </tr>
          <tr bgcolor="#E8F1FF">
            <td align="right" valign="top">教师照片：</td>
            <td style="PADDING-LEFT: 10px"><input name="tpdz" type="text" id="tpdz" value="<%=rscp("tpdz")%>" size="50">
                <input name="Submit1" type="button" onClick="window.open('upfile.asp?formname=myform&editname=tpdz&uppath=images/guanggao&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="上 传"></td>
          </tr>
          <tr bgcolor="#E8F1FF">
            <td height="30"></td>
            <td height="30" style="PADDING-LEFT: 10px">
              <input type="submit" name="Submit" value="修改">
              <input type="reset" name="Submit2" value="重置">
            </td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
</body>
</html>
