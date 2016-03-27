<!--#include file="conn.asp"-->
<%
if request("send")="ok" then

	username=trim(request.form("username"))
	usermail=trim(request.form("usermail"))

	if username="" or usermail="" or request.form("Comments")="" then
	response.write "<script language='javascript'>"
	response.write "alert('必须填写用户名、邮箱、留言内容，请检查是否填写完整！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if Instr(username,"=")>0 or Instr(username,"%")>0 or Instr(username,chr(32))>0 or Instr(username,"?")>0 or Instr(username,";")>0 or Instr(username,",")>0 or Instr(username,"'")>0 or Instr(username,chr(34))>0 or Instr(username,chr(9))>0 or Instr(username,"")>0 or Instr(username,"$")>0 or Instr(username,">")>0 or Instr(username,"<")>0  then
	response.write "<script language='javascript'>"
	response.write "alert('您输入的用户名中含有非法字符，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	end if

	if Instr(usermail,".")<=0 or Instr(usermail,"@")<=0 or len(usermail)<10 or len(usermail)>50 then
	response.write "<script language='javascript'>"
	response.write "alert('您输入的电子邮件地址格式不正确，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	end if

	if len(request.form("Comments"))>200 then
	response.write "<script language='javascript'>"
	response.write "alert('留言内容太长了，请不要超过200个字符！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	set rs=Server.CreateObject("ADODB.RecordSet")
	sql="select * from Feedback where online='1' order by Postdate desc"
	rs.open sql,cn,1,3

			rs.Addnew
			rs("username")=Request("username")
			rs("comments")=Request("comments")
			rs("usermail")=Request("usermail")
			rs("face")=Request("face")
			rs("pic")=Request("pic")
			rs("url")=Request("url")
			rs("qq")=Request("qq")
			view=cstr(view)
			if view<>"0" then view="1"
			rs("online")=view
			rs("IP")=Request.serverVariables("REMOTE_ADDR")
			rs.Update
		rs.close
		set rs=nothing
	response.write "<script language='javascript'>"
	if view="0" then
	response.write "alert('留言提交成功，留言须经管理员审核才能发布。');"
	else
	response.write "alert('留言提交成功，单击“确定”返回留言列表！');"
	end if
	response.write "location.href='zxly_view.asp?dbid=d9';"	
	response.write "</script>"
	response.end
end if
maxlength=200
%><html>
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
.body{width:1050px;margin:0 auto;}
.context{width:760px;min-height:800px;margin:0 20px 50px 0; }
.right{width:260px;float:right;}
.toppic{text-align:center;padding-left:20px;}
-->
</style></head>
<body>
<!--#include file="top.asp"--->
<div class="body">
<div class="right"><!--#include file="rightsy.asp"---></div>
<div class="context">
<div class="toppic"><img src="images/toppic.png" alt=""></div>
<a href="zxly_view.asp?dbid=d9">&nbsp;</a>
<table cellSpacing="1" borderColorDark="#ffffff" cellPadding="4" width="642" align="center" bgColor="#D6C3A5" borderColorLight="#000000" border="0">
<form action=zxly.asp?lb=2 method=post name="book">
                        <tr >
                          <td   width="20%" align=right ><span class="STYLE1">您的姓名：</span></td>
                          <td ><span class="STYLE1">
                            <input type=text name="UserName" size="30" maxlength=16>
                              *</span></td>
                        </tr>
                        <tr >
                          <td   width="20%" align=right ><span class="STYLE1">您的邮箱：</span></td>
                          <td height="25" ><span class="STYLE1">
                            <input type=text name="UserMail" size="30"  maxlength=50>
                              *</span></td>
                        </tr>
                        <tr >
                          <td   width="20%" align=right ><span class="STYLE1">您的网站：</span></td>
                          <td height="25" ><input name="url" type=text value="http://" size="30"  maxlength=100></td>
                        </tr>
                        <tr >
                          <td   width="20%" align=right ><span class="STYLE1">其它联系方式：</span></td>
                          <td height="25" ><span class="STYLE1">
                            <input type=text value="" name="QQ" size="30"  maxlength=100>
                            （如QQ、MSN等）</span></td>
                        </tr>
                        <tr >
                          <td   width="20%" align=right><span class="STYLE1">留言内容：<br>
                              （<%=maxlength%>字以内）</span></td>
                          <td height="25" ><span style="color: #18455A">
                            <textarea name="Comments" rows="7" cols="50" style="overflow:auto;"></textarea>
                          </span></td>
                        </tr>
                        <tr >
                          <td   width="20%" align=right><span class="STYLE1">请选择表情：</span></td>
                          <td ><span class="STYLE1">
                            <input type="radio" value="1" name="face" checked>
                              <img border=0 src="images/face/face1.gif">
                              <input type="radio" value="2" name="face">
                              <img border=0 src="images/face/face2.gif">
                              <input type="radio" value="3" name="face">
                              <img border=0 src="images/face/face3.gif">
                              <input type="radio" value="4" name="face">
                              <img src="IMAGES/face/face4.gif" width="20" height="20" border=0>
                              <input type="radio" value="5" name="face">
                              <img border=0 src="images/face/face5.gif">
                              <input type="radio" value="6" name="face">
                              <img border=0 src="images/face/face6.gif">
                              <input type="radio" value="7" name="face">
                              <img border=0 src="images/face/face7.gif"> <br>
                              <input type="radio" value="8" name="face">
                              <img border=0 src="images/face/face8.gif">
                              <input type="radio" value="9" name="face">
                              <img border=0 src="images/face/face9.gif">
                              <input type="radio" value="10" name="face">
                              <img border=0 src="images/face/face10.gif">
                              <input type="radio" value="11" name="face">
                              <img border=0 src="images/face/face11.gif">
                              <input type="radio" value="12" name="face">
                              <img border=0 src="images/face/face12.gif">
                              <input type="radio" value="13" name="face">
                              <img border=0 src="images/face/face13.gif">
                              <input type="radio" value="14" name="face">
                              <img border=0 src="images/face/face14.gif"> <br>
                              <input type="radio" value="15" name="face">
                              <img border=0 src="images/face/face15.gif">
                              <input type="radio" value="16" name="face">
                              <img border=0 src="images/face/face16.gif">
                              <input type="radio" value="17" name="face">
                              <img border=0 src="images/face/face17.gif">
                              <input type="radio" value="18" name="face">
                              <img border=0 src="images/face/face18.gif">
                              <input type="radio" value="19" name="face">
                              <img border=0 src="images/face/face19.gif">
                              <input type="radio" value="20" name="face">
                              <img border=0 src="images/face/face20.gif"> </span></td>
                        </tr>
                        <tr >
                          <td width="20%" align=right ><span class="STYLE1">请选择头像：</span></td>
                          <td ><div align="center" class="STYLE1">
                              <input type="radio" value="1" name="pic" checked>
                              <img src="images/face/pic1.gif" width=40 border=0>
                              <input type="radio" value="2" name="pic">
                              <img border=0 src="images/face/pic2.gif" width=40>
                              <input type="radio" value="3" name="pic">
                              <img border=0 src="images/face/pic3.gif" width=40>
                              <input type="radio" value="4" name="pic">
                              <img border=0 src="images/face/pic4.gif" width=40>
                              <input type="radio" value="5" name="pic">
                              <img border=0 src="images/face/pic5.gif" width=40> <br>
                              <input type="radio" value="6" name="pic">
                              <img border=0 src="images/face/pic6.gif" width=40>
                              <input type="radio" value="7" name="pic">
                              <img border=0 src="images/face/pic7.gif" width=40>
                              <input type="radio" value="8" name="pic">
                              <img border=0 src="images/face/pic8.gif" width=40>
                              <input type="radio" value="9" name="pic">
                              <img border=0 src="images/face/pic9.gif" width=40>
                              <input type="radio" value="10" name="pic">
                              <img border=0 src="images/face/pic10.gif" width=40> </div></td>
                        </tr>
                        <tr bgColor="#ebebeb">
                          <td height="29" colSpan="2" ><div align="right">
                              <input type="submit" value="提交留言" name="Submit3">
                              <input type="reset" value="重新填写" name="Submit22">
                              <input type=hidden name=send value=ok>
                            &nbsp; </div></td>
                        </tr>
                      </form>
</table>
</div>



<!--#include file="bottom.asp"--->


<!-- ImageReady Slices (okk1.psd) -->
<!-- End ImageReady Slices -->
</body>
</html>
