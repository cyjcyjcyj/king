<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->

<%response.buffer=true%>
<%
if request("deldata")<>"ɾ��ѡ����¼" then
myself=request.servervariables("PATH_INFO")

Set rs = Server.CreateObject("ADODB.Recordset")


sql="Select * From ad order by id desc"
rs.Open sql,Conn,3, 2
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>admin</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body marginwidth="0">

<table width="100" height="25" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"><a href="default.asp?Action=Addfiles">>�ϴ���ӹ��<</a></td>
  </tr>
</table>

<%
Action=Request("Action")
Select Case Action
	Case "Addfiles"
		Call Addfiles
	Case "Add"
		Call Add
	Case "edit"
		Call edit
	Case "upedit"
		Call upedit
End Select


Sub Add

	if request("title")="" then
	response.Write("<script>alert('���ⲻ��Ϊ�գ�');history.go(-1)</script>")
	Response.end
	end if 
	
	if request("imgpath")="" then
	response.Write("<script>alert('ͼƬ·������Ϊ�գ�');history.go(-1)</script>")
	Response.end
	end if
	 
	if request("link")="" then
	response.Write("<script>alert('���ӵ�ַ����Ϊ�գ�������������#!');history.go(-1)</script>")
	Response.end
	end if

sql="Insert into ad(Title,imgpath,link,[time]) Values('"&Request("Title")&"','"&Request("imgpath")&"','"&Request("link")&"','"&Request("time")&"') "
conn.execute sql
Response.Write("<script>alert('ͼƬ��ӳɹ���');location.href='default.asp'</script>")
End Sub




Sub upedit
		if request("title")="" then
		response.Write("<script>alert('���ⲻ��Ϊ�գ�');history.go(-1)</script>")
		Response.end
		end if
		if request("imgpath")="" then
		response.Write("<script>alert('ͼƬ·������Ϊ�գ�');history.go(-1)</script>")
		Response.end
		end if 
		if request("link")="" then
		response.Write("<script>alert('���ӵ�ַ����Ϊ�գ�������������#!');history.go(-1)</script>")
		Response.end
		end if


  Sql="Update ad Set Title='"&Request("Title")&"',imgpath='"&Request("imgpath")&"',link='"&Request("link")&"',[time]='"&Request("time")&"' where ID="&Request("ID")
 'response.Write(sql)
  Conn.execute sql
  Response.Write("<script>alert('�޸ĳɹ���');location.href='default.asp'</script>")
End Sub
%>

<% If rs.eof or rs.bof Then
			response.Write("û�м�¼")
			response.End()
			end if %>
<table width="660" border="0" align="center" cellpadding="0" cellspacing="0">
  <form action="default.asp" method="post" name="theform" >
    <tr> 
      <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="40" height="24" align="center"><strong>ѡ��</strong></td>
            <td width="20" align="center"><strong>ID</strong></td>
            <td align="center"><strong><font color="#CC0000">���ϴ����ͼƬ</font></strong></td>
            <td width="60" align="center"><strong>���ʱ��</strong></td>
            <td width="40" align="center"><strong>����</strong></td>
          </tr>
          <%
dim page
rs.pagesize=10

'��ǰҳ��
if not isempty(request("page")) then
page=request("page")
else
page=1
end if

page=CLng(page)

if page<rs.pagecount+1 then
rs.absolutepage=page
else
page=rs.pagecount
end if

for ipage=1 to rs.pagesize%>
          <tr> 
            <td height="23" align="center"> <span class="content"> 
              <INPUT type=checkbox value=<%=rs("id")%>
                  name=<%="del"&ipage%>>
              </span></td>
            <td height="1" align="center"><span class="content"><%=rs("id")%></span></td>
            <td height="23" align="center"><a href="../<%=rs("imgpath")%>" target="blank" title="<%=rs("title")%>"><%=rs("imgpath")%></a></td>
            <td align="center"><span class="arial"><%=rs("time")%></span></td>
            <td align="center"><a href="default.asp?Action=edit&ID=<%=Rs("ID")%>">�޸�</a></td>
          </tr>
          <%
rs.movenext
if rs.eof then exit for
next%>
          <tr> 
            <td height="40" colspan="5" class="arial"> <input name="deldata" type="submit" class="bt"  value="ɾ��ѡ����¼"> 
              <input type=hidden name=pagesize value=<%=rs.pagesize%>> 
              <%if page>1 then
				  response.write "<a href=""?page=1"">��ҳ</a>&nbsp;&nbsp;&nbsp;<a href=?page="&page-1&">��һҳ</a>&nbsp;&nbsp;&nbsp;"
				  end if
				  if page<rs.pagecount then
                  
				  response.write "<a href=?page="&page+1&">��һҳ</a>&nbsp;&nbsp;&nbsp;<a href=?page="&rs.pagecount&">βҳ</a>"
				  end if
				  response.write "&nbsp;&nbsp;&nbsp;"&page&"/"&rs.pagecount&""				  
				  
				  rs.close
set rs=nothing%>
            </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
<%

conn.close
else
dim delname,i,id

for i=1 to request("pagesize")
delname="del"&i

if request(delname)<>"" then
id=request(delname)

del_sql="delete from ad where id="&id&""
Conn.execute del_sql
end if
next
Conn.close
set Conn=nothing
response.redirect "default.asp"
end if
%>


<%Sub addfiles%>
<form name="form1" method="post" action="default.asp?Action=Add">
  <table height="25" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><strong><font color="#CC0000">�ϴ����ͼƬ</font></strong></td>
    </tr>
  </table>
  <table width="660" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>ͼƬ�ϴ���</td>
                </tr>
              </table></td>
            <td><iframe name="ad" frameborder=0 width=100% height=25 scrolling=no src=PupFiles.asp?Position=imgpath></iframe></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>ͼƬ·����</td>
                </tr>
              </table></td>
            <td><input name="imgpath" type="text" class="input" id="imgpath">
            </td>
            <td width="60">���ʱ�䣺 </td>
            <td width="80"> <input name="time" type="text" class="input" id="time" value="<%=date()%>"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>ͼƬ���⣺</td>
                </tr>
              </table></td>
            <td><input name="title" type="text" class="input" id="title"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>���ӵ�ַ��</td>
                </tr>
              </table></td>
            <td><input name="link" type="text" class="input" id="link" value="#"></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td align="center">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>&nbsp;</td>
                </tr>
              </table></td>
            <td><input name="Submit" type="submit" class="bt" value="���"></td>
          </tr>
        </table> 
        
      </td>
    </tr>
  </table>
</form>
<%End Sub%>

<%Sub edit%>
<%
Set Rs6= conn.execute("Select Title,imgpath,link,time from ad where ID="&Request("ID"))
%>
<form name="form1" method="post" action="?Action=upedit" onSubmit="return CheckForm()">
  <table height="25" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><strong><font color="#CC0000">�޸��ϴ����ͼƬ</font></strong></td>
    </tr>
  </table>
  <table width="660" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>ͼƬ�ϴ���</td>
                </tr>
              </table></td>
            <td><iframe name="ad" frameborder=0 width=100% height=25 scrolling=no src=PupFiles.asp?Position=imgpath></iframe></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>ͼƬ·����</td>
                </tr>
              </table></td>
            <td><input name="imgpath" type="text" class="input" value="<%=rs6(1)%>"> 
              <input name="ID" type="hidden" id="ID" value="<%=Request("ID")%>"> 
            </td>
            <td width="60">���ʱ�䣺 </td>
            <td width="80">
<input name="time" type="text" class="input" id="time" value="<%=rs6(3)%>"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>ͼƬ���⣺</td>
                </tr>
              </table></td>
            <td><input name="title" type="text" class="input" id="title" value="<%=rs6(0)%>"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>���ӵ�ַ��</td>
                </tr>
              </table></td>
            <td><input name="link" type="text" class="input" id="link" value="<%=rs6(2)%>"></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td align="center">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="35"> <table width="60" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td>&nbsp;</td>
                </tr>
              </table></td>
            <td><input name="Submit" type="submit" class="bt" id="Submit" value="�޸�" onclick="return CheckForm()"></td>
          </tr>
        </table> 
        
      </td>
    </tr>
  </table>
</form>
<%End Sub%>
</body>
</html>