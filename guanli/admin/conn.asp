<%
response.buffer=true '启用缓冲处理

dim conn 
dim connstr
'on error resume next
connstr="DBQ="+server.mappath("ad.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
If Err Then
        err.Clear
        Set Conn = Nothing
        Response.Write "数据库连接出错，请检查数据库连接文件中的数据库参数设置。"
        Response.End
    End If 
sub Chkhttp()
server_vv=len(Request.ServerVariables("SERVER_NAME"))
server_v1=left(Cstr(Request.ServerVariables("HTTP_REFERER")),server_vv)
server_v2=left(Cstr("http://"&Request.ServerVariables("SERVER_NAME")),server_vv)
if server_v1<>server_v2 or server_v1="" or server_v1="" then
response.write("<script>alert('错误：禁止从站点外部提交数据!.')</script>")
response.end
end if
end sub
%>