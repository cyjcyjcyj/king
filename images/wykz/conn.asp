<%
dim conn,rs,sql,dbpath
set conn=server.createobject("adodb.connection")
dbpath="database/#9a3bd6ad57bea93b.mdb"		'���ݿ�·��
call conndate()

sub conndate()
	If Err Then err.Clear
	On Error Resume Next
	conn.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & server.mappath(dbpath)
	If Err Then
		err.Clear
		Set conn=Nothing
		Response.Write "<center><font color=red>��Ǹ�����ݿ����Ӵ�������conn.asp�ļ���</font></center>"
		Response.End
	End If
end sub

Sub closeconn()
    On Error Resume Next
    If IsObject(conn) Then
        conn.Close
        Set conn = Nothing
    End If
    If Err Then Err.Clear
End Sub
%>
