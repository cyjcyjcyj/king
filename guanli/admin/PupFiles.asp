<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="upload.inc" -->
<script language=javascript>
<!--
	function addcontent(tmpstr)
	{
		parent.document.form1.<%=Request("Position")%>.value = parent.document.form1.<%=Request("Position")%>.value + tmpstr;
	}

//-->
</script>
<%
Session.Timeout = 999
if Request("action")="up" then

	dim upload,file,formName,formPath,iCount,filename,fileExt
	set upload=new upload_5xSoft ''�����ϴ�����


	 formPath="../upadimg"

 ''��Ŀ¼���(/)
		if right(formPath,1)<>"/" then formPath=formPath&"/"

		iCount=0
			for each formName in upload.file ''�г������ϴ��˵��ļ�
			 set file=upload.file(formName)  ''����һ���ļ���
				 if file.filesize<100 then
				 	response.write "<font size=2>����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
					response.end
				 end if
 	
 				if file.filesize>1000*1024 then
				 	response.write "<font size=2>ͼƬ��С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
					response.end
				 end if

		 fileExt=lcase(right(file.filename,4))

			 if instr(".ram.rm.asf.wma.jpg.gif.jpeg.jpe.png.zip.rar.bmp.mp3.swf.wav.txt.doc.rtf.psd.tif.mid.tga.iff.pcx.dcx.pbm.pgm.ppm.pnm.miff.xbm.xpm.ico.icl.emf.hru.jif.prc.wrl.wbmp",fileExt)<1  then
				response.write "<font size=2>�ļ���ʽ���ԡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
				response.end
			 end if

		filename=Replace(Replace(Replace(now(),"-",""),":","")," ","")&FileExt
		savefile=formPath+filename
		 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
			  file.SaveAs Server.mappath(savefile)   ''������
		 end if
	 %>
  <script language="JavaScript">
 <%
 if fileExt=".gif" or fileExt=".jpg" or fileExt=".jpeg" or fileExt=".bmp"   or fileExt= ".png"  then
 %>
addcontent("<%="upadimg/"&filename%>")
<%Elseif  fileExt = ".swf" then
%>
addcontent("<%="uploadimg/"&filename%>\n")
<%
Elseif fileExt=".Mp3" or fileExt=".asf" or fileExt=".wav" or fileExt="mid" then
%>
addcontent("<%="upadimg/"&filename%>\n")
<%
else
%>
addcontent("[url]<%="upadimg/"&filename%>[/url]\n")
<%
end if
%>
</script>
 <%
	set file=nothing
	iCount=iCount+1
	next
	set upload=nothing  ''ɾ���˶���
	if err.number<>0 then
	Response.Write("�ļ��ϴ�����!�������Ա��ϵ:QQ(36279010)Email(yurix9209@sina.com)")
	Response.End()
	end if
	Response.Write("ͼƬ�ɹ��ϴ�")
	Response.End()
end if
%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <form name="form1" enctype="multipart/form-data" method="post" action="?action=up&Position=<%=Request("Position")%>"><tr>
      <td valign="top"><table width="375" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td> <input name="file" type="file" style="width:300px"> <input name="Submit" type="submit" class="bt" value="�ύ"> 
            </td>
          </tr>
        </table> </td>
  </tr></form>
</table>
