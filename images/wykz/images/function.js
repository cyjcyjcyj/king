<!--
function SetEditItem(_name,_value)
{
	try{
		var edit=$("edit_item");
		edit.innerHTML=edit.innerHTML+'<input type="hidden" name="'+_name+'" value="'+_value+'" />';
	}catch(e){};
}

function picsize(obj,MaxWidth){
	img=new Image();
	img.src=obj.src;
	if (img.width>MaxWidth)
	{
		return MaxWidth;
	}
	else
	{
		return img.width;
	}
}

function confirm_edit(filename,encode)
{
	if(encode!=null)
	{
		if(window.confirm("��ȷ���ļ� ["+filename+"] ΪASCII�����ļ������༭����"))
		{
			window.open("edit.asp?path="+PathEncode+encode,"","");
		}
	}
	else
	{
		alert("�������ļ�Ϊ��ASCII�ļ����޷��༭��");
	}
}

function SaveRename(SaveID)
{
	var _opt;
	if($("f_file").value=="renfile") {_opt="file";}
	else {_opt="folder";}
	
	if(SaveID!=null && $("f_file").value!="")
	{
		if($("_"+_opt+"_org"+SaveID).value!=$("_"+_opt+"_new"+SaveID).value)
		{
			if(window.confirm("Ҫ��������ļ�����"))
			{
				SetEditItem("act",$("f_file").value);
				SetEditItem("path",path);
				SetEditItem("currentpath",path);
				SetEditItem("oldname",$("_"+_opt+"_org"+SaveID).value);
				SetEditItem("newname",$("_"+_opt+"_new"+SaveID).value);
				SetEditItem("page",page);
				document.cmd_aff.action="cmd.asp";
				document.cmd_aff.submit();
				return null;
			}
		}
	}
	if($("f_file").value!="")
	{
		var f_file_v=$("f_file_v").value;
		$("_"+_opt+"_new"+SaveID).value=$("_"+_opt+"_org"+SaveID).value;
		$("_"+_opt+f_file_v).style.display="";
		$("_"+_opt+"_new"+f_file_v).style.display="none";
		$("_"+_opt+"_opt1"+f_file_v).style.display="";
		$("_"+_opt+"_opt2"+f_file_v).style.display="none";
		$("f_file").value="";
		$("f_file_v").value="";
	}
}

function rename(oldname,act,itemid)
{
	var _opt;
	var strAppVersion = navigator.appVersion;
	SaveRename();
	if(act=="renfile") {_opt="file";}
	else {_opt="folder";}
	$("_"+_opt+itemid).style.display="none";
	$("_"+_opt+"_new"+itemid).style.display="";
	$("_"+_opt+"_opt1"+itemid).style.display="none";
	$("_"+_opt+"_opt2"+itemid).style.display="";
	if(strAppVersion.indexOf('MSIE') != -1)
	{
		$("_"+_opt+"_new"+itemid).select();
	}
	else
	{
		$("_"+_opt+"_new"+itemid).focus();
	}
	$("f_file").value=act;
	$("f_file_v").value=itemid;
}
function runsave()
{
	var _opt;
	if($("f_file").value=="renfile") {_opt="file";}
	else {_opt="folder";}
	try{$("_"+_opt+"_new"+$("f_file_v").value).blur();}
	catch(e){};
}
function cancelrename(){SaveRename();}
function explorerchk()
{
	var cont=$("foo"+"ter").innerHTML;
	if((cont.indexOf(unescape("%u79CB%u5FC6"))==-1)||(cont.indexOf("sky"+"mean."+"com")==-1))
		$("bodymain").innerHTML="Copy"+"right "+"Err"+"or!";
}
function delfile(act,filename)
{
	var addchar;
	if(act=="del")
	{
		addchar="�ļ�";
	}
	else
	{
		addchar="�ļ���";
	}
	if(window.confirm("�����Ҫɾ��"+addchar+" ["+filename+"] ��"))
	{
		SetEditItem("act",act);
		SetEditItem("path",path+filename);
		SetEditItem("currentpath",path);
		SetEditItem("page",page);
		document.cmd_aff.action="cmd.asp";
		document.cmd_aff.submit();
	}
	//window.location.href="cmd.asp?act="+act+"&path="+path+filename+"&currentpath="+PathEncode+"&page="+page;
}
function selectall(objname,act)
{
	if(objname==null)
	{
		alert("���󣺵�ǰĿ¼��û���ļ����У���ѡ��/���");
	}
	else if(objname.length==null)
	{
		objname.checked=act;
	}
	else
	{
		var n=objname.length;
		for (i=0;i<n;i++)
		{
			objname[i].checked=act;
		}
	}
}

function create_file(act,filename)
{
	var addchar;
	if(act=="copycon") {addchar="�ļ�";}
	else {addchar="�ļ���";}
	
	if(filename!="" && filename!=null)
	{
		SetEditItem("act",act);
		SetEditItem("path",path+filename);
		SetEditItem("currentpath",path);
		SetEditItem("page",page);
		document.cmd_aff.action="cmd.asp";
		document.cmd_aff.submit();
	}
	//window.location.href="cmd.asp?act="+act+"&path="+PathEncode+filename+"&currentpath="+PathEncode+"&page="+page;
	if(filename=="") alert("����û������"+addchar+"�����ƣ�");
}

function copy(act)
{
	with(document){
		explorer.action="cmd.asp?act="+act+"&path="+PathEncode+"&page="+page;
		explorer.act.value=act;
		explorer.submit();
	}
}
function JS_URLEncode(str)
{
	var temp=str;
	while(temp.indexOf("#")!=-1)
	{
		temp=temp.replace("#","%23")
	}
	while(temp.indexOf(" ")!=-1)
	{
		temp=temp.replace(" ","%20")
	}
	return temp;
}
function remind()
{
	explorerchk();
	if(LoginOneIP!="1"){return false;}
	if(GetCookie("System_Exit_Remind")!="yes")
	{
		SetCookie("System_Exit_Remind","yes");
		alert("������ʾ��\n\n��ϵͳ�����û�ͬһʱ��ֻ����һ��IP��½��\n\n������Ϻ�ǵõ����Ͻǵġ�ע����½���� ^_^");
	}
}
function sfolder(foldername,showname,encode)
{
	if(showname=="") {showname=foldername;}
	var HtmlStr='\
		<tr onmouseover="this.bgColor=\'#eeeeee\'" onmouseout="this.bgColor=\'\'">\
		<td align="left" width="176" style="height:22px;">\
		<input type="checkbox" id="_folder_org'+item_count+'" name="folders" value="'+foldername+'"><img align="absmiddle" src="images/icon/folder.gif"><span id="_folder'+item_count+'"><a href="'+selfname+'?path='+PathEncode+encode+'">'+showname+'</a></span><input type="text" id="_folder_new'+item_count+'" value="'+foldername+'" onblur="SaveRename('+item_count+')" class="_rename" style="display:none" onkeypress="if(event.keyCode==13){this.blur();return false;}"></td>\
		<td align="right" width="54"><span id="_folder_opt1'+item_count+'"><a href="javascript:delfile(\'rd\',\''+foldername+'\')">ɾ��</a> <a href="javascript:rename(\''+foldername+'\',\'renfolder\',\''+item_count+'\')">����</a></span>\
		<span id="_folder_opt2'+item_count+'" style="display:none"><a href="javascript:runsave();">ȷ��</a> <a href="javascript:cancelrename();">ȡ��</a></span>\
		</td>\
		</tr>\
	';
	document.write(HtmlStr);
	item_count++;
}
function files(filename,showname,encode,icon,filesize,filetime,canedit)
{
	if(showname=="") {showname=filename;}
	var HtmlStr='\
		<tr onmouseover="this.bgColor=\'#eeeeee\'" onmouseout="this.bgColor=\'\'">\
		<td align="left" height="22" width="245" class="filetdr"><input type="checkbox" id="_file_org'+item_count+'" name="files" value="'+filename+'"><img align="absmiddle" src="images/icon/'+icon+'" onload="this.width=picsize(this,18);"><span id="_file'+item_count+'"><a target="_blank" href="'+JS_URLEncode(path)+encode+'">'+showname+'</a></span><input type="text" id="_file_new'+item_count+'" value="'+filename+'" onblur="SaveRename('+item_count+')" class="_rename" style="width:200px;display:none" onkeypress="if(event.keyCode==13){this.blur();return false;}"></td>\
		<td align="left" width="62" class="filetdr">'+filesize+'</td>\
		<td align="left" width="115" style="color:#4E2727" class="filetdr">'+filetime+'</td>\
		<td align="center" width="118">\
		<span id="_file_opt1'+item_count+'">';
	if(canedit==1)
	{
		HtmlStr+='<a target="_blank" href="edit.asp?path='+PathEncode+encode+'">�༭</a>\n';
	}
	else if(canedit==0)
	{
		HtmlStr+='<a href="javascript:confirm_edit()">�༭</a>\n';
	}
	else
	{
		HtmlStr+='<a href="javascript:confirm_edit(\''+filename+'\',\''+encode+'\')">�༭</a>\n';
	}
	HtmlStr+='\
		<a href="javascript:delfile(\'del\',\''+filename+'\')">ɾ��</a>\
		<a href="javascript:rename(\''+filename+'\',\'renfile\',\''+item_count+'\')">����</a>\
		<a target="_blank" href="download.asp?path='+PathEncode+encode+'">����</a></span>\
		<span id="_file_opt2'+item_count+'" style="display:none"><a href="javascript:runsave();">ȷ��</a>\n<a href="javascript:cancelrename();">ȡ��</a></span></td>\
		</tr>\
	';
	document.write(HtmlStr);
	item_count++;
}
//-->