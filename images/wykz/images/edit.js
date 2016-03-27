<!--
function useado()
{
	if(document.filepos.adodb.checked){
		$("toolbarright").style.display="";
		$("filepath").style.width="340px";
	}else{
		$("toolbarright").style.display="none";
		$("filepath").style.width="500px";
	}
}
function closeabout(o)
{
	if(o==null) $("aboutmain").style.display="none";
	else $("aboutmain").style.display="";
	return false;
}
function hiddencopy()
{
	if($("aboutmain").style.display=="")
		if((event.keyCode==13||event.keyCode==32)||event.keyCode==27){event.returnValue=false;closeabout();}
}
function savefile()
{
	var Form=document.savefile;
	var charcode=$("charset").value;
	if(charcode!="")
	{
		if(!window.confirm("当前文件以"+charcode+"编码读取，要以此编码保存文件吗？"))
		{
			$("charset").value="";
			$("useado").value="";
		}
	}
	if($("filepath").value!="") {Form.path.value=$("filepath").value;}
	Form.submit();
}
function store(filename)
{
	filename=escape(filename);
	if(GetCookie("System_Edit_Store"+filename)==null)
		SetCookie("System_Edit_Store"+filename,document.savefile.fcontent.value);
}
function recover(filename)
{
	if(window.confirm("确定要还原文件["+filename+"]最初始内容吗？"))
	{
		try{
			filename=escape(filename);
			if(GetCookie("System_Edit_Store"+filename)!=null)
			{
				document.savefile.fcontent.value=GetCookie("System_Edit_Store"+filename);
			}
			else
			{
				document.savefile.reset();
			}
			$("content_change").innerHTML="&nbsp;*";
		}catch(e){};
	}
}
//-->