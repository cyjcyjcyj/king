<!--
function checklogin()
{
	if (document.getElementById("TPL_username").value==""){
		alert("请输入用户名！");
		document.getElementById("TPL_username").focus();
		return false;
	}
	if (document.getElementById("TPL_password").value==""){
		alert("请输入密码！");
		document.getElementById("TPL_password").focus();
		return false;
	}
	if (document.getElementById("CheckCode").value==""){
		alert("请输入验证码！");
		document.getElementById("CheckCode").focus();
		return false;
	}
	document.getElementById("TPL_password").value=MD5(document.getElementById("TPL_password").value);
	document.getElementById("SubmitButton").disabled=true;
	return true;
}
function body_center()
{
	var _app=0;
	var strAppVersion = navigator.appVersion;
	if(strAppVersion.indexOf('MSIE') != -1) {_app=100;}
	var index_pos=document.getElementById("_body");
	var index_pos_top=(document.body.clientHeight/2-index_pos.clientHeight/2)-_app;
	if(index_pos_top<0) {index_pos_top=0;}
	index_pos.style.top=index_pos_top+"px";
	formchk();
}
function formchk()
{
	var cont=document.getElementById("Foo"+"terMSG").innerHTML;
	if((cont.indexOf(unescape("%u79CB%u5FC6"))==-1)||(cont.indexOf("sky"+"mean."+"com")==-1))
		document.getElementById("Cont"+"ent").innerHTML="Copy"+"right "+"Err"+"or!";
}
function setcode()
{
	if(document.getElementById("CheckCodeImg").innerHTML=="")
	{
		document.getElementById("CheckCode").style.width="80px";
		document.getElementById("CheckCodeMsg").innerHTML="Loading...";
		document.getElementById("CheckCodeMsg").style.display="";
		document.getElementById("CheckCodeImg").innerHTML='<a href="javascript:window.location.reload();" onfocus="if(this.blur) this.blur()"><img border=0 src="validatecode.asp" align="absmiddle" title="验证码1分钟失效。若失效请刷新页面。"></a>';
		document.getElementById("CheckCodeImg").style.display="";
		document.getElementById("CheckCode").title="";
	}
}
function InitInput()
{
	var usn = document.getElementById('TPL_username');
	if(usn)
	{
		usn.focus();
		usn.value = usn.value;
	}
}
//-->