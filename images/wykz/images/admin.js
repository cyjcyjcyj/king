<!--
function formcheck(_form,ty)
{
	try{
		if(_form.username.value=="")
		{
			alert("用户名不能为空！");
			return false;
		}
	}catch(e){};
	if(ty==1)
		if(_form.password.value=="")
		{
			alert("密码不能为空！");
			return false;
		}
	if(_form.password.value!="" || _form.password2.value!="")
	{
		if(_form.password.value!=_form.password2.value)
		{
			if(ty==2)
				alert("两次输入密码不相同！若不更改密码请留空。");
			else
				alert("两次输入密码不相同！");
			return false;
		}
	}
	try{
		if(_form.adminpath.value=="")
		{
			alert("目录不能为空！");
			return false;
		}
	}catch(e){};
	_form.password.value=MD5(_form.password.value);
	_form.password2.value=MD5(_form.password2.value);
	try{_form.oldpassword.value=MD5(_form.oldpassword.value);}catch(e){};
}
//-->