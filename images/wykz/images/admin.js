<!--
function formcheck(_form,ty)
{
	try{
		if(_form.username.value=="")
		{
			alert("�û�������Ϊ�գ�");
			return false;
		}
	}catch(e){};
	if(ty==1)
		if(_form.password.value=="")
		{
			alert("���벻��Ϊ�գ�");
			return false;
		}
	if(_form.password.value!="" || _form.password2.value!="")
	{
		if(_form.password.value!=_form.password2.value)
		{
			if(ty==2)
				alert("�����������벻��ͬ�������������������ա�");
			else
				alert("�����������벻��ͬ��");
			return false;
		}
	}
	try{
		if(_form.adminpath.value=="")
		{
			alert("Ŀ¼����Ϊ�գ�");
			return false;
		}
	}catch(e){};
	_form.password.value=MD5(_form.password.value);
	_form.password2.value=MD5(_form.password2.value);
	try{_form.oldpassword.value=MD5(_form.oldpassword.value);}catch(e){};
}
//-->