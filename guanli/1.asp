<%xwlb_id=2
   response.write "<script language='javascript'>"
	response.write "alert('操作失败，没有选择合适参数，请单击“确定”返回！');"
	response.write "location.href='xwlb.asp?xwlb_id="&xwlb_id&"';"
	response.write "</script>"
	response.end
	%>
