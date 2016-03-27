<!--#include file="config.asp"-->
<%
NumCode (8)
Function NumCode(CodeType)
    Response.Expires = -1
    Response.AddHeader "Pragma", "no-cache"
    Response.AddHeader "cache-ctrol", "no-cache"
    On Error Resume Next
    Dim zNum, rNum, i, j, listnum, listcode
    Dim Ados, Ados1
    listcode = "0123456789abcdefghijklmnopqrstuvwxyz"
    Randomize Timer
    Dim zimg(6), NStr
    For i = 0 To 5
        rNum = CStr(CInt(35 * Rnd))	'将35改为9即为使用纯数字密码
        zimg(i) = rNum
        listnum = listnum & Mid(listcode, rNum + 1, 1)
    Next
    if len(trim(request("t")))<1 then
	Session.TimeOut=1
    else
	if IsNumeric(trim(request("t"))) then
		Session.TimeOut=int(trim(request("t")))
	else
		Session.TimeOut=1
	end if
    end if
    if len(trim(request("v")))<1 then
	Session(SessionPrefix&"syscode")=listnum
    else
	Session(SessionPrefix&trim(request("v")))=listnum
    end if
    Dim Pos
    Set Ados = Server.CreateObject("Adodb.Stream")
    Ados.Mode = 3
    Ados.Type = 1
    Ados.Open
    Set Ados1 = Server.CreateObject("Adodb.Stream")
    Ados1.Mode = 3
    Ados1.Type = 1
    Ados1.Open
    Ados.LoadFromFile (Server.mappath("images/body.fix"))
    Ados1.write Ados.read(2880)
    For i = 0 To 5
        Ados.Position = (35 - zimg(i)) * 480
        Ados1.Position = i * 480
        Ados1.write Ados.read(480)
    Next
    Ados.LoadFromFile (Server.mappath("images/head.fix"))
    Pos = LenB(Ados.read())
    Ados.Position = Pos
    For i = 0 To 15 Step 1
        For j = 0 To 5
            Ados1.Position = i * 32 + j * 480
            Ados.Position = Pos + 30 * j + i * 270
            Ados.write Ados1.read(30)
        Next
    Next
    Response.ContentType = "image/BMP"
    Ados.Position = 0
    Response.BinaryWrite Ados.read()
    Ados.Close: Set Ados = Nothing
    Ados1.Close: Set Ados1 = Nothing
    'If Err Then Session(SessionPrefix&"CheckCode") = "999999"
End Function
%>