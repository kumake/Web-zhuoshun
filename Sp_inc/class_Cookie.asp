<%
Dim Ck
Set Ck = New HotCMS_Class_Cookie
Class HotCMS_Class_Cookie
	Public Ver, ClassName,Name
	
	Private Sub Class_Initialize()
		session.Timeout = 999
		Ver = "1.1"
		ClassName = "HotCMS_Class_Cookie"
		Name = Request.ServerVariables("APPL_MD_PATH")
		Name = Replace(Name,"/","_")
	End Sub

	Public Default Property Get D(vName)
		D = Request.Cookies(Name & "_" & vName)
	End Property

	Public Property Let D(vName, vValue)
		Response.Cookies(Name & "_" & vName) = vValue
	End Property
	
	Public Sub Clear()
		Dim elm
		for each elm in Request.Cookies
			If Left(elm, Len(Name)+1) = Name & "_" Then response.Cookies(elm) = ""
		Next
	End Sub
	
	Public Sub RemovePrefix(prefix)
		Dim elm
		for each elm in Request.Cookies
			If Left(elm, Len(Name)+Len(prefix)+1) = Name & "_" & prefix Then response.Cookies(elm) = ""
		Next
	End Sub
	
	Public Property Get Count(prefix)
		Dim elm
		Count = 0
		For Each elm in Request.Cookies
			If Left(elm, Len(Name)+Len(prefix)+1) = Name & "_" & prefix Then Count = Count+1
		Next
	End Property 
End Class
%>