<%
	Dim Sn
	Set Sn = New HotCMS_Class_Session
Class HotCMS_Class_Session
	Public Ver, ClassName
	Public Name
	
	Private Sub Class_Initialize()
		session.Timeout = 999
		Ver = "1.1"
		ClassName = "HotCMS_Class_Session"
		Name = Request.ServerVariables("APPL_MD_PATH")
		Name = Replace(Name,"/","_")
	End Sub

	Public Default Property Get D(vName)
		D = Session(Name & "_" & vName)
	End Property

	Public Property Let D(vName, vValue)
		Session(Name & "_" & vName) = vValue
	End Property
	
	Public Sub Clear()
		Dim elm
		For Each elm in Session.Contents
			If Left(elm, Len(Name)+1) = Name & "_" Then Session(elm) = Null
		Next
	End Sub
	
	Public Sub RemovePrefix(prefix)
		Dim elm
		For Each elm in Session.Contents
			If Left(elm, Len(Name)+Len(prefix)+1) = Name & "_" & prefix Then Session(elm) = Null
		Next
	End Sub
	
	Public Property Get Count(prefix)
		Dim elm
		Count = 0
		For Each elm in Session.Contents
			If Left(elm, Len(Name)+Len(prefix)+1) = Name & "_" & prefix Then Count = Count+1
		Next
	End Property 
End Class
%>