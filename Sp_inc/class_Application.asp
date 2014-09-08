<%
	'默认基本对象
	'在二代中，各个通用类自己负责变量声明，这个类对象名，都是默认了的，不许修改
	Dim An
	Set An = New HotCMS_Class_Application

	'简要说明
	'主要为了方便编程，以及防止不同系统的会话识别
	'技术问题，暂时不支持对象的保存
Class HotCMS_Class_Application
	Public Ver, ClassName	'类版本,名称
	Public Name, ReloadTime
	
	'类初始化
	Private Sub Class_Initialize()
		Ver = "1.1"
		ClassName = "HotCMS_Class_Application"
		Name = Request.ServerVariables("APPL_MD_PATH")
		Name = Replace(Name,"/","_")
		ReloadTime = 1	'数据缓存一分钟
	End Sub

	'取得缓存。缓存是以数组的形式存放于Application中的，数据第一个元素就是数据，第二个元素就是建立缓存的时间
	Public Default Property Get D(vName)
		Dim Key, LocalValue
		Key = Name & "_" & vName
		LocalValue = Application(Key)
		If IsArray(LocalValue) Then
			If IsDate(LocalValue(1)) Then
				If DateDiff("s", CDate(LocalValue(1)),Now()) > (60 * ReloadTime) Then
					Application.Lock()
					Application.Contents.Remove(Key)
					Application.UnLock()
					Exit Property
				End If
				D = LocalValue(0)
			End If
		End If
	End Property
	
	'保存缓存
	Public Property Let D(vName, vValue)
		Dim LocalValue(1)
		LocalValue(0) = vValue
		LocalValue(1) = Now()
		Application.Lock()
		Application(Name & "_" & vName) = LocalValue
		Application.UnLock()
	End Property
	
	'清除本对象管理的所有缓存
	Public Sub Clear()
		Dim i, elm
		'很奇怪的事，发现要清楚多次才能清除干净，所以加了一个循环
		For i = 1 To 5
			For Each elm In Application.Contents
				If Left(elm, Len(Name)+1) = Name & "_" Then
					Application.Lock()
					Application.Contents.Remove(elm)
					Application.UnLock()
				End If
			Next
		Next
	End Sub
	
	'移除所有前缀为prefix的缓存
	Public Sub RemovePrefix(prefix)
		'F.DebugStep "RemovePrefix(""" & prefix & """)"
		Dim i, elm
		'很奇怪的事，发现要清楚多次才能清除干净，所以加了一个循环
		For i = 1 To 5
			For Each elm in Application.Contents
				If Left(elm, Len(Name)+Len(prefix)+1) = Name & "_" & prefix Then
					Application.Lock()
					Application.Contents.Remove(elm)
					Application.UnLock()
				End If
			Next
		Next
	End Sub
	
	'返回前缀为prefix的缓存个数
	Public Property Get Count(prefix)
		Dim elm
		Count = 0
		For Each elm in Application.Contents
			If Left(elm, Len(Name)+Len(prefix)+1) = Name & "_" & prefix Then Count = Count+1
		Next
	End Property 
End Class
%>