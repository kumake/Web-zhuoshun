<%
	'Ĭ�ϻ�������
	'�ڶ����У�����ͨ�����Լ������������������������������Ĭ���˵ģ������޸�
	Dim An
	Set An = New HotCMS_Class_Application

	'��Ҫ˵��
	'��ҪΪ�˷����̣��Լ���ֹ��ͬϵͳ�ĻỰʶ��
	'�������⣬��ʱ��֧�ֶ���ı���
Class HotCMS_Class_Application
	Public Ver, ClassName	'��汾,����
	Public Name, ReloadTime
	
	'���ʼ��
	Private Sub Class_Initialize()
		Ver = "1.1"
		ClassName = "HotCMS_Class_Application"
		Name = Request.ServerVariables("APPL_MD_PATH")
		Name = Replace(Name,"/","_")
		ReloadTime = 1	'���ݻ���һ����
	End Sub

	'ȡ�û��档���������������ʽ�����Application�еģ����ݵ�һ��Ԫ�ؾ������ݣ��ڶ���Ԫ�ؾ��ǽ��������ʱ��
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
	
	'���滺��
	Public Property Let D(vName, vValue)
		Dim LocalValue(1)
		LocalValue(0) = vValue
		LocalValue(1) = Now()
		Application.Lock()
		Application(Name & "_" & vName) = LocalValue
		Application.UnLock()
	End Property
	
	'����������������л���
	Public Sub Clear()
		Dim i, elm
		'����ֵ��£�����Ҫ�����β�������ɾ������Լ���һ��ѭ��
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
	
	'�Ƴ�����ǰ׺Ϊprefix�Ļ���
	Public Sub RemovePrefix(prefix)
		'F.DebugStep "RemovePrefix(""" & prefix & """)"
		Dim i, elm
		'����ֵ��£�����Ҫ�����β�������ɾ������Լ���һ��ѭ��
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
	
	'����ǰ׺Ϊprefix�Ļ������
	Public Property Get Count(prefix)
		Dim elm
		Count = 0
		For Each elm in Application.Contents
			If Left(elm, Len(Name)+Len(prefix)+1) = Name & "_" & prefix Then Count = Count+1
		Next
	End Property 
End Class
%>