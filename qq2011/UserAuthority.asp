<%
	'''''''''�ж��û��Ƿ��ǺϷ���½�û�
	if CK("userid")<>"" and CK("username")<>"" and CK("userrole")<>"" then
	'''''
	else
		response.Write "<script>alert('Sp_CMSϵͳ��ʾ:\r\n\n\n��û�е�½!���ߵ�½�Ự���ݶ�ʧ!');parent.location.href='/admin/index.asp';</script>"
	end if
	
	''''''''
	'���ؼ�¼�����棬��ֹ�ظ���ѯ����˷ѡ�����������Ա�֤��ֻ���������ݿ������²���������
	Set UserAuthorityCache = Server.CreateObject("Scripting.Dictionary")
	Function UserAuthority(userpower)
		If IsObject(UserAuthorityCache) Then
			If UserAuthorityCache.Exists("UserAuthority") Then
				UserAuthority = UserAuthorityCache("UserAuthority")
				Exit Function
			End If
		End If
		
		Dim vRs
		Dim sql:sql = "select ModelCode,PowerAction from Sp_ManagePower where ID in("&userpower&")"
		UserAuthority = Con.QueryData(sql)		
		'��ӵ�����
		UserAuthorityCache.Add "UserAuthority", UserAuthority
	End Function

	''''''''�ж��û��Ƿ����Ȩ��
	'''useruid �û�UID
	'''CurrentModel ��ǰģ������
	'''CurrentAction ��ǰ��Ҫ��Ȩ��
	function CheckUserAuthority(CurrentModel,CurrentAction)	
		'''''����ҪȨ��
		if CurrentModel="" and CurrentAction="" then
		else
			Dim Authority_flag:Authority_flag = false
			Dim AuthorityArray:AuthorityArray = UserAuthority(CK("userpower"))
			for i=0 to ubound(AuthorityArray,2)
				if AuthorityArray(0,i)=CurrentModel and CurrentAction =AuthorityArray(1,i) then
					Authority_flag = true
					exit for
				end if
			next
			if Authority_flag=false then
				HtmlAlert "��û�и�Ȩ��#��½�Ự���ݶ�ʧ"
			end if
		end if
	end function	
%>