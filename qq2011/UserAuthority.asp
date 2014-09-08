<%
	'''''''''判断用户是否是合法登陆用户
	if CK("userid")<>"" and CK("username")<>"" and CK("userrole")<>"" then
	'''''
	else
		response.Write "<script>alert('Sp_CMS系统提示:\r\n\n\n你没有登陆!或者登陆会话数据丢失!');parent.location.href='/admin/index.asp';</script>"
	end if
	
	''''''''
	'本地记录集缓存，防止重复查询造成浪费。放在这里，可以保证在只有连接数据库的情况下才声明对象
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
		'添加到缓存
		UserAuthorityCache.Add "UserAuthority", UserAuthority
	End Function

	''''''''判断用户是否具有权限
	'''useruid 用户UID
	'''CurrentModel 当前模块名称
	'''CurrentAction 当前需要的权限
	function CheckUserAuthority(CurrentModel,CurrentAction)	
		'''''不需要权限
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
				HtmlAlert "你没有该权限#登陆会话数据丢失"
			end if
		end if
	end function	
%>