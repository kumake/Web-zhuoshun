<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台管理页面:::菜单</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<SCRIPT src="Scripts/prototype.lite.js" type=text/javascript></SCRIPT>
	<SCRIPT src="Scripts/moo.fx.js" type=text/javascript></SCRIPT>
	<SCRIPT src="Scripts/moo.fx.pack.js" type=text/javascript></SCRIPT>
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<!--//-->
	<%
	'response.Write CK("usermenu")
	Dim menutype:menutype = VerificationUrlParam("menutype","string","")
	if menutype="" or menutype="help" then
		''显示帮助
	%>
		<DIV class=class>
			<DIV class=title-left>系统帮助</DIV>
			<DIV class=menu-left>
			<A href="manage.index.help.asp" target="main">后台系统帮助</A><BR>
			</DIV>
		</DIV>
	<DIV class=class>
		<DIV class=title-left>快捷菜单</DIV>
		<DIV class=menu-left>
		<A href="manage.Faster.add.asp" target="main">添加关键菜单</A><BR>
		<A href="manage.Faster.list.asp" target="main">菜单管理</A><BR>
		-----------------------------<BR>
		<%
		set rsf = Con.Query("select MenuName,MenuUrl from Sp_FasterMenu")
		if rsf.recordcount<>0 then
		do while not rsf.eof
			response.Write "<A href="""&rsf("MenuUrl")&""" target=""main"">"&rsf("MenuName")&"</A><BR>"
		rsf.movenext
		loop
		end if
		rsf.close
		set rsf = nothing
		%>
		
		</DIV>
	</DIV>
		<DIV class=class>
		<DIV class=title-left>站长工具</DIV>
		<DIV class=menu-left>
		<A href="manage.index.baidu.asp" target="main">百度收录查询</A><BR>
		<A href="manage.index.server.asp" target="main">服务器配置</A><BR>
		<A href="manage.index.baiduMap.asp" target="main">生成百度地图</A><BR>
		<A href="manage.index.googleMap.asp" target="main">生成Google地图</A><BR>
		<A href="manage.index.space.asp" target="main">空间容量查询</A><BR>
		</DIV>
	</DIV>
	<%
	elseif menutype="setting" then
	if instr(CK("usermenu"),"ModelSet")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>模型设置管理</DIV>
		<DIV class=menu-left>
		<A href="setting.Model.list.asp" target="main">模型管理</A><BR>
		<A href="setting.Model.add.asp" target=main>增加模型</A><BR>
		<A href="setting.Model.Field.list.asp" target=main>模型字段列表</A><BR>
		<A href="setting.Model.Field.add.asp" target=main>增加模型字段</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Template")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>模板管理</DIV>
		<DIV class=menu-left>
		<A href="setting.Model.template.asp" target="main">批量导入栏目模板</A><BR>
		<A href="setting.Model.template.asp" target="main">模板管理</A><BR>
		<A href="setting.Model.js.asp" target="main">Javascript模板管理</A><BR>
		<A href="#" onClick="javascript:alert('开发中...');" >自定义sql模板</A><BR>
		</DIV>	
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Lable")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>标签管理</DIV>
		<DIV class=menu-left>
		<A href="setting.Model.label.asp" target="main">标签管理</A><BR>
		<A href="setting.Model.label.add.asp" target="main">增加标签</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Area")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>地区管理</DIV>
		<DIV class=menu-left>
		<A href="setting.city.list.asp" target="main">地区管理</A><BR>
		<A href="setting.city.add.asp" target=main>增加地区</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"SetForm")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>自定义表单管理</DIV>
		<DIV class=menu-left>
			<A href="setting.selfForm.list.asp" target=main>自定义表单</A><BR>
			<A href="setting.selfForm.add.asp" target=main>增加自定义表单</A><BR>
			<A href="setting.selfForm.Field.list.asp" target=main>自定义表单字段</A><BR>
			<A href="setting.selfForm.Field.add.asp" target=main>增加表单字段</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"SetSinglePage")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>单页类别管理</DIV>
		<DIV class=menu-left>
		<A href="setting.singlePage.list.asp" target="main">类别管理</A><BR>
		<A href="setting.singlePage.add.asp" target=main>增加单页</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Dirctiary")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>数据字典</DIV>
		<DIV class=menu-left>
		<A href="setting.dictionary.Category.asp" target=main>数据字典类别</A><BR>
		<A href="setting.dictionary.Category.add.asp" target=main>增加字典类别</A><BR>
		<A href="setting.dictionary.asp" target=main>字典</A>&nbsp;<A href="setting.dictionary.add.asp" target="main" style="color:#FF0000;">增加字典</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"BaseConfig")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>全局设置</DIV>
		<DIV class=menu-left>
		<A href="setting.config.asp" target=main>全局系统设置</A><BR>
		<A href="setting.tables.clear.asp" target=main>一键清除表记录</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Caching")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>缓存设置</DIV>
		<DIV class=menu-left>
		<A href="setting.caching.asp" target=main>系统缓存清理</A><BR>
		<A href="setting.trace.asp" target=main>变量跟踪</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"ExpField")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>字段扩展</DIV>
		<DIV class=menu-left>
		<A href="setting.user.Field.expand.asp" target=main>用户字段扩展</A><BR>
		<A href="setting.sys.Field.expand.asp" target=main>表字段扩展</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"ManageRole")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>用户、角色及用户权限</DIV>
		<DIV class=menu-left>
		<A href="manage.Group.list.asp" target=main>管理员角色管理</A><BR>
		<A href="manage.Group.add.asp" target=main>增加角色</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"SystemLog")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>系统日志管理</DIV>
		<DIV class=menu-left>
			<A href="system.Logs.asp?category=login" target=main>登录日志</A><BR>
			<A href="system.Logs.asp?category=action" target=main>操作日志</A>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"WebSiteInterface")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>网站接口管理</DIV>
		<DIV class=menu-left>
			<A href="system.interface.asp?category=login" target=main>接口设置</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	elseif menutype="admin" then
		if instr(CK("usermenu"),"ManageManager")>0 then
	%>
		<DIV class=class>
			<DIV class=title-left>管理员管理</DIV>
			<DIV class=menu-left>
				<A href="manage.Group.User.list.asp" target=main>管理员管理</A><BR>
				<A href="manage.Group.User.add.asp" target=main>增加管理员</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"ManageUser")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>用户管理</DIV>
			<DIV class=menu-left>
				<A href="pulg.User.Group.asp" target=main>用户组管理</A><BR>
				<A href="pulg.User.Group.add.asp" target=main>增加用户组</A><BR>
				<A href="pulg.User.list.asp" target=main>用户管理</A><BR>
				<A href="pulg.User.add.asp" target=main>增加用户</A><BR>
				<A href="pulg.User.score.transaction.asp" target=main>用户账户交易记录</A><BR>
			</DIV>
		</DIV>
		<DIV class=class>
			<DIV class=title-left>点卡管理</DIV>
			<DIV class=menu-left>
				<A href="pulg.Card.List.asp" target=main>管理点卡</A><BR>
				<A href="pulg.Card.Create.asp" target=main>批量生成点卡</A><BR>
				<A href="pulg.Card.User.add.asp" target=main>批量赠送点数</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"MessageSend")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>短信&群发</DIV>
			<DIV class=menu-left>
				<A href="pulg.Msg.send.asp" target=main>群发短信</A><BR>
				<A href="pulg.email.send.asp" target=main>群发邮件</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"SelfForm")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>自定义表单信息管理</DIV>
			<DIV class=menu-left>
			<%
			set rs = Con.Query("select id,formname from sp_form")
			if rs.recordcount<>0 then
			do while not rs.eof
			%>
				<A href="pulg.selfForm.info.asp?formID=<%=rs("id")%>" target=main><%=rs("formname")%>信息</A><BR>
			<%
			rs.movenext
			loop
			end if
			%>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"EmailFeed")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>Email信息订阅管理</DIV>
			<DIV class=menu-left>
				<A href="pulg.email.list.asp" target=main>管理邮件地址</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"orderForm")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>产品订单管理</DIV>
			<DIV class=menu-left>
				<A href="pulg.order.list.asp" target=main>产品订单</A><BR>
				<A href="pulg.order.pay.list.asp" target=main>支付方式维护</A><BR>
				<A href="pulg.order.deliver.list.asp" target=main>配送方式维护</A><BR>
				<A href="pulg.order.cart.list.asp" target=main>购物车管理</A><BR>
				<A href="pulg.order.total.asp" style="color:#FF0000; text-decoration:underline;" target=main>产品销售统计</A><BR>
			</DIV>
		</DIV>
		<DIV class=class>
			<DIV class=title-left>在线支付</DIV>
			<DIV class=menu-left>
				<A href="pulg.order.list.asp" target=main>管理支付记录</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"Ads")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>广告管理系统</DIV>
			<DIV class=menu-left>
				<A href="pulg.ads.list.asp" target=main>广告管理</A><BR>
				<A href="pulg.ads.add.asp" target=main>增加广告</A><BR>
			</DIV>
		</DIV>
		<DIV class=class>
			<DIV class=title-left>轮显广告系统</DIV>
			<DIV class=menu-left>
				<A href="pulg.rollads.list.asp" target=main>广告管理</A><BR>
				<A href="pulg.rollads.add.asp" target=main>增加广告</A><BR>
			</DIV>
		</DIV>
		<DIV class=class>
			<DIV class=title-left>公告系统</DIV>
			<DIV class=menu-left>
				<A href="pulg.indexads.list.asp" target=main>公告管理</A><BR>
				<A href="pulg.indexads.add.asp" target=main>增加公告</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"Link")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>友情链接管理</DIV>
			<DIV class=menu-left>
				<A href="pulg.link.list.asp" target=main>友情链接管理</A><BR>
				<A href="pulg.link.add.asp" target=main>增加友情链接</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"Vote")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>在线调查管理</DIV>
			<DIV class=menu-left>
				<A href="pulg.vote.list.asp" target=main>调查管理</A><BR>
				<A href="pulg.vote.add.asp" target=main>增加调查</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"Job")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>人才招聘管理</DIV>
			<DIV class=menu-left>
				<A href="pulg.job.list.asp" target=main>岗位管理</A><BR>
				<A href="pulg.job.add.asp" target=main>增加岗位</A><BR>
				<A href="pulg.job.resume.asp" target=main>人才库管理</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"GuestBook")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>留言板管理</DIV>
			<DIV class=menu-left>
			<%
			set rs = Con.Query("select id,categoryname from Sp_dictionary where categoryID="&config_guestbookCategoryID&"")
			if rs.recordcount<>0 then
			do while not rs.eof
			%>
				<A href="pulg.guestbook.list.asp?guestbookID=<%=rs("id")%>" target="main"><%=rs("categoryname")%></A><BR>
			<%
			rs.movenext
			loop
			end if
			%>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"UploadFile")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>上传文件管理</DIV>
			<DIV class=menu-left>
				<A href="pulg.upload.asp" target=main>文件管理</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"DataBase")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>数据库管理</DIV>
			<DIV class=menu-left>
				<A href="plug.manage.database.sql.asp" target=main>sql高级管理</A><BR>
				<A href="plug.manage.database.back.asp" target=main>数据库备份还原</A><BR>
				<A href="plug.manage.database.upgrade.asp" target=main>数据库升级</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"BBSAdmin")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>论坛管理</DIV>
			<DIV class=menu-left>
				<A href="../Sp_Forum/admin/index.asp" target="_blank">论坛管理</A><BR>
			</DIV>
		</DIV>
	<%
		end if
	%>
		<DIV class=class>
			<DIV class=title-left>数据采集</DIV>
			<DIV class=menu-left>
				<A href="pulg.Gather.Add.asp" target="_blank">增加任务</A><BR>
				<A href="pulg.Gather.List.asp" target="_blank">管理任务</A><BR>
			</DIV>
		</DIV>
	<%
	elseif menutype="model" then
	'''''模型管理
		'response.Write CK("userrole")
		set rs = Con.Query("select * from Sp_Model where IsDisplay =1")
		if rs.recordcount<>0 then
		do while not rs.eof
		''''''''''''''''''
		if instr(CK("usermenu"),""&rs("modeltable")&"")>0 then
	%>
		<DIV class=class>
			<DIV class=title-left><%=rs("ModelName")%>管理</DIV>
			<DIV class=menu-left>
			    <%if rs("modelCategoryID")<>1 then%>
				<a href="setting.dictionary.asp?CategoryID=<%=rs("modelCategoryID")%>&location=Model" target=main>类别管理</a><BR>
				<%end if%>
				<A href="Model_Frame.asp?LeftUrl=Model_Mune.asp?ModelID=<%=rs("ID")%>&MainUrl=Model_list.asp?ModelID=<%=rs("ID")%>" target=main>信息管理</A><BR>
				<A href="Model_add.asp?ModelID=<%=rs("ID")%>" target=main>增加信息</A><BR>
				<A href="Model_list.asp?ModelID=<%=rs("ID")%>&action=Recycle" target=main>回收站信息管理</A><BR>
				<%if rs("IsModelNeedScore")=1 then%>
				<A href="Model_buy_list.asp?modelID=0&ParentModelID=<%=rs("ID")%>" target=main>信息购买记录</A><BR>
				<%end if%>
				<%if rs("ISAllowComment")=1 then%>
				<A href="Model_comment_list.asp?modelID=<%=rs("ID")%>" target=main>信息评论</A><BR>
				<%end if%>
			</DIV>
		</DIV>
	<%
		end if
		''''''''
		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
	
	if instr(CK("usermenu"),"SignleContent")>0 then
	%>
		<DIV class=class>
			<DIV class=title-left>单页管理</DIV>
			<DIV class=menu-left>
			<%
				function SinglePageCategoryMenu(SinglePageCategoryArray,parentID)
					for Sarraystep=0 to ubound(SinglePageCategoryArray,2) step 1
						if SinglePageCategoryArray(1,Sarraystep)=parentID and SinglePageCategoryArray(2,Sarraystep)<>"关于系统" then
							if parentID<>0 then 							
								response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"&SinglePageCategoryArray(0,Sarraystep)&".<A href='SinglePage_Add.asp?ModelID="&SinglePageCategoryArray(0,Sarraystep)&"' target=main>"
								response.Write SinglePageCategoryArray(2,Sarraystep)&"</a><BR>"
							else
								response.Write SinglePageCategoryArray(0,Sarraystep)&".<A href='SinglePage_Add.asp?ModelID="&SinglePageCategoryArray(0,Sarraystep)&"' target=main><u>"&SinglePageCategoryArray(2,Sarraystep)&"</u></a><BR>"
							end if
							SinglePageCategoryMenu SinglePageCategoryArray,SinglePageCategoryArray(0,Sarraystep)
						end if
					next
				end function
		
				Dim SinglePageCategoryArray:SinglePageCategoryArray = Con.QueryData("select ID,parentID,singlePagename from Sp_SinglePage")
				if Ubound(SinglePageCategoryArray,2)<>0 then SinglePageCategoryMenu SinglePageCategoryArray,0
			%>
			</DIV>
		</DIV>
	<%
	end if
	end if
	%>
</DIV>
<SCRIPT type=text/javascript>
    var contents = document.getElementsByClassName('menu-left');
    var toggles = document.getElementsByClassName('title-left');

    var myAccordion = new fx.Accordion(
        toggles, contents, {opacity: true, duration: 400}   
    );
    myAccordion.showThisHideOpen(contents[0]);  
</SCRIPT>
</BODY>
</HTML>
