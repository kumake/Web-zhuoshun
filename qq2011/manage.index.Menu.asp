<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨����ҳ��:::�˵�</TITLE>
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
		''��ʾ����
	%>
		<DIV class=class>
			<DIV class=title-left>ϵͳ����</DIV>
			<DIV class=menu-left>
			<A href="manage.index.help.asp" target="main">��̨ϵͳ����</A><BR>
			</DIV>
		</DIV>
	<DIV class=class>
		<DIV class=title-left>��ݲ˵�</DIV>
		<DIV class=menu-left>
		<A href="manage.Faster.add.asp" target="main">��ӹؼ��˵�</A><BR>
		<A href="manage.Faster.list.asp" target="main">�˵�����</A><BR>
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
		<DIV class=title-left>վ������</DIV>
		<DIV class=menu-left>
		<A href="manage.index.baidu.asp" target="main">�ٶ���¼��ѯ</A><BR>
		<A href="manage.index.server.asp" target="main">����������</A><BR>
		<A href="manage.index.baiduMap.asp" target="main">���ɰٶȵ�ͼ</A><BR>
		<A href="manage.index.googleMap.asp" target="main">����Google��ͼ</A><BR>
		<A href="manage.index.space.asp" target="main">�ռ�������ѯ</A><BR>
		</DIV>
	</DIV>
	<%
	elseif menutype="setting" then
	if instr(CK("usermenu"),"ModelSet")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>ģ�����ù���</DIV>
		<DIV class=menu-left>
		<A href="setting.Model.list.asp" target="main">ģ�͹���</A><BR>
		<A href="setting.Model.add.asp" target=main>����ģ��</A><BR>
		<A href="setting.Model.Field.list.asp" target=main>ģ���ֶ��б�</A><BR>
		<A href="setting.Model.Field.add.asp" target=main>����ģ���ֶ�</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Template")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>ģ�����</DIV>
		<DIV class=menu-left>
		<A href="setting.Model.template.asp" target="main">����������Ŀģ��</A><BR>
		<A href="setting.Model.template.asp" target="main">ģ�����</A><BR>
		<A href="setting.Model.js.asp" target="main">Javascriptģ�����</A><BR>
		<A href="#" onClick="javascript:alert('������...');" >�Զ���sqlģ��</A><BR>
		</DIV>	
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Lable")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>��ǩ����</DIV>
		<DIV class=menu-left>
		<A href="setting.Model.label.asp" target="main">��ǩ����</A><BR>
		<A href="setting.Model.label.add.asp" target="main">���ӱ�ǩ</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Area")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>��������</DIV>
		<DIV class=menu-left>
		<A href="setting.city.list.asp" target="main">��������</A><BR>
		<A href="setting.city.add.asp" target=main>���ӵ���</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"SetForm")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>�Զ��������</DIV>
		<DIV class=menu-left>
			<A href="setting.selfForm.list.asp" target=main>�Զ����</A><BR>
			<A href="setting.selfForm.add.asp" target=main>�����Զ����</A><BR>
			<A href="setting.selfForm.Field.list.asp" target=main>�Զ�����ֶ�</A><BR>
			<A href="setting.selfForm.Field.add.asp" target=main>���ӱ��ֶ�</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"SetSinglePage")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>��ҳ������</DIV>
		<DIV class=menu-left>
		<A href="setting.singlePage.list.asp" target="main">������</A><BR>
		<A href="setting.singlePage.add.asp" target=main>���ӵ�ҳ</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Dirctiary")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>�����ֵ�</DIV>
		<DIV class=menu-left>
		<A href="setting.dictionary.Category.asp" target=main>�����ֵ����</A><BR>
		<A href="setting.dictionary.Category.add.asp" target=main>�����ֵ����</A><BR>
		<A href="setting.dictionary.asp" target=main>�ֵ�</A>&nbsp;<A href="setting.dictionary.add.asp" target="main" style="color:#FF0000;">�����ֵ�</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"BaseConfig")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>ȫ������</DIV>
		<DIV class=menu-left>
		<A href="setting.config.asp" target=main>ȫ��ϵͳ����</A><BR>
		<A href="setting.tables.clear.asp" target=main>һ��������¼</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"Caching")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>��������</DIV>
		<DIV class=menu-left>
		<A href="setting.caching.asp" target=main>ϵͳ��������</A><BR>
		<A href="setting.trace.asp" target=main>��������</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"ExpField")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>�ֶ���չ</DIV>
		<DIV class=menu-left>
		<A href="setting.user.Field.expand.asp" target=main>�û��ֶ���չ</A><BR>
		<A href="setting.sys.Field.expand.asp" target=main>���ֶ���չ</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"ManageRole")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>�û�����ɫ���û�Ȩ��</DIV>
		<DIV class=menu-left>
		<A href="manage.Group.list.asp" target=main>����Ա��ɫ����</A><BR>
		<A href="manage.Group.add.asp" target=main>���ӽ�ɫ</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"SystemLog")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>ϵͳ��־����</DIV>
		<DIV class=menu-left>
			<A href="system.Logs.asp?category=login" target=main>��¼��־</A><BR>
			<A href="system.Logs.asp?category=action" target=main>������־</A>
		</DIV>
	</DIV>
	<%
	end if
	if instr(CK("usermenu"),"WebSiteInterface")>0 then
	%>
	<DIV class=class>
		<DIV class=title-left>��վ�ӿڹ���</DIV>
		<DIV class=menu-left>
			<A href="system.interface.asp?category=login" target=main>�ӿ�����</A><BR>
		</DIV>
	</DIV>
	<%
	end if
	elseif menutype="admin" then
		if instr(CK("usermenu"),"ManageManager")>0 then
	%>
		<DIV class=class>
			<DIV class=title-left>����Ա����</DIV>
			<DIV class=menu-left>
				<A href="manage.Group.User.list.asp" target=main>����Ա����</A><BR>
				<A href="manage.Group.User.add.asp" target=main>���ӹ���Ա</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"ManageUser")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>�û�����</DIV>
			<DIV class=menu-left>
				<A href="pulg.User.Group.asp" target=main>�û������</A><BR>
				<A href="pulg.User.Group.add.asp" target=main>�����û���</A><BR>
				<A href="pulg.User.list.asp" target=main>�û�����</A><BR>
				<A href="pulg.User.add.asp" target=main>�����û�</A><BR>
				<A href="pulg.User.score.transaction.asp" target=main>�û��˻����׼�¼</A><BR>
			</DIV>
		</DIV>
		<DIV class=class>
			<DIV class=title-left>�㿨����</DIV>
			<DIV class=menu-left>
				<A href="pulg.Card.List.asp" target=main>����㿨</A><BR>
				<A href="pulg.Card.Create.asp" target=main>�������ɵ㿨</A><BR>
				<A href="pulg.Card.User.add.asp" target=main>�������͵���</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"MessageSend")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>����&Ⱥ��</DIV>
			<DIV class=menu-left>
				<A href="pulg.Msg.send.asp" target=main>Ⱥ������</A><BR>
				<A href="pulg.email.send.asp" target=main>Ⱥ���ʼ�</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"SelfForm")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>�Զ������Ϣ����</DIV>
			<DIV class=menu-left>
			<%
			set rs = Con.Query("select id,formname from sp_form")
			if rs.recordcount<>0 then
			do while not rs.eof
			%>
				<A href="pulg.selfForm.info.asp?formID=<%=rs("id")%>" target=main><%=rs("formname")%>��Ϣ</A><BR>
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
			<DIV class=title-left>Email��Ϣ���Ĺ���</DIV>
			<DIV class=menu-left>
				<A href="pulg.email.list.asp" target=main>�����ʼ���ַ</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"orderForm")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>��Ʒ��������</DIV>
			<DIV class=menu-left>
				<A href="pulg.order.list.asp" target=main>��Ʒ����</A><BR>
				<A href="pulg.order.pay.list.asp" target=main>֧����ʽά��</A><BR>
				<A href="pulg.order.deliver.list.asp" target=main>���ͷ�ʽά��</A><BR>
				<A href="pulg.order.cart.list.asp" target=main>���ﳵ����</A><BR>
				<A href="pulg.order.total.asp" style="color:#FF0000; text-decoration:underline;" target=main>��Ʒ����ͳ��</A><BR>
			</DIV>
		</DIV>
		<DIV class=class>
			<DIV class=title-left>����֧��</DIV>
			<DIV class=menu-left>
				<A href="pulg.order.list.asp" target=main>����֧����¼</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"Ads")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>������ϵͳ</DIV>
			<DIV class=menu-left>
				<A href="pulg.ads.list.asp" target=main>������</A><BR>
				<A href="pulg.ads.add.asp" target=main>���ӹ��</A><BR>
			</DIV>
		</DIV>
		<DIV class=class>
			<DIV class=title-left>���Թ��ϵͳ</DIV>
			<DIV class=menu-left>
				<A href="pulg.rollads.list.asp" target=main>������</A><BR>
				<A href="pulg.rollads.add.asp" target=main>���ӹ��</A><BR>
			</DIV>
		</DIV>
		<DIV class=class>
			<DIV class=title-left>����ϵͳ</DIV>
			<DIV class=menu-left>
				<A href="pulg.indexads.list.asp" target=main>�������</A><BR>
				<A href="pulg.indexads.add.asp" target=main>���ӹ���</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"Link")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>�������ӹ���</DIV>
			<DIV class=menu-left>
				<A href="pulg.link.list.asp" target=main>�������ӹ���</A><BR>
				<A href="pulg.link.add.asp" target=main>������������</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"Vote")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>���ߵ������</DIV>
			<DIV class=menu-left>
				<A href="pulg.vote.list.asp" target=main>�������</A><BR>
				<A href="pulg.vote.add.asp" target=main>���ӵ���</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"Job")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>�˲���Ƹ����</DIV>
			<DIV class=menu-left>
				<A href="pulg.job.list.asp" target=main>��λ����</A><BR>
				<A href="pulg.job.add.asp" target=main>���Ӹ�λ</A><BR>
				<A href="pulg.job.resume.asp" target=main>�˲ſ����</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"GuestBook")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>���԰����</DIV>
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
			<DIV class=title-left>�ϴ��ļ�����</DIV>
			<DIV class=menu-left>
				<A href="pulg.upload.asp" target=main>�ļ�����</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"DataBase")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>���ݿ����</DIV>
			<DIV class=menu-left>
				<A href="plug.manage.database.sql.asp" target=main>sql�߼�����</A><BR>
				<A href="plug.manage.database.back.asp" target=main>���ݿⱸ�ݻ�ԭ</A><BR>
				<A href="plug.manage.database.upgrade.asp" target=main>���ݿ�����</A><BR>
			</DIV>
		</DIV>
		<%
		end if
		if instr(CK("usermenu"),"BBSAdmin")>0 then
		%>
		<DIV class=class>
			<DIV class=title-left>��̳����</DIV>
			<DIV class=menu-left>
				<A href="../Sp_Forum/admin/index.asp" target="_blank">��̳����</A><BR>
			</DIV>
		</DIV>
	<%
		end if
	%>
		<DIV class=class>
			<DIV class=title-left>���ݲɼ�</DIV>
			<DIV class=menu-left>
				<A href="pulg.Gather.Add.asp" target="_blank">��������</A><BR>
				<A href="pulg.Gather.List.asp" target="_blank">��������</A><BR>
			</DIV>
		</DIV>
	<%
	elseif menutype="model" then
	'''''ģ�͹���
		'response.Write CK("userrole")
		set rs = Con.Query("select * from Sp_Model where IsDisplay =1")
		if rs.recordcount<>0 then
		do while not rs.eof
		''''''''''''''''''
		if instr(CK("usermenu"),""&rs("modeltable")&"")>0 then
	%>
		<DIV class=class>
			<DIV class=title-left><%=rs("ModelName")%>����</DIV>
			<DIV class=menu-left>
			    <%if rs("modelCategoryID")<>1 then%>
				<a href="setting.dictionary.asp?CategoryID=<%=rs("modelCategoryID")%>&location=Model" target=main>������</a><BR>
				<%end if%>
				<A href="Model_Frame.asp?LeftUrl=Model_Mune.asp?ModelID=<%=rs("ID")%>&MainUrl=Model_list.asp?ModelID=<%=rs("ID")%>" target=main>��Ϣ����</A><BR>
				<A href="Model_add.asp?ModelID=<%=rs("ID")%>" target=main>������Ϣ</A><BR>
				<A href="Model_list.asp?ModelID=<%=rs("ID")%>&action=Recycle" target=main>����վ��Ϣ����</A><BR>
				<%if rs("IsModelNeedScore")=1 then%>
				<A href="Model_buy_list.asp?modelID=0&ParentModelID=<%=rs("ID")%>" target=main>��Ϣ�����¼</A><BR>
				<%end if%>
				<%if rs("ISAllowComment")=1 then%>
				<A href="Model_comment_list.asp?modelID=<%=rs("ID")%>" target=main>��Ϣ����</A><BR>
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
			<DIV class=title-left>��ҳ����</DIV>
			<DIV class=menu-left>
			<%
				function SinglePageCategoryMenu(SinglePageCategoryArray,parentID)
					for Sarraystep=0 to ubound(SinglePageCategoryArray,2) step 1
						if SinglePageCategoryArray(1,Sarraystep)=parentID and SinglePageCategoryArray(2,Sarraystep)<>"����ϵͳ" then
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
