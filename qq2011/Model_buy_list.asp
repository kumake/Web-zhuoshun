<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	''modelID=0&ParentModelID=57&ModelItemID=1
	'action
	Dim action:action = VerificationUrlParam("action","string","")
	''父模型ID
	Dim ParentModelID:ParentModelID = VerificationUrlParam("ParentModelID","int","0")
	'信息标识ID
	Dim ItemID:ItemID = VerificationUrlParam("ModelItemID","int","0")
	'模型名称
	Dim ModelName
	'存放该模型数据数据表
	Dim ModelTableName

	if ParentModelID=0 then
		response.Write "模型解释出现错误!"
	else
		'取对应模型的数据存放数据表
		Modelarray = Con.QueryRow("select ID,modelname,modeltable from Sp_Model where Id="&ParentModelID&"",0) 
		ModelName = Modelarray(1)		
		ModelTableName = Modelarray(2)	
	end if	
	'response.Write ModelTableName
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">模型主题信息用户购买记录</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="allCheck"></TH>
			    <TH width="65%" align="left">模型主题</TH>
			    <TH width="10%" align="left">用户名称</TH>
			    <TH width="10%" align="left">积分</TH>
			    <TH width="10%" align="left">交易金额</TH>
			  </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strwhere:strwhere = "M.ModelID=B.modelID and U.id=B.userUID and M.id=B.ModelItemID"
				if ParentModelID<>0 then strwhere =strwhere &" and B.modelID="&ParentModelID&""
				if ItemID<>0 then  strwhere =strwhere &" and B.ModelItemID="&ItemID&""
				Dim sql
				if total = 0 then 
					sql = "select count(*) as total from Sp_User U, Sp_ModelPurchaseLog B,"&ModelTableName&" M where "&strwhere&""
					'response.Write sql
					'response.End()
					Dim Tarray:Tarray = Con.QueryData(sql)
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				sql = Con.QueryPageByNum("U.username,M.title,B.*","Sp_User U, Sp_ModelPurchaseLog B,"&ModelTableName&" M", ""&strwhere&"", "B.ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><a href="Model_Edit.asp?ModelID=<%=rs("modelID")%>&id=<%=rs("ModelItemID")%>"><%=rs("title")%></a></TD>
				<TD><%=rs("username")%></TD>
				<TD><%=rs("score")%></TD>
				<TD><%=year(rs("buytime"))%>-<%=month(rs("buytime"))%>-<%=day(rs("buytime"))%></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
