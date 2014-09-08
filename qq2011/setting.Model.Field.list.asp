<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","")
	Dim isSubmodelID:isSubmodelID = VerificationUrlParam("isSubmodelID","int","-1")
	
	Dim ID:ID = VerificationUrlParam("ID","int","")
	if action<>"" and action="del" then
		Dim Fieldarray:Fieldarray = Con.QueryRow("select M.ModelTable,F.fieldname from Sp_Model M,Sp_ModelField F where F.ModelID=M.ID and F.ID="&ID&"",0)
		if ubound(Fieldarray)>0 then
			'删除表记录
			Con.Execute("delete * from Sp_ModelField where ID="&ID&"")
			'删除物理表结构 
			if isSubmodelID=0 then
			Con.Execute("alter table "&Fieldarray(0)&" DROP COLUMN "&Fieldarray(1)&"")
			if config_Verison = 2 then Con.Execute("alter table "&Fieldarray(0)&" DROP COLUMN "&Fieldarray(1)&"_en")
			end if
			Alert "删除字段成功","setting.Model.Field.list.asp?ModelID="&ModelID&""
		else
			Alert "数据查询失败",""
		end if		
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">模型字段列表</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<div class="divpadding">
		模型检索：
		<select name="ModelID" style="width:100px;" class="select" onChange="javascript:if(this.value!='') location.href='?ModelID='+this.value;">
		<option value="">选择模型</option>
		<%
			set Rs = Con.Query("select id,ModelName from SP_Model")
			if rs.recordcount<>0 then
			do while not rs.eof
			response.Write "<option value="&rs("id")&""
			if ModelID<>"" then
			if Cint(ModelID)=Cint(rs("ID")) then 
			response.Write " selected"
			end if
			end if
			response.Write ">"&rs("ModelName")&"</option>"
			rs.movenext
			loop
			end if
		%>
		</select>
		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH align="left">模型名称</TH>
			    <TH align="left">字段名称</TH>
			    <TH align="left">字段类型</TH>
			    <TH align="left">页面显示形式</TH>
			    <TH align="left">字段页面描述</TH>
			    <TH align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strWhere:strWhere = "F.ModelID=M.ID"				
				if ModelID <>"" then strWhere = strWhere & " and M.ID = "&ModelID&""
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Model M,Sp_ModelField F where "&strWhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("F.*,M.ModelName","Sp_Model M,Sp_ModelField F", ""&strWhere&"", "F.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD>&nbsp;</TD>
				<TD><%=rs("ModelName")%></TD>
				<TD><%=rs("fieldname")%></TD>
				<TD><%=rs("fieldtype")%></TD>
				<TD><%=rs("fieldUI")%></TD>
				<TD><%=rs("FieldDescription")%></TD>
				<TD><a href="setting.Model.Field.Edit.asp?ModelID=<%=rs("ModelID")%>&ID=<%=rs("ID")%>">修改</a>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&ModelID=<%=rs("ModelID")%>&ID=<%=rs("ID")%>&isSubmodelID=<%=rs("isSubmodelID")%>';">删除</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<input name="btnsearch" value="批量删除" class="button" type="button">
		</div>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
