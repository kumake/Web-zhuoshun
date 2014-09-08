<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim tableName:tableName = VerificationUrlParam("tableName","string","")
	Dim fieldName:fieldName = VerificationUrlParam("fieldName","string","")
	Dim fieldtype:fieldtype = VerificationUrlParam("fieldtype","string","")
	Dim fieldlength:fieldlength = VerificationUrlParam("fieldLen","int","4")
	if action<>"" and action="add" then
		'''增加物理表结构
		Dim fieldSql
		Dim fieldSql_En
		select case fieldtype
			case "char","nvarchar"
				fieldSql = "alter table "&tableName&" add "&fieldName&" "&fieldtype&"("&fieldlength&")"
				fieldSql_En = "alter table "&tableName&" add "&fieldName&"_en "&fieldtype&"("&fieldlength&")"
			case "int"
				fieldSql = "alter table "&tableName&" add "&fieldName&" "&fieldtype&""
				fieldSql_En = "alter table "&tableName&" add "&fieldName&"_en "&fieldtype&""
			case "ntext"
				fieldSql = "alter table "&tableName&" add "&fieldName&" "&fieldtype&""
				fieldSql_En = "alter table "&tableName&" add "&fieldName&"_en "&fieldtype&""
			case "datetime"
				fieldSql = "alter table "&tableName&" add "&fieldName&" "&fieldtype&""
				fieldSql_En = "alter table "&tableName&" add "&fieldName&"_en "&fieldtype&""
			case else 
				fieldSql = "alter table "&tableName&" add "&fieldName&" "&fieldtype&"(255)"
				fieldSql_En = "alter table "&tableName&" add "&fieldName&"_en "&fieldtype&"(255)"
		end select	  
		Con.execute(fieldSql)
		'Con.execute(fieldSql_En)
		''''
		Alert "操作成功","plug.manage.database.upgrade.asp"
	elseif action="del" then
			'删除物理表结构 
			Con.Execute("alter table "&tableName&" DROP COLUMN "&fieldName&"")
			'Con.Execute("alter table "&tableName&" DROP COLUMN "&fieldName&"_en")
		Alert "操作成功","plug.manage.database.upgrade.asp"
	end if
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">数据库升级</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<div style="font-size:12px; color:#999999; padding-bottom:8px;">红色的表名表示最新系统包含字段</div>
		<DIV style="float:left; width:400px;">
			<TABLE cellSpacing="1" width="100%">
				<%
				upgrateStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath("../upgrade/upgrate.mdb") &";Jet OLEDB:Database Password=sp_cms_manage;"
				set UICon = Server.CreateObject("ADODB.Connection")
				UICon.open upgrateStr
				Set rstSchema = UICOn.OpenSchema(20)
				Do Until rstSchema.EOF '遍历数据库表
				if rstSchema("TABLE_TYPE")="TABLE" then 
				%>			  
			  <Tr>
				<Th colspan="5" style="color:#FF0000;"><%=rstSchema("TABLE_NAME")%></TD>
			  </Tr>
			  <%
			  	'遍历表中字段
				set rs=UICon.execute("select * from "& rstSchema("TABLE_NAME"))
				for i=0 to rs.fields.count-1
			  %>
			  <TR class="content-td2">
				<TD width="13%"></TD>
				<TD width="26%"><%=rs(i).name%></TD>
				<TD width="20%"><%=rs(i).type%></TD>
				<TD><%=rs(i).DefinedSize%></TD>
				<TD><a href="?action=add&tableName=<%=rstSchema("TABLE_NAME")%>&fieldName=<%=rs(i).name%>&fieldtype=<%=VerAccessType(rs(i).type)%>&fieldLen=<%=rs(i).DefinedSize%>" style="color:#999999; font-size:10px;">增加</a></TD>
			  </TR>
			  <%
				next 
			  %>
			<%
			    end	if 
				rstSchema.MoveNext
				Loop 
			%>
			</TABLE>
		</DIV>
		<DIV style="float:left; width:100px;"></DIV>
		<DIV style="float:left; width:400px;">
			<TABLE cellSpacing="1" width="100%">
				<%
				Str = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath("../Sp_data/SP_cms.mdb") &";Jet OLEDB:Database Password=sp_cms_manage;"
				set UICon = Server.CreateObject("ADODB.Connection")
				UICon.open Str
				Set rstSchema = UICOn.OpenSchema(20)
				Do Until rstSchema.EOF '遍历数据库表
				if rstSchema("TABLE_TYPE")="TABLE" then 
					
				%>			  
			  <Tr>
				<Th colspan="5"><%=rstSchema("TABLE_NAME")%></TD>
			  </Tr>
			  <%
			  	'遍历表中字段
				set rs=UICon.execute("select * from "& rstSchema("TABLE_NAME"))
				for i=0 to rs.fields.count-1
			  %>
			  <TR class="content-td2">
				<TD width="13%"></TD>
				<TD width="26%"><%=rs(i).name%></TD>
				<TD width="20%"><%=rs(i).type%></TD>
				<TD><%=rs(i).DefinedSize%></TD>
				<TD><a href="?action=del&tableName=<%=rstSchema("TABLE_NAME")%>&fieldName=<%=rs(i).name%>&fieldtype=<%=VerAccessType(rs(i).type)%>&fieldLen=<%=rs(i).DefinedSize%>" style="color:#999999; font-size:10px;">删除</a></TD>
			  </TR>
			  <%
				next 
			  %>
			<%
			    end	if 
				rstSchema.MoveNext
				Loop 
			%>
			</TABLE>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
