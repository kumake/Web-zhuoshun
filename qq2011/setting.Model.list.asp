<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ExpModelID:ExpModelID = VerificationUrlParam("ExpModelID","int","-1")
	Dim id:id = VerificationUrlParam("id","int","0")
	if action<>"" and action="del" then
		''取模型对应的详细信息
		Dim ModelArray:ModelArray = Con.QueryRow("select * from Sp_Model where ID="&ID&"",0)
		Dim ModelTable:ModelTable = ModelArray(2)
		'response.Write ModelTable
		'response.End()
		''删除信息
		Con.Execute("delete * from Sp_Model where ID="&id&"")
		Con.Execute("delete * from Sp_ModelField where ModelID="&id&"")
		''更新关系模型的ID
		Con.Execute("update Sp_Model set ParentModelID=0 where ParentModelID="&id&"")
		Con.Execute("update Sp_Model set SubModelID=0 where SubModelID="&id&"")
		''删除物理表结构		
		if ExpModelID=0 then		
			Con.Execute("drop table "&ModelTable&"")
		end if
		''''删除模型用户组关系
		Con.Execute("delete * from Sp_ModelGroup where ModelID="&id&"")
		'''删除模型信息购买记录
		Con.Execute("delete * from Sp_ModelPurchaseLog where modelID="&id&"")
		''''删除模型对应的权限管理表中的分配记录
		Con.Execute("delete * from Sp_ManagePower where ModelID="&id&"")
		'''''
		ALert "删除模型成功","Setting.Model.List.asp"
	elseif action="clear" then
		Dim table:table = VerificationUrlParam("table","string","")
		if submodelID=0 then		
		Con.Execute("delete * from "&table&"")
		'''删除模型信息购买记录
		Con.Execute("delete * from Sp_ModelPurchaseLog where modelID="&id&"")
		end if
		ALert "清空表成功","Setting.Model.List.asp"
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">模型列表</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="AllCheck"></TH>
			    <TH width="25%" align="left">模型名称</TH>
			    <TH width="25%" align="left">模型表名称</TH>
			    <TH width="15%" align="left">模型类别对应数据字典</TH>
			    <TH width="7%" align="left">是否显示</TH>
			    <TH width="23%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Model M,Sp_dictionaryCategory D where D.ID=M.modelCategoryID")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("M.*,D.dictionary","Sp_Model M,Sp_dictionaryCategory D", "D.ID=M.modelCategoryID", "M.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><%=rs("ModelName")%></TD>
				<TD><%=rs("ModelTable")%></TD>
				<TD><%=rs("dictionary")%></TD>
				<TD><%if rs("IsDisplay")=0 then:response.Write "隐藏":else:response.Write "显示":end if%></TD>
				<TD><a href="setting.Model.Field.list.asp?ModelID=<%=rs("ID")%>">字段</a>&nbsp;<a href="setting.Model.edit.asp?ModelID=<%=rs("ID")%>">修改</a>&nbsp;<%if rs("issystem")=0 then%><a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&id=<%=rs("ID")%>&ExpModelID=<%=rs("ExpModelID")%>';">删除</a><%end if%>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=clear&table=<%=rs("Modeltable")%>&ExpModelID=<%=rs("ExpModelID")%>&id=<%=rs("ID")%>';">清空表</a>&nbsp;<%if rs("ExpModelID")=0 then response.Write "<a href='ExpandModel.asp?ModelID="&rs("id")&"'><span style='color:#999;'>从此扩展</span></a>"%></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,false,false%>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
