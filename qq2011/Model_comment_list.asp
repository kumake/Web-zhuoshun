<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	'action
	Dim action:action = VerificationUrlParam("action","string","")
	'模型ID
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","0")
	'信息标识ID
	Dim ModelItemID:ModelItemID = VerificationUrlParam("ModelItemID","int","0")
	'评论信息标识ID
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	'存放该模型数据数据表
	Dim ModelTableName
	'模型名称
	Dim ModelName
	if action="del" and ItemID<>0 then
		Con.Execute("delete * from Sp_Comment where ID="&ItemID&"")
		Alert "删除成功","Model_comment_list.asp?modelID="&ModelID&"&ModelItemID="&ModelItemID&""
	elseif action="batchdel" then
		Dim BatchrestoreItemID:BatchrestoreItemID = F.ReplaceFormText("ItemUID")
		Con.Execute("delete * from Sp_Comment where ID in ("&BatchrestoreItemID&")")
		'''''''
		WriteLog CK("username"),"action","批量回收站模型 "&ModelName&" 内容"
		'''''''
		Alert "批量删除成功","Model_comment_list.asp?modelID="&ModelID&"&ModelItemID="&ModelItemID&""
	end if
	if ModelID=0 then
		Alter "模型解释出现错误!",""
	else
		'取对应模型的数据存放数据表
		Modelarray = Con.QueryRow("select ID,modelname,modeltable from Sp_Model where Id="&ModelID&"",0) 
		ModelName = Modelarray(1)		
		ModelTableName = Modelarray(2)		
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript">
		//////批量操作信息
		function SelectALLCheck(obj)
		{
			for (var i=0;i<document.forms[0].elements.length;i++) 
			{ 
				var e = document.forms[0].elements[i]; 
				if (e.name != obj.name && e.disabled==false && e.type=="checkbox") 
				{
					e.checked = obj.checked;
				}				
			} 			
		}
		/////
		function checkedItem()
		{
			var Item = 0;
			for (var i=0;i<document.forms[0].elements.length;i++) 
			{ 
				var e = document.forms[0].elements[i]; 
				if (e.name == "ItemUID" && e.disabled==false && e.type=="checkbox") 
				{
					if(e.checked)
					{
						Item = Item + 1
					}
				}				
			} 
			return Item;
		}
		//////
		function batchAction()
		{
			if(checkedItem()==0)
			{
				alert("Sp_CMS系统提示:\r\n\n请选择要批量操作的信息!");
				return false;
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?ModelID=<%=ModelID%>&ModelItemID=<%=ModelItemID%>&action=batchdel" name="ModelItemList" onSubmit="javascript:return batchAction();" method="post">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">评论管理</A></LI>
	  <LI><A href="model_list.asp?ModelID=<%=ModelID%>">返回<%=ModelName%></A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="2%" align=left><input type="checkbox" name="allCheck" onclick="javascript:SelectALLCheck(this);"></TH>
			    <TH align=left width="20%" >标题</TH>
			    <TH align=left width="10%" >类型</TH>
			    <TH align=left width="10%" >主题/联系人</TH>
			    <TH align=left width="30%" >评论内容/联系方式</TH>
			    <TH width="8%" align=left>发布人</TH>
			    <TH width="10%" align=left>更新时间</TH>
			    <TH width="5%" align=left></TH>
		      </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strWhere		'查询条件
				strWhere = " C.ModelID = "&ModelID&" and C.ModelItemID = MT.ID"
				if ModelItemID<>0 then strWhere = strWhere &" and C.ModelItemID="&ModelItemID&" "
				'response.Write strWhere
				'response.End()
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from "&ModelTableName&" MT, Sp_Comment C where "&strWhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("C.ID as ItemID,C.content,C.title as Ctitle,C.username,C.Ip,C.posttime,C.CommentTYpe,Mt.ID,Mt.title",""&ModelTableName&" MT, Sp_Comment C", ""&strWhere&"", "C.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemUID" value="<%=rs("ItemID")%>"></TD>
				<TD><a href="Model_Edit.asp?ModelID=<%=ModelID%>&id=<%=rs("id")%>"><%=rs("title")%></a></TD>
				<TD><%if rs("CommentTYpe")="C" then:response.Write "评论":else:response.Write "求职应聘":end if%></TD>
				<TD><%=rs("Ctitle")%></TD>
				<TD><%=rs("content")%></TD>
				<TD><%=rs("username")%></TD>
				<TD><%=Year(rs("postTime"))%>-<%=month(rs("postTime"))%>-<%=day(rs("postTime"))%></TD>
				<TD align="center"><a href="javascript:if(confirm('确定删除?删除后信息不可恢复!')) location.href='Model_comment_list.asp?modelID=<%=ModelID%>&ModelItemID=<%=ModelItemID%>&ItemID=<%=rs("ItemID")%>&action=del';">删除</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<input name="btnrestore" value="批量删除" class="button" type="submit">		
		</div>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
