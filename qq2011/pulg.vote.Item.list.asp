<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim voteID:voteID = VerificationUrlParam("voteID","int","0")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="del" then
		''删除信息
		Con.Execute("delete * from Sp_VoteItem where id="&ItemID&"")		
		ALert "删除成功","pulg.vote.Item.list.asp?voteID="&voteID&""
	elseif action="save" then
		Dim voteItem:voteItem =F.ReplaceFormText("txt_voteItem")
		Dim voteNum:voteNum =F.ReplaceFormText("txt_voteNum")
		Dim voteMemo:voteMemo =F.ReplaceFormText("txt_voteMemo")
		Con.Execute("insert into Sp_VoteItem (VoteID,voteItem,voteNum,voteMemo) values ("&VoteID&",'"&voteItem&"',"&voteNum&",'"&voteMemo&"')")		
		ALert "增加成功","pulg.vote.Item.list.asp?voteID="&voteID&""
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
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_voteItem").value=="")
			{
				alert("Sp_CMS提示\r\n\n子项名称必须填写!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">调查子项管理</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent1',this)" href="javascript:void(0)">增加子项</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent2',this)" href="javascript:void(0)">投票图示</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="40%" align="left">子项名称</TH>
			    <TH width="15%" align="left">子项备注</TH>
			    <TH width="15%" align="left">子项票数</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strwhere:strwhere = "voteID = "&voteID&""
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_VoteItem where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("*","Sp_VoteItem", ""&strwhere&"", "ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><%=rs("voteItem")%></TD>
				<TD><%=rs("votememo")%></TD>
				<TD><%=rs("voteNum")%></TD>
				<TD><a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&VoteID=<%=voteID%>&ItemID=<%=rs("ID")%>';">删除</a></TD>
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
	<DIV class="tagContent selectTag content" id="tagContent1" style="display:none;">
		<!--第1个标签//-->
		<form action="?action=save&VoteID=<%=VoteID%>" method="post" onSubmit="javascript:return checkform();">
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">子项名称：</TD>
				<TD><input name="txt_voteItem" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>子项默认票数：</TD>
				<TD><input name="txt_voteNum" class="input" type="text" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>备注：</TD>
				<TD><textarea name="txt_voteMemo" cols="40" rows="5" class="input"></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="增加子项" class="button" type="submit">
		</div>
		</form>
	</DIV>
	<DIV class="tagContent selectTag content" id="tagContent2" style="display:none;">
		<DIV>
		新功能开发中...
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
