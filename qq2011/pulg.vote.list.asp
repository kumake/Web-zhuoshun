<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","")
	if action<>"" and action="del" then
		''删除信息
		Con.Execute("delete * from Sp_Vote where ID="&ItemID&"")
		Con.Execute("delete * from Sp_VoteItem where voteID="&ItemID&"")		
		ALert "删除成功","pulg.vote.list.asp"
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">调查管理</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="40%" align="left">调查名称</TH>
			    <TH width="15%" align="left">开始时间</TH>
			    <TH width="15%" align="left">结束时间</TH>
			    <TH width="5%" align="left">是否单选</TH>
			    <TH width="5%" align="left">显示</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strwhere:strwhere = "1=1"
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Vote where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("*","Sp_Vote", ""&strwhere&"", "ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><a href="pulg.vote.Item.list.asp?voteID=<%=rs("ID")%>"><%=rs("votename")%></a>&nbsp;<a href="pulg.vote.Item.list.asp?voteID=<%=rs("ID")%>" class="huitext">子项</a></TD>
				<TD><%=year(rs("validstarttime"))%>-<%=month(rs("validstarttime"))%>-<%=day(rs("validstarttime"))%></TD>
				<TD><%=year(rs("validendtime"))%>-<%=month(rs("validendtime"))%>-<%=day(rs("validendtime"))%></TD>
				<TD><%if rs("IsSingle")=1 then:response.Write "单选":else:response.Write "复选":end if%></TD>
				<TD><%if rs("isactive")=1 then:response.Write "是":else:response.Write "否":end if%></TD>
				<TD><a href="pulg.vote.edit.asp?ItemID=<%=rs("ID")%>">修改</a>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&ItemID=<%=rs("ID")%>';">删除</a>&nbsp;<a href="pulg.vote.Item.list.asp?voteID=<%=rs("ID")%>">详细</a></TD>
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
		<br>
		<span class="huitext">说明:在要显示投票的地方加上:&lt;script src=调用地址&gt;&lt;/script&gt;</span>
</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
