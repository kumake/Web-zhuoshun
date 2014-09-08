<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","")
	Dim JobID:JobID = VerificationUrlParam("JobID","int","")	
	if action<>"" and action="del" then
		''删除信息
		Con.Execute("delete * from Sp_GuestBook where ID="&ItemID&"")
		ALert "删除成功","plug.guestbook.list.asp"
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">人才库管理</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="15%" align="left">岗位名称</TH>
			    <TH width="15%" align="left">用户名称</TH>
			    <TH width="10%" align="left">性别</TH>
			    <TH width="25%" align="left">联系方式</TH>
			    <TH width="10%" align="left">状态</TH>
				<TH width="10%" align="left">提交日期</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strwhere:strwhere = "1=1"
				if JobID<>"" then strwhere = strwhere &" and JobID = "&JobID&"" 
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Resume where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("*","Sp_Resume", ""&strwhere&"", "ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><%=rs("jobname")%></TD>
				<TD><%=rs("username")%>&nbsp;<a href="plug.job.resume.view.asp?ItemID=<%=rs("ID")%>" class="huitext">详细</a></TD>
				<TD><%=rs("sex")%></TD>
				<TD><%=rs("tel")%>&nbsp;&nbsp;<%if rs("mobel")<>"" then response.Write rs("mobel")%></TD>
				<TD align="center">
				<%
				if rs("Ishire")=-1 then
				response.Write "淘汰"
				elseif rs("Ishire")=1 then
				response.Write "通知面世"
				elseif rs("Ishire")=2 then
				response.Write "正式录用"
				else
				response.Write "新加"
				end if
				%>
				</TD>
				<td><%=year(rs("posttime"))%>-<%=month(rs("posttime"))%>-<%=day(rs("posttime"))%></td>
				<TD><a href="plug.job.resume.view.asp?ItemID=<%=rs("id")%>">详细</a>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&ItemID=<%=rs("ID")%>';">删除</a></TD>
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
