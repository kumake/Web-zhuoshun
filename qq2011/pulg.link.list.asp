<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","")
	if action<>"" and action="del" then
		''删除信息
		Con.Execute("delete * from Sp_FriendLink where ID="&ItemID&"")
		ALert "删除成功","pulg.link.list.asp"
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">友情链接管理</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="10%" align="left">类别名称</TH>
			    <TH width="15%" align="left">连接类型</TH>
			    <TH width="15%" align="left">站点名称</TH>
			    <TH width="20%" align="left">站点地址</TH>
			    <TH width="20%" align="left">logo</TH>
			    <TH width="5%" align="left">显示</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strwhere:strwhere = "F.linkCategory=D.id"
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_FriendLink F,Sp_dictionary D where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("F.*,D.Categoryname","Sp_FriendLink F,Sp_dictionary D", ""&strwhere&"", "F.ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><%=rs("Categoryname")%></TD>
				<TD><%if rs("urltype")=1 then:response.Write "logo":else:response.Write "文本":end if%></TD>
				<TD><%=rs("sitename")%></TD>
				<TD><%=rs("siteurl")%></TD>
				<TD><%if rs("urltype")=1 and rs("sitelogourl")<>"" then response.Write "<img src='"&rs("sitelogourl")&"' align='absmiddle'>"%></TD>
				<TD><%if rs("ISdisplay")=1 then:response.Write "是":else:response.Write "否":end if%></TD>
				<TD><a href="pulg.link.edit.asp?ItemID=<%=rs("ID")%>">修改</a>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&ItemID=<%=rs("ID")%>';">删除</a></TD>
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
		<div>
			<TABLE cellSpacing="1" width="100%">
			<TR class="content-td1">
				<TH align="left" colspan="2">代码调用</TH>
			</TR>
			<TR class="content-td1">
				<TD>asp</TD>
				<TD>
				<textarea name="textarea" cols="120" rows="2"><script src="http://<%=request.ServerVariables("HTTP_HOST")%>/plugIn/link.show.asp"></script></textarea>
				</TD>
			</TR>
			</TABLE>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
