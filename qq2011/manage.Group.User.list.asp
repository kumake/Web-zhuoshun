<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim id:id = VerificationUrlParam("id","int","0")
	if action<>"" and action="del" then
		Con.Execute("delete * from Sp_Manage where ID="&id&"")
		ALert "删除成功","manage.Group.User.list.asp"
	end if	
	
	dim RoleArray:RoleArray = Con.QueryData("select ID,RoleName from Sp_ManageRole")
	function listUserRole(userrole)
		dim temprole:temprole = split(trim(userrole),",")
		for i=0 to ubound(temprole)
			for j=0 to ubound(RoleArray,2)
				if Cint(RoleArray(0,j))=Cint(temprole(i)) then 
				response.Write "<a href='manage.Group.edit.asp?action=view&ID="&RoleArray(0,j)&"'>"&RoleArray(1,j)&"</a>&nbsp;&nbsp;&nbsp;"
				exit for
				end if
			next
		next
	end function
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A>管理员管理</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="10%" align="left">管理员名称</TH>
			    <TH width="70%" align="left">所属角色</TH>
			    <TH width="15%" align="left"></TH>
			  </TR>
			  <%
			  '''response.Write CK("userrole")
			  '''
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Manage where ParentID="&CK("userid")&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("*","Sp_Manage", "ParentID="&CK("userid")&"", "ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD>&nbsp;</TD>
				<TD><%=rs("UserName")%></TD>
				<TD>
				<%listUserRole rs("UserRole")%>
				</TD>
				<TD><a href="manage.Group.User.edit.asp?ID=<%=rs("ID")%>">修改</a>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&id=<%=rs("ID")%>';">删除</a></TD>
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
