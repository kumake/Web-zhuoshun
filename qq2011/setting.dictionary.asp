<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim location:location = VerificationUrlParam("location","string","Dictionary")
	Dim pathstr:pathstr = VerificationUrlParam("pathstr","string","")	
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","")
	Dim CategoryID:CategoryID = VerificationUrlParam("CategoryID","int","0")
	if action<>"" and action="del" then
		''删除信息
		Con.Execute("delete * from Sp_dictionary where ID="&ItemID&"")
		ALert "删除字典成功","setting.dictionary.asp?CategoryID="&CategoryID&"&location="&location&""
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">数据字典类别</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<div class="divpadding">
		        <%
				''''''
				if location<>"" and location = "Model" then
				else
				%>
				<select name="txt_categoryID" onChange="javascript:location.href='?CategoryID='+this.value;">
				<%
					set Rs = Con.Query("select id,dictionary from Sp_dictionaryCategory")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&""
					if Cint(CategoryID)=Cint(rs("id")) then
					response.Write " selected"
					end if
					response.Write ">"&rs("dictionary")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>
				<%
				'''''
				end if
				%>
		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="50%" align="left">
				</TH>
			    <TH width="30%" align="left">字典名</TH>
			    <TH width="15%" align="left"></TH>
			  </TR>
			  <%
				'''取数据存放进ModelCategoryArray
				Dim ModelCategoryArray:ModelCategoryArray = Con.QueryData("select D.ID,D.categoryID,D.categoryname,D.Pathstr,D.ParentID from Sp_dictionaryCategory C,Sp_dictionary D where C.id=D.categoryID order by D.id desc")
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strwhere:strwhere = "D.categoryID = C.ID"
				if CategoryID<>0 then strwhere = strwhere & " and D.categoryID="&categoryID&""
				if pathstr<>"" then strwhere = strwhere &" and D.Pathstr like '"&pathstr&"%'"
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_dictionaryCategory C,Sp_dictionary D where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("D.ID,D.categoryID,D.categoryname,D.Pathstr,D.parentID,C.dictionary","Sp_dictionaryCategory C,Sp_dictionary D", ""&strwhere&"", "D.ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><%ListDictionaryLeav ModelCategoryArray,rs("Pathstr")%></TD>
				<TD><%=rs("categoryname")%></TD>
				<TD>
				<%
				if rs("ID")<>1 then
				%>
				<a href="setting.dictionary.edit.asp?CategoryID=<%=rs("categoryID")%>&ItemID=<%=rs("ID")%>&location=<%=location%>">修改</a>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&CategoryID=<%=rs("CategoryID")%>&ItemID=<%=rs("ID")%>&location=<%=location%>';">删除</a>
				<%end if%>
				</TD>
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
		<div class="divpadding">
		  <input name="btnadd" value="增加类别字典" class="button" <%if CategoryID = 1 then response.Write "disabled"%> onClick="javascript:location.href='setting.dictionary.add.asp?CategoryID=<%=CategoryID%>&location=<%=location%>';" type="button">
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
