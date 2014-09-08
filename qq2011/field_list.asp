<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">模型字段排列</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<div class="divpadding">
		信息检索：
		<select name="selcategory" style="width:100px;" class="select" onchange="javascript:if(this.value!='') location.href='?modelID='+this.value;">
		<option value="">选择模型</option>
		<%
			Dim modelID:modelID = VerificationUrlParam("modelID","int","0")
			set rs1 = Con.Query("select id,modelname from Sp_Model")
			if rs1.recordcount<>0 then
				do while not rs1.eof
		%>
		<option value="<%=rs1("id")%>" <%if Cint(modelID)=Cint(rs1("id")) then response.Write "selected"%>><%=rs1("modelname")%></option>
		<%
				rs1.movenext
				loop
			end if
		%>
		</select>
		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH align=left><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH align=left>模型名称</TH>
			    <TH align=left>字段名称</TH>
			    <TH align=left>字段类型</TH>
			    <TH align=left>字段页面描述</TH>
			    <TH align=left>是否必填</TH>
			    <TH align=left></TH>
		      </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Model M, Sp_ModelField F where M.ID=F.ModelID and M.ID="&modelID&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("F.ID,M.ID as MID","Sp_Model M,Sp_ModelField F", "M.ID=F.ModelID and M.ID="&modelID&"", "F.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD width="5%">&nbsp;</TD>
				<TD width="20%">&nbsp;</TD>
				<TD width="20%">&nbsp;</TD>
				<TD width="20%">&nbsp;</TD>
				<TD width="20%">&nbsp;</TD>
				<TD width="5%">&nbsp;</TD>
				<TD width="10%" align="center"><a href="field_add.asp?id=<%=rs("id")%>">修改</a>&nbsp;<a href="javascript:if(confirm('确定删除?删除后信息不可恢复!')) location.href='field_del.asp?action=del&id=<%=rs("id")%>';">删除</a></TD>
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
