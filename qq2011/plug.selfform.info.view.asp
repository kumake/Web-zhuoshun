<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	'模型ID
	Dim FormID:FormID = VerificationUrlParam("FormID","int","0")
	Dim ID:ID = VerificationUrlParam("ID","int","0")
	'存放该模型数据的二维数组
	Dim Formarray
	Dim ModelTableName
	'''''''''''''''''''''''
	if FormID=0 then
		response.Write "模型解释出现错误!"
	else
		'取对应模型数据和对应模型的自定义字段
		Formarray = Con.QueryData("select M.Formtable,F.fieldname,F.FieldDescription from Sp_Form M,Sp_FormField F where M.ID=F.FormID and M.Id="&FormID&"") 
		'response.Write ubound(Modelarray,1)
		'response.End()
		if ubound(Formarray,1)=0 then 
			response.Write "出现错误,请在模型字段中添加该模型的字段!至少添加一个字段"
			response.End()	
		else
			ModelTableName = Formarray(0,0)		
		end if
		set rs = Con.Query("select * from "&ModelTableName&" where ID="&ID&"")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<form action="" method="post">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)"><%=ModelName%>自定义表单信息展示</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="15%" align=left>&nbsp;</TH>
				<TH align=left>&nbsp;</TH>
		      </TR>
			  <%
			  for i=0 to ubound(Formarray,2)
			  %>
			  <TR class="content-td1">
				<TD><%=Formarray(2,i)%>：</TD>
				<TD><%=rs(""&Formarray(1,i)&"")%></TD>
			  </TR>
			  <%
			  next
			  %>
			  <TR class="content-td1">
				<TD>提交时间：</TD>
				<TD><%=rs("posttime")%></TD>
			  </TR>
			</TABLE>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
%>
