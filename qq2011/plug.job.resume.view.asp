<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then
		Dim Companyevaluation:Companyevaluation = F.ReplaceFormText("txt_Companyevaluation")
		Dim Ishire:Ishire = F.ReplaceFormText("txt_Ishire")
		Dim sql:sql = "update Sp_Resume set Companyevaluation='"&Companyevaluation&"',Ishire="&Ishire&" where ID="&ItemID&""
		'response.Write sql
		'response.End()
		Con.Execute(sql)
		Alert "成功","plug.job.resume.view.asp?ItemID="&ItemID&""
	end if
	''''''
	if ItemID=0 then
		Alert "出现错误",""
	else
		set rs = Con.Query("select * from Sp_Resume where ID="&ItemID&"")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/calender.js" type="text/javascript"></script>
	<!--有些字段必须填写的javascript的验证//-->
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_Companyevaluation").value=="")
			{
				alert("Sp_CMS提示\r\n\n人力资源评价必须填写!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save&ItemID=<%=ItemID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">人才信息</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="15%" align=left>&nbsp;</TH>
				<TH align=left>&nbsp;</TH>
		      </TR>
			  <TR class="content-td1">
				<TD>职位：</TD>
				<TD><%=rs("jobname")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>用户名：</TD>
				<TD><%=rs("username")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>性别：</TD>
				<TD><%=rs("sex")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>电话：</TD>
				<TD><%=rs("tel")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>移动电话：</TD>
				<TD><%=rs("mobel")%></TD>
			  </TR>
			  <%
				''''自定义字段展示
				'所有自定义字段是根据表名称来区分
				Dim tablename:tablename = "Sp_Resume"
				Dim rsFieldArray:rsFieldArray = con.queryData("select F.ID,F.fieldname,F.FieldUI,F.FieldDescription,F.fieldtype,F.fieldlength,F.isNull,F.Attribute,F.defaultValue from Sp_Table T,Sp_TableField F where T.id=F.tableID and T.tablename='"&tablename&"'")
				'''用户组自定义字段
				if ubound(rsFieldArray,1)=0 then
				else
				for i=0 to ubound(rsFieldArray,2)
				%>
				<tr class="content-td1">
				  <td><%=rsFieldArray(3,i)%>：</td>
				  <td><%=FieldWebUI(rsFieldArray(1,i),rsFieldArray(2,i),rs(""&rsFieldArray(1,i)&""))%></td>
				</tr>
				<%
				'
				'''用户组自定义字段
				next
				end if
				%>
			  <TR class="content-td1">
				<TD>简历提交日期：</TD>
				<TD><%=year(rs("posttime"))%>-<%=month(rs("posttime"))%>-<%=day(rs("posttime"))%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>状态：</TD>
				<TD>
				<input type="radio" name="txt_Ishire" value="-1" <%if rs("Ishire")=-1 then response.Write "checked"%>>淘汰&nbsp;
				<input type="radio" name="txt_Ishire" value="1" <%if rs("Ishire")=1 then response.Write "checked"%>>通知面试&nbsp;
				<input type="radio" name="txt_Ishire" value="2" <%if rs("Ishire")=2 then response.Write "checked"%>>正式录用&nbsp;
				<input type="radio" name="txt_Ishire" value="0" <%if rs("Ishire")=0 then response.Write "checked"%>>新加&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>人力资源评价：</TD>
				<TD><textarea name="txt_Companyevaluation" cols="60" rows="8" class="input"><%=rs("Companyevaluation")%></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="人才设置" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
%>
