<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="save" then 	
		Dim departname:departname = F.ReplaceFormText("txt_departname")
		Dim jobname:jobname = F.ReplaceFormText("txt_jobname")
		Dim personnum:personnum = F.ReplaceFormText("txt_personnum")
		Dim jobrequire:jobrequire = F.ReplaceFormText("txt_jobrequire")
		Dim jobduty:jobduty = F.ReplaceFormText("txt_jobduty")
		Dim validtime:validtime = F.ReplaceFormText("txt_validtime")
		Dim isactive:isactive = F.ReplaceFormText("txt_isactive")
		Dim IsOnlineDeliverResuem:IsOnlineDeliverResuem = F.ReplaceFormText("txt_IsOnlineDeliverResuem")
		Dim memo:memo = F.ReplaceFormText("txt_memo")
		if IsOnlineDeliverResuem="" then IsSingle = 0
		if isactive="" then isactive = 0		
		'''''''增加记录
		Con.execute("update Sp_Job set departname="&departname&",jobname='"&jobname&"',personnum="&personnum&",jobrequire='"&jobrequire&"',jobduty='"&jobduty&"',validtime='"&validtime&"',isactive="&isactive&",IsOnlineDeliverResuem="&IsOnlineDeliverResuem&",[memo]='"&memo&"' where ID="&ItemID&"")
		Alert "修改招聘岗位成功","pulg.job.list.asp"
	end if
	if ItemID = 0 then
		Alert "参数传递失败",""
	else
		set rs = Con.Query("select * from Sp_Job where ID="&ItemID&"")
		if rs.recordcount=0 then
			Alert "找不到记录",""
		else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/calender.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_jobname").value=="")
			{
				alert("Sp_CMS提示\r\n\n岗位名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_personnum").value=="")
			{
				alert("Sp_CMS提示\r\n\n招聘人数必须填写!");
				return false;			
			}
			if(document.getElementById("txt_jobrequire").value=="")
			{
				alert("Sp_CMS提示\r\n\n岗位要求必须填写!");
				return false;			
			}
			if(document.getElementById("txt_jobduty").value=="")
			{
				alert("Sp_CMS提示\r\n\n岗位职责必须填写!");
				return false;			
			}
			if(document.getElementById("txt_validtime").value=="")
			{
				alert("Sp_CMS提示\r\n\n招聘截止时间必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">人才招聘增加职位</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">部门名称：</TD>
				<TD>
				<select name="txt_departname">
				<%
					set Rs1 = Con.Query("select id,categoryname from Sp_dictionary where categoryID = "&config_JobDepartCategoryID&"")
					if rs1.recordcount<>0 then
					do while not rs1.eof
					response.Write "<option value="&rs1("id")&""
					if rs1("id") = rs("departname") then response.Write " selected"
					response.Write ">"&rs1("categoryname")&"</option>"
					rs1.movenext
					loop
					end if
				%>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>岗位名称：</TD>
				<TD><input name="txt_jobname" type="text" value="<%=rs("jobname")%>" class="input"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>招聘人数：</TD>
				<TD><input name="txt_personnum" type="text" value="<%=rs("personnum")%>" class="input" value="1"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>岗位职责：</TD>
				<TD><textarea name="txt_jobduty" cols="60" rows="4"><%=rs("jobduty")%></textarea><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>岗位需求：</TD>
				<TD><textarea name="txt_jobrequire" cols="60" rows="4"><%=rs("jobrequire")%></textarea><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>截止时间：</TD>
				<TD><input name="txt_validtime" type="text" value="<%=rs("validtime")%>" class="input" onfocus="javascript:HS_setDate(this);"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>			
			  <TR class="content-td1">
				<TD>是否在线投递简历：</TD>
				<TD><input name="txt_IsOnlineDeliverResuem" class="input" type="checkbox" value="1"  <%if rs("IsOnlineDeliverResuem")=1 then response.Write "checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否显示：</TD>
				<TD><input name="txt_isactive" class="input" type="checkbox" value="1" <%if rs("isactive")=1 then response.Write "checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>备注：</TD>
				<TD><textarea name="txt_memo" cols="60" rows="8" class="input"><%=rs("memo")%></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="修改岗位" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
	end if
%>
