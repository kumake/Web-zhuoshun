<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="save" then 
		dim votename:votename = F.ReplaceFormText("txt_votename")
		dim ItemNum:ItemNum = F.ReplaceFormText("txt_ItemNum")
		dim validstarttime:validstarttime = F.ReplaceFormText("txt_validstarttime")
		dim validendtime:validendtime = F.ReplaceFormText("txt_validendtime")
		dim IsSingle:IsSingle = F.ReplaceFormText("txt_IsSingle")
		dim isactive:isactive = F.ReplaceFormText("txt_isactive")
		dim votememo:votememo = F.ReplaceFormText("txt_votememo")
		if IsSingle="" then IsSingle = 0
		if isactive="" then isactive = 0		
		'''''''增加记录
		Con.execute("update Sp_Vote set votename='"&votename&"',ItemNum="&ItemNum&",validstarttime='"&validstarttime&"',validendtime='"&validendtime&"',IsSingle="&IsSingle&",isactive="&isactive&",votememo='"&votememo&"' where ID="&ItemID&"")
		Alert "修改投票成功","pulg.vote.list.asp"
	end if
	if ItemID=0 then 
		Alert "参数传递失败",""
	else
		set rs = Con.Query("select * from Sp_Vote where ID="&ItemID&"")
		if rs.recordcount=0 then
			Alert "查找信息失败",""
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
			if(document.getElementById("txt_votename").value=="")
			{
				alert("Sp_CMS提示\r\n\n调查名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_validstarttime").value=="")
			{
				alert("Sp_CMS提示\r\n\n开始时间必须填写!");
				return false;			
			}
			if(document.getElementById("txt_validendtime").value=="")
			{
				alert("Sp_CMS提示\r\n\n结束时间必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">修改调查</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">调查名称：</TD>
				<TD><input name="txt_votename" type="text" class="input" size="80" value="<%=rs("votename")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>开始时间：</TD>
				<TD><input name="txt_validstarttime" type="text" class="input" value="<%=rs("validstarttime")%>" onfocus="javascript:HS_setDate(this);"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>结束时间：</TD>
				<TD><input name="txt_validendtime" type="text" class="input" value="<%=rs("validendtime")%>" onfocus="javascript:HS_setDate(this);"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>子项数目：</TD>
				<TD><input name="txt_ItemNum" class="input" type="text" value="<%=rs("ItemNum")%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否单选：</TD>
				<TD><input name="txt_IsSingle" class="input" type="checkbox" value="1" <%if rs("IsSingle")=1 then response.Write "checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否显示：</TD>
				<TD><input name="txt_isactive" class="input" type="checkbox" value="1" <%if rs("isactive")=1 then response.Write "checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>备注：</TD>
				<TD><textarea name="txt_votememo" cols="60" rows="8" class="input"><%=rs("votememo")%></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="修改投票" class="button" type="submit">
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
