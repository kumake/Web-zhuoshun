<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="replay" then
		Dim replay:replay = F.ReplaceFormText("txt_replay")
		Dim IsAudit:IsAudit = F.ReplaceFormText("txt_IsAudit")
		Dim replaytime:replaytime = F.ReplaceFormText("txt_replaytime")
		Con.Execute("update Sp_GuestBook set IsAudit="&IsAudit&",replay='"&replay&"',replaytime='"&replaytime&"' where ID="&ItemID&"")
		Alert "回复成功","plug.guestbook.view.asp?ItemID="&ItemID&""
	end if
	''''''
	if ItemID=0 then
		Alert "出现错误",""
	else
		set rs = Con.Query("select * from Sp_GuestBook where ID="&ItemID&"")
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
			if(document.getElementById("replay").value=="")
			{
				alert("Sp_CMS提示\r\n\n回复内容必须填写!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=replay&ItemID=<%=ItemID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">留言内容展示</A></LI>
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
				<TD>用户名：</TD>
				<TD><%=rs("username")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>联系电话：</TD>
				<TD><%=rs("tel")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>电子邮件：</TD>
				<TD><%=rs("email")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>留言内容：</TD>
				<TD><%=rs("content")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>提交时间：</TD>
				<TD><%=rs("addtime")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>提交IP：</TD>
				<TD><%=rs("ip")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否审核：</TD>
				<TD>
				<input type="radio" name="txt_IsAudit" value="1" <%if rs("IsAudit")=1 then response.Write "checked"%>>审核&nbsp;
				<input type="radio" name="txt_IsAudit" value="0" <%if rs("IsAudit")=0 then response.Write "checked"%>>未审核&nbsp;
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>回复内容：</TD>
				<TD><textarea name="txt_replay" cols="60" rows="8"><%=rs("replay")%></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>回复时间：</TD>
				<TD><input name="txt_replaytime" type="text" class="input" value="<%=rs("replaytime")%>" onfocus="javascript:HS_setDate(this);"></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="回复留言" class="button" type="submit">
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
