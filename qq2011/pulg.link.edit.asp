<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="save" then 
  		dim linkCategory:linkCategory = F.ReplaceFormText("txt_linkCategory")
		dim urltype:urltype = F.ReplaceFormText("txt_urltype")
		dim sitename:sitename = F.ReplaceFormText("txt_sitename")
		dim siteurl:siteurl = F.ReplaceFormText("txt_siteurl")
		dim sitelogourl:sitelogourl = F.ReplaceFormText("txt_sitelogourl")		
		dim IsDisplay:IsDisplay = F.ReplaceFormText("txt_ISdisplay")
		if IsDisplay="" then IsDisplay = 0
		'''''''增加记录
		Con.execute("update Sp_FriendLink set linkCategory="&linkCategory&",urltype="&urltype&",sitename='"&sitename&"',siteurl='"&siteurl&"',sitelogourl='"&sitelogourl&"',ISdisplay="&ISdisplay&" where ID="&ItemID&"")
		Alert "修改链接成功","pulg.link.list.asp"
	end if
	if ItemID = 0 then
		Alert "参数传递失败",""
	else
		set rs = Con.Query("select * from Sp_FriendLink where ID="&ItemID&"")
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
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_sitename").value=="")
			{
				alert("Sp_CMS提示\r\n\n站点名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_siteurl").value=="")
			{
				alert("Sp_CMS提示\r\n\n站点链接必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">修改友情链接</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">类别名称：</TD>
				<TD>
				<select name="txt_linkCategory">
				<%
					set Rs1 = Con.Query("select id,categoryname from Sp_dictionary where categoryID = "&config_FriendLinkCategoryID&"")
					if rs1.recordcount<>0 then
					do while not rs1.eof
					response.Write "<option value="&rs1("id")&""
					if rs1("id") = rs("linkCategory") then response.Write " selected"
					response.Write ">"&rs1("categoryname")&"</option>"
					rs1.movenext
					loop
					end if
				%>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>链接类型：</TD>
				<TD>
				<input name="txt_urltype" type="radio" onClick="javascript:document.getElementById('logourl').style.display='none';" value="0" <%if rs("urltype")=0 then response.Write "checked"%>> 文本&nbsp;
				<input name="txt_urltype" type="radio" onClick="javascript:document.getElementById('logourl').style.display='';" value="1" <%if rs("urltype")=1 then response.Write "checked"%>> Logo&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>站点名称：</TD>
				<TD><input name="txt_sitename" type="text" class="input" value="<%=rs("sitename")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>站点链接：</TD>
				<TD><input name="txt_siteurl" type="text" class="input" value="<%=rs("siteurl")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1" style="display:<% if rs("urltype")=1 then:response.Write "":else:response.Write "none":end if%>;" id="logourl">
				<TD>Logo地址：</TD>
				<TD>
				<input name="txt_sitelogourl" type="text" class="input" value="<%=rs("sitelogourl")%>" size="80"><br>
				<iframe name="file_txt_sitelogourl" frameborder="0"   width="100%" height="30px;" scrolling="no" src="File_Upload.asp?obj=txt_sitelogourl"></iframe>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否显示：</TD>
				<TD><input name="txt_ISdisplay" class="input" type="checkbox" <%if rs("ISdisplay")=1 then response.Write "checked"%> value="1"></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="修改链接" class="button" type="submit">
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
