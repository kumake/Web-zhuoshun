<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
  		dim linkCategory:linkCategory = F.ReplaceFormText("txt_linkCategory")
		dim urltype:urltype = F.ReplaceFormText("txt_urltype")
		dim sitename:sitename = F.ReplaceFormText("txt_sitename")
		dim siteurl:siteurl = F.ReplaceFormText("txt_siteurl")
		dim sitelogourl:sitelogourl = F.ReplaceFormText("txt_sitelogourl")		
		dim IsDisplay:IsDisplay = F.ReplaceFormText("txt_ISdisplay")
		if IsDisplay="" then IsDisplay = 0
		'''''''���Ӽ�¼
		Con.execute("insert into Sp_FriendLink (linkCategory,urltype,sitename,siteurl,sitelogourl,sitememo,ISdisplay) values ("&linkCategory&","&urltype&",'"&sitename&"','"&siteurl&"','"&sitelogourl&"','',"&ISdisplay&")")
		Alert "�������ӳɹ�","pulg.link.list.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_sitename").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\nվ�����Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_siteurl").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\nվ�����ӱ�����д!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�޸���������</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">������ƣ�</TD>
				<TD>
				<select name="txt_linkCategory">
				<%
					set Rs = Con.Query("select id,categoryname from Sp_dictionary where categoryID = "&config_FriendLinkCategoryID&"")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&">"&rs("categoryname")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�������ͣ�</TD>
				<TD>
				<input name="txt_urltype" type="radio" onClick="javascript:document.getElementById('logourl').style.display='none';" value="0" checked> �ı�&nbsp;
				<input name="txt_urltype" type="radio" onClick="javascript:document.getElementById('logourl').style.display='';" value="1"> Logo&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>վ�����ƣ�</TD>
				<TD><input name="txt_sitename" type="text" class="input" size="60"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>վ�����ӣ�</TD>
				<TD><input name="txt_siteurl" type="text" class="input" size="60"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1" style="display:none;" id="logourl">
				<TD>Logo��ַ��</TD>
				<TD>
				<input name="txt_sitelogourl" type="text" class="input" size="80"><br>
				<iframe name="file_txt_sitelogourl" frameborder="0"   width="100%" height="30px;" scrolling="no" src="File_Upload.asp?obj=txt_sitelogourl"></iframe>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ���ʾ��</TD>
				<TD><input name="txt_ISdisplay" class="input" type="checkbox" value="1"></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="����ģ��" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
