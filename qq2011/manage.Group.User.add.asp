<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Md5.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ID:ID = VerificationUrlParam("ID","int","0")
	if action<>"" and action="save" then 
  		dim username:username = F.ReplaceFormText("txt_username")
		dim userRole:userRole = F.ReplaceFormText("txt_userRole")
		dim userpwd:userpwd = F.ReplaceFormText("txt_userpwd")
		'response.Write rolepower
		'response.End()
		Dim sql:sql = "insert into Sp_Manage (username,userpwd,userRole,ParentID) values ('"&username&"','"&md5(userpwd)&"','"&userRole&"',"&CK("userid")&")"
		Con.execute(sql)
		Alert "���ӹ���Ա�ɹ�","manage.Group.User.list.asp"
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
			if(document.getElementById("txt_username").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n����Ա���Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_userpwd").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n����Ա���������д!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" >���ӹ���Ա</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">����Ա���ƣ�</TD>
				<TD><input name="txt_username" class="input" type="text"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">����Ա���룺</TD>
				<TD><input name="txt_userpwd" class="input" type="text"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<Th colspan="2">������ɫ</Th>
			  </TR>
			  </TR>
			  <TR class="content-td1">
			    <TD colspan="2">
				<%
				sql ="select ID,RoleName from Sp_ManageRole where ID in ("&CK("userrole")&")"
				dim RoleArray:RoleArray = Con.QueryData(sql)
				for j=0 to ubound(RoleArray,2)
				%>
				<input type="checkbox" name="txt_userRole" value="<%=RoleArray(0,j)%>">&nbsp;<%=RoleArray(1,j)%>&nbsp;&nbsp;
				<%
				next
				%>
				</TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="���ӹ���Ա" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
