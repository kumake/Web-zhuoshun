<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_MD5.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then
		Dim usergroup:usergroup = F.ReplaceFormText("txt_usergroup")
		Dim username:username = F.ReplaceFormText("txt_username")
		Dim userpwd:userpwd = F.ReplaceFormText("txt_userpwd")
		Dim email:email = F.ReplaceFormText("txt_email")
		''''�����û�
		con.execute("insert into Sp_user (groupID,IsVip,username,userpwd,useremail) values ("&usergroup&",0,'"&username&"','"&MD5(userpwd)&"','"&email&"')")
		''''
		''''��̳�û�����
		'dim sql:sql = "insert into SpForum_user (name,password,mark,grade,alltopicnum,lasttime,type,del) values ('"&username&"','"&md5(userpwd)&"',50,0,0,'"&now()&"',0,0)"
		'response.Write sql
		'response.End()	
		'UIcon.execute(sql)
		'UIcon.query("update SpForum_user set userinfo='"&md5(userpwd)&"','||--||1|images/headpic/1.gif|70|120||2008-7-30|' where username='"&username&"'")	
		Alert "�����û��ɹ�!","pulg.User.list.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
		<script language="javascript" type="text/javascript">
			function checkform()
			{
				if(document.form1.all.txt_username.value=="")
				{
					alert("�û���������д��");
					return false;
				}
				if(document.form1.all.txt_userpwd.value=="")
				{
					alert("���������д��");
					return false;
				}
				if(document.form1.all.txt_email.value=="")
				{
					alert("���������д��");
					return false;
				}
				return true;
			}
		</script>
</HEAD>
<BODY>
<form name="form1" id="form1" action="?action=save" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�û�����</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">�û��飺</TD>
				<TD>
				<select name="txt_usergroup">
				  <%
				  '''�û���
				  set rsgroup = con.Query("select ID,UserGroup from Sp_UserGroup")
				  if rsgroup.recordcount<>0 then
				  do while not rsgroup.eof
				  	response.Write "<option value='"&rsgroup("id")&"'>"&rsgroup("UserGroup")&"</option>"
				  rsgroup.movenext
				  loop
				  rsgroup.close
				  set rsgroup = nothing
				  end if
				  %>
			    </select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�û�����</TD>
				<TD><input name="txt_username" type="text" class="input"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�û����룺</TD>
				<TD><input name="txt_userpwd" type="text" class="input"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�����ʼ���</TD>
				<TD><input name="txt_email" type="text" class="input"></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�����û�" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
