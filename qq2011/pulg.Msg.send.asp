<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim email:email = VerificationUrlParam("email","string","")
	if action<>"" and action="send" then 
		Dim usergroup:usergroup = F.ReplaceFormText("usergroup")
		Dim ItemUID:ItemUID = F.ReplaceFormText("ItemUID")
		Dim Username:Username = F.ReplaceFormText("txt_Username")
		Dim title:title = F.ReplaceFormText("txt_title")
		Dim content:content = F.ReplaceFormText("txt_content")
		if usergroup=0 then
			''''���͸��ƶ���
			set rs = Con.Query("select * from Sp_User where username='"&Username&"'")
			if rs.recordcount=0 then
				Alert "ϵͳ�����ڸ��û�",""
				response.End()
			end if
			Con.Execute("insert into Sp_Message (senduser,ReceiveUser,title,content) values ('ϵͳ����Ա','"&Username&"','"&title&"','"&content&"')")
			Alert "���ͳɹ�","pulg.msg.send.asp"
		else
			''''���û��鷢��
			set rs = Con.Query("select Username from Sp_User where groupID in ("&ItemUID&")")
			if rs.recordcount=0 then
				Alert "ϵͳ���û�",""
			else
				do while not rs.eof
					Con.Execute("insert into Sp_Message (senduser,ReceiveUser,title,content) values ('ϵͳ����Ա','"&rs("Username")&"','"&title&"','"&content&"')")
				rs.movenext
				loop
			end if
			Alert "���ͳɹ�","pulg.msg.send.asp"
		end if
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
		function checkedItem()
		{
			var Item = 0;
			for (var i=0;i<document.forms[0].elements.length;i++) 
			{ 
				var e = document.forms[0].elements[i]; 
				if (e.name == "ItemUID" && e.disabled==false && e.type=="checkbox") 
				{
					if(e.checked)
					{
						Item = Item + 1
					}
				}				
			} 
			return Item;
		}
		//////////////
		function checkform()
		{
			if(document.getElementById("groupsend").checked)
			{
				if(checkedItem()==0)
				{
					alert("Sp_CMS��ʾ\r\n\n��ѡ���û���!");
					return false;			
				}			
			}
			else
			{
				if(document.getElementById("txt_Username").value=="")
				{
					alert("Sp_CMS��ʾ\r\n\n�û���������д!");
					return false;			
				}
			}
			if(document.getElementById("txt_title").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�ʼ����������д!");
				return false;			
			}
			if(document.getElementById("txt_content").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�ʼ����ݱ�����д!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=send&email=<%=email%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">Ⱥ������</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD colspan="2">
				<input name="usergroup" id="groupsend" type="radio" onClick="javascript:document.getElementById('spanuser').style.display='none';document.getElementById('spangroup').style.display='';" value="1">
				���û��鷢��&nbsp;
				<input name="usergroup" id="usersend" onClick="javascript:document.getElementById('spangroup').style.display='none';document.getElementById('spanuser').style.display='';" type="radio" value="0">���û�����&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1" id="spangroup" style="display:;">
				<TD width="13%">&nbsp;</TD>
				<TD>
				<%
					set rs = Con.Query("select ID,UserGroup from Sp_UserGroup")
					if rs.recordcount<>0 then
						do while not rs.eof
							response.Write "<input type='checkbox' value='"&rs("id")&"' name='ItemUID'>&nbsp;"&rs("UserGroup")&""
						rs.movenext
						loop
					end if
				%>
				</TD>
			  </TR>		
			  <TR class="content-td1" id="spanuser" style="display:none;">
				<TD>�ռ������ƣ�</TD>
				<TD><input name="txt_Username" type="text" class="input" size="60">
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>���ű��⣺</TD>
				<TD><input name="txt_title" type="text" class="input" size="60">
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�������ݣ�</TD>
				<TD><textarea name="txt_content" cols="60" rows="10" class="input"></textarea>
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="���Ͷ���" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
