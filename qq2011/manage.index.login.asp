
<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Md5.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<%

	'Dim site_Status:site_Status = "false"
	'Dim site_IdentifyingCode:site_IdentifyingCode = ""
	'Dim site_Domain:site_Domain = request.ServerVariables("HTTP_HOST")
	'if site_Domain<>"localhost" or site_Domain<>"127.0.0.1" or instr(site_Domain,"192.168")<0 then
		''''''������������Ȩ�ļ��Ƚ�
		'site_IdentifyingCode = encryptCode(site_Domain)
		'''''''�ж��Ƿ��½
		'site_Status = ValidationCode(site_IdentifyingCode)
		'if site_Status="false" then
			'Alert "��Ȩ��֤����ȷ!","copyright.asp"
		'end if
	'else
	''''''���ز���׼��ʹ��
	'end if
	'''''1, 2, 3, 4, 5, 6, 7, 8, 9, 10,1, 2, 3, 4, 5
	'''''���庯��ȥ���ظ���Ȩ��
	function splitUserpower(userpower)
		Dim temparray
		Dim tempstr
		temparray = split(trim(userpower),",")
		for i=0 to ubound(temparray)	
			if instr(tempstr,temparray(i))>0 then 
			else
				if tempstr<>"" then
				tempstr = tempstr &","&temparray(i)
				else
				tempstr = temparray(i)
				end if
			end if
		next
		splitUserpower = tempstr
	end function

	Sn.Clear()
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="login" then 
  		dim username:username = F.ReplaceFormText("txt_username")
  		dim userpwd:userpwd = F.ReplaceFormText("txt_userpwd")
		dim vercode:vercode = F.ReplaceFormText("txt_vercode")
		dim userpower
		dim usermenu
		if vercode<>session("GetCode") then
			Alert "��֤���������!","manage.index.login.asp"
		else
			set RS = Con.Query("select * from Sp_Manage where username='"&username&"' and userpwd='"&mD5(userpwd)&"'")
			if rs.recordcount<>0 then
				CK("userid") = rs("id")
				CK("username") = rs("username")
				CK("userrole") = rs("userrole")
				WriteLog username,"login","�ɹ���½"
				'''''
				'response.Redirect("manage.index.load.asp")
				'''''ȡ��ϸȨ��
				'response.Write CK("userrole")
				'response.End()				
				set rs1 = Con.Query("select RolePower from Sp_ManageRole where id in ("&CK("userrole")&")")
				if rs1.recordcount<>0 then
					do while not rs1.eof
						if userpower<>"" then
						userpower = userpower & "," &rs1("RolePower")
						else
						userpower = rs1("RolePower")
						end if
					rs1.movenext
					loop
				end if
				CK("userpower") = splitUserpower(userpower)
				CK("usermenu") = splitUserpower(userpower)
				'if trim(CK("userpower"))<>"" then
					'''''ȡ��ϸ�˵�
					'set rs2 = Con.Query("select ModelCode from Sp_ManagePower where ModelCode in ('"&CK("userpower")&"') and ParentID=0")
					'if rs2.recordcount<>0 then
						'do while not rs2.eof
							'if userpower<>"" then
							'usermenu = usermenu & "," &rs2("ModelCode")
							'else
							'usermenu = rs2("ModelCode")
							'end if
						'rs2.movenext
						'loop
					'end if
					'CK("usermenu") = splitUserpower(usermenu)
				'end if
				response.Redirect("manage.index.asp")
				'Alert "��¼�ɹ�!",""
			else
				WriteLog username,"login","��½ʧ��,��������û�������"
				Alert "�Ҳ�����¼,���������û��������������!","manage.index.login.asp"
			end if
		end if
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<TITLE>asp���̨����Ա����̨</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_username").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�û������Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_userpwd").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n����Ա���������д!");
				return false;			
			}
			if(document.getElementById("txt_vercode").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n��֤�������д!");
				return false;			
			}			
			return true;
		}
	</script>
</head>
<body>
<form action="?action=login" method="post" onSubmit="javascript:return checkform();">
<DIV id="wrap">
	<UL id=tags>
	  <LI class="selectTag" style="LEFT: 0px; TOP: 0px; padding-left:20px; padding-right:20px;">��̨����Ա��¼<span style="color:#FF0000;"> QzxCms 09��</span></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content">
		<!--��3����ǩ//-->
		<DIV>
			<TABLE cellSpacing=1 width="500">
			<TR class="content-td1">
				<TD width="14%">�û�����</TD>
				<TD><input name="txt_username" class="input" type="text"></TD>
			</TR>
			<TR class="content-td1">
				<TD>��&nbsp;&nbsp;&nbsp;�룺</TD>
				<TD><input name="txt_userpwd" class="input" type="password"></TD>
			</TR>
			<TR class="content-td1">
				<TD>��֤�룺</TD>
				<TD><input name="txt_vercode" class="input" type="text">&nbsp;<img align="absmiddle" onClick="javascript:this.src='../Sp_inc/class_Checkcode.asp';" style="border:1px solid #999999; cursor:hand;" src="../Sp_inc/class_Checkcode.asp"></TD>
			</TR>
			<TR class="content-td1">
				<TD colspan="2"><input name="btnuserlogin" value="��  ½" class="button" type="submit"></TD>
			</TR>
		  </TABLE>
		  <br>
		  Copyright (c) 2008-2010  All Rights Reserved 
	
		</DIV>
	</DIV>
</DIV>
</div>
</form>
</body>
</html>