<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ID:ID = VerificationUrlParam("ID","int","0")
	if action<>"" and action="save" then 
  		dim RoleName:RoleName = F.ReplaceFormText("txt_RoleName")
		dim rolepower:rolepower = F.ReplaceFormText("txt_rolepower")
		'response.Write rolepower
		'response.End()
		Con.execute("update Sp_ManageRole set RoleName='"&RoleName&"',RolePower='"&rolepower&"',RoleCode='"&rolepower&"' where ID="&ID&" ")
		Alert "�޸Ľ�ɫ�ɹ�","manage.Group.list.asp"
	end if
	if ID=0 then
		Alert "��������ʧ��",""
	else
		set rs = Con.Query("select * from Sp_ManageRole where ID="&ID&"")
		if rs.recordcount=0 then
			Alert "��Ӧ��Ϣ�Ҳ���",""
		else
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
			if(document.getElementById("txt_RoleName").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n��ɫ���Ʊ�����д!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save&ID=<%=ID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" >�޸Ľ�ɫ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">��ɫ���ƣ�</TD>
				<TD><input name="txt_RoleName" class="input" type="text" value="<%=rs("RoleName")%>" <%if action="view" then response.Write "disabled"%> ><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<Th colspan="2">��ɫȨ��</Th>
			  </TR>
			  </TR>
			  <%
			  dim PowerArray:PowerArray = Con.QueryData("select ID,ModelName,ModelCode,PowerAction,powerMemo,ParentID from Sp_ManagePower order by ParentID asc")
			  if ubound(PowerArray,2)>0 then
			  Dim parentID:parentID = 0 
			  for i=0 to ubound(PowerArray,2)
			  	if PowerArray(5,i)=0 then
				parentID = PowerArray(0,i)
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="txt_rolepower" value="<%=PowerArray(2,i)%>" <%if instr(rs("RoleCode"),PowerArray(2,i))>0 then response.Write "checked"%> value="<%=rs("RoleName")%>" <%if action="view" then response.Write "disabled"%>><%=PowerArray(1,i)%></TD>
				<TD><span class="huitext"><input type="checkbox" name="List_action" value="list">�鿴<input type="checkbox" name="List_action" value="edit">�޸�<input type="checkbox" name="List_action" value="add">����<input type="checkbox" name="List_action" value="del">ɾ��</span></TD>
			  </TR>
			  <%
			  end if
			  next
			  end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
		<%if action<>"view" then%>
		  <input name="btnsearch" value="�޸Ľ�ɫ" class="button" type="submit">
		<%end if%>
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
