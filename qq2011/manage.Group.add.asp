<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
  		dim RoleName:RoleName = F.ReplaceFormText("txt_RoleName")
		dim rolepower:rolepower = F.ReplaceFormText("txt_rolepower")
		Con.execute("insert into Sp_ManageRole (RoleName,RolePower,RoleCode) values ('"&RoleName&"','"&rolepower&"','"&rolepower&"')")
		''''ȡ���ֵ
		Dim RoleItem:RoleItem = Con.QueryRow("select top 1 ID from Sp_ManageRole order by id desc",0)
		''''���³�������Ա�˺�
		Con.execute("update Sp_Manage set userrole = userrole&',"&RoleItem(0)&"' where username='admin'")
		Alert "���ӽ�ɫ�ɹ�","manage.Group.list.asp"
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
<form action="?action=save" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" >���ӽ�ɫ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="14%">��ɫ���ƣ�</TD>
				<TD colspan="5"><input name="txt_RoleName" class="input" type="text" value=""><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<Th colspan="6">��ɫȨ��</Th>
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
					<TD><input type="checkbox" name="txt_rolepower" value="<%=PowerArray(2,i)%>"><%=PowerArray(1,i)%></TD>
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
		  <input name="btnsearch" value="���ӽ�ɫ" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
