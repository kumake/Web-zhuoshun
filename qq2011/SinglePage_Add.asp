<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","")
	Dim ModelName
	Dim ModelPicName
	Dim ModelContent
	Dim IsNeedPic:IsNeedPic = 0
	Dim IsNeedInstr:IsNeedInstr = 0
	Dim ContentInstr
	if action<>"" and action="save" then
		'����
		ModelID = F.ReplaceFormText("txt_ModelID")
		ModelPicName = F.ReplaceFormText("txt_picname")
		ModelContent = F.ReplaceFormText("txt_content")
		ContentInstr  = F.ReplaceFormText("txt_Instr")
		Con.Execute("update user_Singlepage set picname='"&ModelPicName&"',Instr='"&ContentInstr&"',content='"&ModelContent&"' where singlepageID="&ModelID&"")
		Alert "�޸���Ϣ�ɹ�","?ModelID="&ModelID&""
	else
		'ȡֵ
		set rs = Con.Query("select * from user_Singlepage U,Sp_SinglePage S where S.ID=U.singlepageID and U.singlepageID="&ModelID&"")
		if rs.recordcount<>0 then
			ModelName = rs("singlePagename")
			ModelPicName = rs("picname")
			ModelContent =rs("content")
			IsNeedPic = rs("IsNeedPic")
			IsNeedInstr = rs("IsNeedInstr")
			ContentInstr = rs("Instr")
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
		function checkform()
		{
			if(document.getElementById("txt_title").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n��Ϣ���������д!");
				return false;			
			}
			if(document.getElementById("txt_content").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n��Ϣ���ݱ�����д!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)"><%=ModelName%>������Ϣ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent1">
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">��Ϣ���⣺</TD>
				<TD><input name="txt_title" class="input" size="80" type="text" value="<%=ModelName%>" readonly><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <%if IsNeedPic=1 then%>
			  <TR class="content-td1">
				<TD>˵����ͼƬ��</TD>
				<TD>
				<input name="txt_picname" class="input" size="80" type="text" value="<%=ModelPicName%>"><br>
				<iframe name="file_txt_picname" frameborder="0"   width="100%" height="30px;" scrolling="no" src="File_Upload.asp?obj=txt_picname"></iframe>
				</TD>
			  </TR>
			  <%
			  end if
			  '''''''���
			  if IsNeedInstr = 1 then
			  %>
			  <TR class="content-td1">
				<TD>��飺</TD>
				<TD>
				<textarea name='txt_Instr' style="display:none;"><%=ContentInstr%></textarea>
				<iframe src="Html/eWebEditor.asp?id=txt_Instr&style=s_light" height="300" width="100%" scrolling="no"></iframe>
				<span class="huitext">&nbsp;����</span>
				</TD>
			  </TR>
			  <%end if%>			  
			  <TR class="content-td1">
				<TD>��ϸ���ݣ�</TD>
				<TD>
				<input name='txt_ModelID' value="<%=ModelID%>" type='hidden'>
				<textarea name='txt_content' style="display:none;"><%=ModelContent%></textarea>
				<%
				if Cint(config_SinglePageEditor) = 0 then
				%>
				<script>
				// ʹ��ʾ��
				var rnd = Math.round(Math.random()*10000);
				__ImageRoot = 'images/'; // �༭��JS�ļ�����ڱ�ҳ��·��
				__FileDialogPage = 'img.asp?ramd='+rnd; // �ϴ��ļ�ҳ������ڵ�ǰҳ��·��
				__SummaryID = 'Summary'; // Ҫ����ժҪ��HTML�ؼ�ID
				__SummaryLen = 100; // ժҪ�ĳ���
				__ContentID = 'txt_content'; // ������ȡ��д���HTML�ؼ�ID
				HtmlEditor('dd','500','300');
				</script>
				<%
				else
				%>
				<iframe src="Html/eWebEditor.asp?id=txt_content&style=s_light" height="300" width="100%" scrolling="no"></iframe>
				<%
				end if
				%>
				<span class="huitext">&nbsp;����</span>
				</TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="����<%=ModelName%>" <%if ModelID=1 then response.Write "disabled"%> class="buttonlen" type="submit">
		</div>
		
		<br>
		<div>
			<TABLE cellSpacing="1" width="100%">
			<TR class="content-td1">
				<TH align="left" colspan="2">�������</TH>
			</TR>
			<TR class="content-td1">
				<TD>asp</TD>
				<TD>
				��飺UserSinglePageInstr("<%=ModelName%>")&nbsp;&nbsp;
				ͼƬ��UserSinglePagePic("<%=ModelName%>")&nbsp;&nbsp;
				��ϸ��UserSinglePage("<%=ModelName%>")
				<br>
				ID���ã�UserSinglePageByID(SinglePagetype,SinglePageid)	 SinglePagetype��ʾȡֵ����("pic","instr","content")  SinglePageid��ʾID��ʶ��
				</TD>
			</TR>
			</TABLE>
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
