<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","")
	'response.Write ID
	'response.End()
	Dim ModelArray
	'''''
	Dim ModelGroupArray
	'''''
	Dim SelModelGroupArray
	'''''''''''
	if ModelID<>"" and isnumeric(ModelID) then
		ModelArray = Con.QueryRow("select ID,ModelName,IsDisplay,modeltable,modelCategoryID,IsUserAllowAdd,listTemplate,contentTemplate,submodelID,ISAllowComment,HtmlFilePath,FileNameRule,ModelNewJsTemplate,ModelHotJsTemplate,ModelTopJsTemplate,ModelComJsTemplate,IsModelNeedScore from Sp_Model where ID="&ModelID&"",0)
		'response.Write ModelArray(2)
		'response.End()
	end if
	if action<>"" and action="save" then 
  		dim modelname:modelname = F.ReplaceFormText("txt_ModelNmae")
		dim modeltable:modeltable = F.ReplaceFormText("txt_ModelTable")
		dim modelCategoryID:modelCategoryID = F.ReplaceFormText("txt_ModelCategoryID")
		dim SubModelID:SubModelID = F.ReplaceFormText("txt_submodelID")
		dim ISAllowComment:ISAllowComment = F.ReplaceFormText("txt_ISAllowComment")
		dim IsDisplay:IsDisplay = F.ReplaceFormText("txt_ISdisplay")
		dim IsUserAllowAdd:IsUserAllowAdd = F.ReplaceFormText("txt_IsUserAllowAdd")
		dim IsModelNeedScore:IsModelNeedScore = F.ReplaceFormText("txt_IsModelNeedScore")
		'dim SubModelFilter:SubModelFilter = F.ReplaceFormText("txt_SubModelFilter")
		''''''''
		dim ModelNewJsTemplate:ModelNewJsTemplate = F.ReplaceFormText("txt_ModelNewJsTemplate")
		dim ModelHotJsTemplate:ModelHotJsTemplate = F.ReplaceFormText("txt_ModelHotJsTemplate")
		dim ModelTopJsTemplate:ModelTopJsTemplate = F.ReplaceFormText("txt_ModelTopJsTemplate")
		dim ModelComJsTemplate:ModelComJsTemplate = F.ReplaceFormText("txt_ModelComJsTemplate")
		'''''''
		dim HtmlFilePath:HtmlFilePath = F.ReplaceFormText("txt_HtmlFilePath")
		dim FileNameRule:FileNameRule = F.ReplaceFormText("txt_FileNameRule")
		dim contentTemplate:contentTemplate = F.ReplaceFormText("txt_contentTempalte")
		dim listTemplate:listTemplate = F.ReplaceFormText("txt_listTempalte")
		dim modelgroup:modelgroup = F.ReplaceFormText("txt_usergroup")
		dim isusergroup:isusergroup = 0
		if modelgroup<>"" then isusergroup = 1
		if IsUserAllowAdd="" then IsUserAllowAdd = 0
		if IsDisplay="" then IsDisplay = 0
		if ISAllowComment ="" then ISAllowComment = 0
		if IsModelNeedScore="" then IsModelNeedScore = 0
		
		Con.execute("update Sp_Model set SubModelID="&SubModelID&",modelCategoryID ="&modelCategoryID&",modelname='"&modelname&"',IsDisplay="&IsDisplay&",IsUserAllowAdd = "&IsUserAllowAdd&",contentTemplate='"&contentTemplate&"',listTemplate='"&listTemplate&"',ISAllowComment="&ISAllowComment&",IsUsergroup="&IsUsergroup&",usergroup='"&modelgroup&"',HtmlFilePath = '"&HtmlFilePath&"',FileNameRule = '"&FileNameRule&"',ModelNewJsTemplate="&ModelNewJsTemplate&",ModelHotJsTemplate="&ModelHotJsTemplate&",ModelTopJsTemplate="&ModelTopJsTemplate&",ModelComJsTemplate="&ModelComJsTemplate&",IsModelNeedScore="&IsModelNeedScore&" where ID="&ModelID&"")
		'''''���±�ѡ��ģ�͵ĸ�ģ��ID		
		Con.execute("update sp_model set ParentModelID = "&ModelID&" where ID = "&SubModelID&"")		
		''''''����ģ���û��鹦��
		con.execute("delete * from Sp_ModelGroup where ModelID="&ModelID&"")
		if modelgroup<>"" then			
			ModelGroupArray = split(modelgroup,",")
			for i=0 to ubound(ModelGroupArray)
				con.execute("insert into Sp_ModelGroup (ModelID,groupID) values ("&ModelID&","&ModelGroupArray(i)&")")
			next
		end if
		Alert "�޸�ģ�ͳɹ�","setting.Model.list.asp"
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
    <script language="javascript" type="text/javascript" src="Scripts/popup.js"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_ModelNmae").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\nģ�����Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_ModelTable").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\nģ�ͱ������д!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save&ModelID=<%=ModelID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�޸�ģ��</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">ģ�����ƣ�</TD>
				<TD><input name="txt_ModelNmae" class="input" type="text" value="<%=ModelArray(1)%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>ģ�ͱ����ƣ�</TD>
				<TD><input name="txt_ModelTable" class="input" type="text" value="<%=ModelArray(3)%>" disabled><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>ģ��ϵͳ���</TD>
				<TD>
				<select name="txt_ModelCategoryID">
				<%
					set Rs = Con.Query("select id,dictionary from Sp_dictionaryCategory")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&""
					if cint(ModelArray(4))=cint(rs("id")) then response.Write " selected"
					response.Write ">"&rs("dictionary")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>&nbsp;<a href="setting.dictionary.add.asp">��������</a>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>ģ���ļ����Ŀ¼��</TD>
				<TD><input name="txt_HtmlFilePath" type="text" class="input" size="30" value="<%=ModelArray(10)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ļ���������</TD>
				<TD><input name="txt_FileNameRule" type="text" class="input" size="30" value="<%=ModelArray(11)%>">&nbsp;&nbsp;<span class="huitext">news_{page}.html</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>������ϢJavascriptģ��ID��</TD>
				<TD><input name="txt_ModelNewJsTemplate" type="text" class="input" size="10" value="<%=ModelArray(12)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ȵ���ϢJavascriptģ��ID��</TD>
				<TD><input name="txt_ModelHotJsTemplate" type="text" class="input" size="10" value="<%=ModelArray(13)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ö���ϢJavascriptģ��ID��</TD>
				<TD><input name="txt_ModelTopJsTemplate" type="text" class="input" size="10" value="<%=ModelArray(14)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƽ���ϢJavascriptģ��ID��</TD>
				<TD><input name="txt_ModelComJsTemplate" type="text" class="input" size="10" value="<%=ModelArray(15)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�б�ģ�壺</TD>
				<TD><input name="txt_listTempalte" value="<%=ModelArray(6)%>" type="text" class="input" size="70">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_listTempalte');">ѡ��</a></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>����ҳģ�壺</TD>
				<TD><input name="txt_contentTempalte" value="<%=ModelArray(7)%>" type="text" class="input" size="70">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_contentTempalte');">ѡ��</a></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ϵģ��(�¼�)��</TD>
				<TD>
				<select name="txt_submodelID">
				<option value="0">-----</option>
				<%
				set rs = Con.Query("select ID,modelname from Sp_model where id<>"&ModelID&"")
				if rs.recordcount<>0 then
				do while not rs.eof
					response.Write "<option value='"&rs("ID")&"'"
					if cint(ModelArray(8))=cint(rs("ID")) then response.Write " selected"
					response.Write ">"&rs("modelname")&"</option>"
				rs.movenext
				loop
				end if
				%>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ��������ۣ�</TD>
				<TD><input name="txt_ISAllowComment" class="input" type="checkbox" value="1" <%if ModelArray(9)=1 then response.Write " checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ���ʾ��</TD>
				<TD><input name="txt_ISdisplay" class="input" type="checkbox" value="1" <%if ModelArray(2)=1 then response.Write " checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>ǰ̨�û��Ƿ�����¼�룺</TD>
		        <TD><input name="txt_IsUserAllowAdd" class="input" type="checkbox" value="1" <%if ModelArray(5)=1 then response.Write " checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>��ģ�͵���Ϣ�Ƿ���Ҫ���ֲ鿴��</TD>
			    <TD><input name="txt_IsModelNeedScore" type="checkbox" value="1" <%if ModelArray(16)=1 then response.Write " checked"%>></TD>
		      </TR>
			  <TR class="content-td1">
			    <TD>�û���Ȩ�ޣ�</TD>
		        <TD>
				<%
				set rsg = con.query("select id,UserGroup from Sp_UserGroup")
				if rsg.recordcount<>0 then
				do while not rsg.eof
					response.Write "<input type='checkbox' name='txt_usergroup' value='"&rsg("id")&"'"
					set rspp = con.query("select * from Sp_ModelGroup where ModelID="&ModelID&" and GroupID="&rsg("id")&"")
					if rspp.recordcount<>0 then response.Write " checked"
					response.Write ">"&rsg("UserGroup")&"<br>"
				rsg.movenext
				loop
				end if
				%>
				</TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�޸�ģ��" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<script language=javascript>
    //ѡ��ģ��
	var g_pop;
    function SelectTemplate(obj)
    {
        g_pop=new Popup({ contentType:1, isReloadOnClose:false,scrollType:'no', width:700, height:320 });
        g_pop.setContent("title","ѡ��ģ��");
        g_pop.setContent("contentUrl","Template.asp?obj="+obj);
    	
        g_pop.build();
        g_pop.show();
        return false;
    }
    
    //�رմ򿪵Ĵ���
    function ClosePop()
    {
        g_pop.close();
    }
</script>
