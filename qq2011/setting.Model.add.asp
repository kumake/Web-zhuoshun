<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
		dim modelgroupArray
  		dim modelname:modelname = F.ReplaceFormText("txt_ModelNmae")
		dim modeltable:modeltable = F.ReplaceFormText("txt_ModelTable")
		dim modelCategoryID:modelCategoryID = F.ReplaceFormText("txt_ModelCategoryID")
		dim SubModelID:SubModelID = F.ReplaceFormText("txt_SubModelID")
		dim ISAllowComment:ISAllowComment = F.ReplaceFormText("txt_ISAllowComment")
		dim IsDisplay:IsDisplay = F.ReplaceFormText("txt_ISdisplay")
		dim IsUserAllowAdd:IsUserAllowAdd = F.ReplaceFormText("txt_IsUserAllowAdd")
		dim IsModelNeedScore:IsModelNeedScore = F.ReplaceFormText("txt_IsModelNeedScore")
		if IsDisplay="" then IsDisplay = 0
		if IsUserAllowAdd="" then IsUserAllowAdd = 0
		if ISAllowComment="" then ISAllowComment = 0
		if IsModelNeedScore="" then IsModelNeedScore = 0
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
		'response.Write modelgroup
		'response.End()
		if modelgroup<>"" then isusergroup = 1
		''''
		dim sql:sql = "insert into Sp_Model (modelname ,modeltable,modelCategoryID,SubModelID,IsDisplay,IsUserAllowAdd,IsModelNeedScore,listTemplate,contentTemplate,ISAllowComment,isusergroup,usergroup,HtmlFilePath,FileNameRule,ModelNewJsTemplate,ModelHotJsTemplate,ModelTopJsTemplate,ModelComJsTemplate) values ('"&modelname&"','"&modeltable&"',"&modelCategoryID&","&SubModelID&","&IsDisplay&","&IsUserAllowAdd&","&IsModelNeedScore&",'"&listTemplate&"','"&contentTemplate&"',"&ISAllowComment&","&isusergroup&",'"&usergroup&"','"&HtmlFilePath&"','"&FileNameRule&"',"&ModelNewJsTemplate&","&ModelHotJsTemplate&","&ModelTopJsTemplate&","&ModelComJsTemplate&")"
		'response.Write sql
		'response.End()
		Con.execute(sql)
		'''''ȡ���µ�ID
		ItemArray = Con.QueryRow("select top 1 id from Sp_Model order by id desc",0)
		'''''���±�ѡ��ģ�͵ĸ�ģ��ID
		Con.execute("update sp_model set ParentModelID = "&ItemArray(0)&" where ID = "&SubModelID&"")		
		''''���������
		Con.execute("select * into "&modeltable&" from Sp_TableStructure where 1<>1")
		Con.execute("alter table "&modeltable&" add primary key (id)")
		'''�޸��ֶ�Ĭ��ֵ
		Con.execute("alter table "&modeltable&" ALTER IndexID int default 0")
		Con.execute("alter table "&modeltable&" ALTER categoryID int default 1")
		Con.execute("alter table "&modeltable&" ALTER IsJumpUrl int default 0")
		Con.execute("alter table "&modeltable&" ALTER InfoType int default 0")
		Con.execute("alter table "&modeltable&" ALTER UserUID int default 0")
		Con.execute("alter table "&modeltable&" ALTER Hots int default 0")
		Con.execute("alter table "&modeltable&" ALTER IsHot int default 0")
		Con.execute("alter table "&modeltable&" ALTER IsTop int default 0")
		Con.execute("alter table "&modeltable&" ALTER IsCommand int default 0")
		Con.execute("alter table "&modeltable&" ALTER CommentCount int default 0")
		Con.execute("alter table "&modeltable&" ALTER IsDel int default 0")
		Con.execute("alter table "&modeltable&" ALTER isusergroup int default 0")
		Con.execute("alter table "&modeltable&" ALTER PostTime datetime default now()")
		Con.execute("alter table "&modeltable&" ALTER LastUpdateTime datetime default now()")
		''''''��Ȩ�޹�����м���ģ�͵�Ȩ�޼�¼
		Con.Execute("insert into Sp_ManagePower (ModelName,ModelCode,ModelID,ParentID) values ('����ģ��_"&modelname&"','"&modeltable&"',"&ItemArray(0)&",0)")
		
		''''''����ģ���û��鹦��
		if modelgroup<>"" then
			modelgroupArray = split(modelgroup,",")
			for i=0 to ubound(modelgroupArray)
				con.execute("insert into Sp_ModelGroup (ModelID,groupID) values ("&ItemArray(0)&","&modelgroupArray(i)&")")
			next
		end if
		Alert "����ģ�ͳɹ�","setting.Model.list.asp"
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
<form action="?action=save" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">����ģ��</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="18%">ģ�����ƣ�</TD>
				<TD><input name="txt_ModelNmae" class="input" type="text" value=""><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>ģ�ͱ����ƣ�</TD>
				<TD><input name="txt_ModelTable" class="input" type="text" value="user_"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>ģ��ϵͳ���</TD>
				<TD>
				<select name="txt_ModelCategoryID">
				<%
					set Rs = Con.Query("select id,dictionary from Sp_dictionaryCategory")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&">"&rs("dictionary")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>&nbsp;<a href="setting.dictionary.add.asp">��������</a>
				</TD>
			  </TR>
			  
			  <TR class="content-td1">
				<TD>ģ���ļ����Ŀ¼��</TD>
				<TD><input name="txt_HtmlFilePath" type="text" class="input" size="30"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ļ���������</TD>
				<TD><input name="txt_FileNameRule" type="text" class="input" size="30">&nbsp;&nbsp;<span class="huitext">news_{page}.html</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>������ϢJavascriptģ��ID��</TD>
				<TD><input name="txt_ModelNewJsTemplate" type="text" class="input" size="10" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ȵ���ϢJavascriptģ��ID��</TD>
				<TD><input name="txt_ModelHotJsTemplate" type="text" class="input" size="10" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ö���ϢJavascriptģ��ID��</TD>
				<TD><input name="txt_ModelTopJsTemplate" type="text" class="input" size="10" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƽ���ϢJavascriptģ��ID��</TD>
				<TD><input name="txt_ModelComJsTemplate" type="text" class="input" size="10" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�б�ģ�壺</TD>
				<TD><input name="txt_listTempalte" type="text" class="input" size="70">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_listTempalte');">ѡ��</a></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>����ҳģ�壺</TD>
				<TD><input name="txt_contentTempalte" type="text" class="input" size="70">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_contentTempalte');">ѡ��</a></TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>��ϵģ��(�¼�)��</TD>
				<TD>
				<select name="txt_SubModelID">
				<option value="0">-----</option>
				<%
				set rs = Con.Query("select ID,modelname from Sp_model")
				if rs.recordcount<>0 then
				do while not rs.eof
					response.Write "<option value='"&rs("ID")&"'>"&rs("modelname")&"</option>"
				rs.movenext
				loop
				end if
				%>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ��������ۣ�</TD>
				<TD><input name="txt_ISAllowComment" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ���ʾ��</TD>
				<TD><input name="txt_ISdisplay" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>ǰ̨�û��Ƿ�����¼�룺</TD>
		        <TD><input name="txt_IsUserAllowAdd" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>��ģ�͵���Ϣ�Ƿ���Ҫ���ֲ鿴��</TD>
			    <TD><input name="txt_IsModelNeedScore" type="checkbox" value="1"></TD>
		      </TR>
			  <TR class="content-td1">
			    <TD>�û���Ȩ�ޣ�</TD>
		        <TD>
				<%
				set rsg = con.query("select id,UserGroup from Sp_UserGroup")
				if rsg.recordcount<>0 then
				do while not rsg.eof
					response.Write "<input type='checkbox' name='txt_usergroup' value='"&rsg("id")&"'>"&rsg("UserGroup")&"<br>"
				rsg.movenext
				loop
				end if
				%>
				</TD>
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
