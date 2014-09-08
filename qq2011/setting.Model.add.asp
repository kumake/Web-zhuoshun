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
		'''''取最新的ID
		ItemArray = Con.QueryRow("select top 1 id from Sp_Model order by id desc",0)
		'''''更新被选中模型的父模型ID
		Con.execute("update sp_model set ParentModelID = "&ItemArray(0)&" where ID = "&SubModelID&"")		
		''''增加物理表
		Con.execute("select * into "&modeltable&" from Sp_TableStructure where 1<>1")
		Con.execute("alter table "&modeltable&" add primary key (id)")
		'''修改字段默认值
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
		''''''向权限管理表中加入模型的权限记录
		Con.Execute("insert into Sp_ManagePower (ModelName,ModelCode,ModelID,ParentID) values ('内容模型_"&modelname&"','"&modeltable&"',"&ItemArray(0)&",0)")
		
		''''''增加模型用户组功能
		if modelgroup<>"" then
			modelgroupArray = split(modelgroup,",")
			for i=0 to ubound(modelgroupArray)
				con.execute("insert into Sp_ModelGroup (ModelID,groupID) values ("&ItemArray(0)&","&modelgroupArray(i)&")")
			next
		end if
		Alert "增加模型成功","setting.Model.list.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
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
				alert("Sp_CMS提示\r\n\n模型名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_ModelTable").value=="")
			{
				alert("Sp_CMS提示\r\n\n模型表必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">增加模型</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="18%">模型名称：</TD>
				<TD><input name="txt_ModelNmae" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>模型表名称：</TD>
				<TD><input name="txt_ModelTable" class="input" type="text" value="user_"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>模型系统类别：</TD>
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
				</select>&nbsp;<a href="setting.dictionary.add.asp">添加新类别</a>
				</TD>
			  </TR>
			  
			  <TR class="content-td1">
				<TD>模型文件存放目录：</TD>
				<TD><input name="txt_HtmlFilePath" type="text" class="input" size="30"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>文件命名规则：</TD>
				<TD><input name="txt_FileNameRule" type="text" class="input" size="30">&nbsp;&nbsp;<span class="huitext">news_{page}.html</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>最新信息Javascript模板ID：</TD>
				<TD><input name="txt_ModelNewJsTemplate" type="text" class="input" size="10" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>热点信息Javascript模板ID：</TD>
				<TD><input name="txt_ModelHotJsTemplate" type="text" class="input" size="10" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>置顶信息Javascript模板ID：</TD>
				<TD><input name="txt_ModelTopJsTemplate" type="text" class="input" size="10" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>推荐信息Javascript模板ID：</TD>
				<TD><input name="txt_ModelComJsTemplate" type="text" class="input" size="10" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>列表模板：</TD>
				<TD><input name="txt_listTempalte" type="text" class="input" size="70">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_listTempalte');">选择</a></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>内容页模板：</TD>
				<TD><input name="txt_contentTempalte" type="text" class="input" size="70">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_contentTempalte');">选择</a></TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>关系模型(下级)：</TD>
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
				<TD>是否允许评论：</TD>
				<TD><input name="txt_ISAllowComment" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否显示：</TD>
				<TD><input name="txt_ISdisplay" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>前台用户是否允许录入：</TD>
		        <TD><input name="txt_IsUserAllowAdd" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>该模型的信息是否需要积分查看：</TD>
			    <TD><input name="txt_IsModelNeedScore" type="checkbox" value="1"></TD>
		      </TR>
			  <TR class="content-td1">
			    <TD>用户组权限：</TD>
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
		  <input name="btnsearch" value="增加模型" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<script language=javascript>
    //选择模板
	var g_pop;
    function SelectTemplate(obj)
    {
        g_pop=new Popup({ contentType:1, isReloadOnClose:false,scrollType:'no', width:700, height:320 });
        g_pop.setContent("title","选择模板");
        g_pop.setContent("contentUrl","Template.asp?obj="+obj);
    	
        g_pop.build();
        g_pop.show();
        return false;
    }
    
    //关闭打开的窗口
    function ClosePop()
    {
        g_pop.close();
    }
</script>
