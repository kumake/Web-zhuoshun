<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
		Dim IsRelativePath:IsRelativePath = F.ReplaceFormText("txt_IsRelativePath")
		Dim Verison:Verison = F.ReplaceFormText("txt_Verison")
		Dim sitename:sitename = F.ReplaceFormText("txt_SiteNmae")
		Dim seokeyword:seokeyword = F.ReplaceFormText("txt_seokeyword")
		Dim seocontent:seocontent = F.ReplaceFormText("txt_seocontent")
		Dim sitename_en:sitename_en = F.ReplaceFormText("txt_SiteNmae_en")
		Dim seokeyword_en:seokeyword_en = F.ReplaceFormText("txt_seokeyword_en")
		Dim seocontent_en:seocontent_en = F.ReplaceFormText("txt_seocontent_en")
		Dim guestSplitWord:guestSplitWord = F.ReplaceFormText("txt_guestSplitWord")	
		Dim guestbookCategoryID:guestbookCategoryID = F.ReplaceFormText("txt_guestbookCategoryID")
		
		Dim guestbookStatus:guestbookStatus = F.ReplaceFormText("txt_guestbookStatus")
		Dim guestbookSplit:guestbookSplit = F.ReplaceFormText("txt_guestbookSplit")
		Dim guestbookAudit:guestbookAudit = F.ReplaceFormText("txt_guestbookAudit")
		Dim guestbookValidation:guestbookValidation = F.ReplaceFormText("txt_guestbookValidation")
		Dim MailSMtp:MailSMtp = F.ReplaceFormText("txt_MailSMtp")
		Dim EmailAccount:EmailAccount = F.ReplaceFormText("txt_EmailAccount")
		Dim EmailAccountUserName:EmailAccountUserName = F.ReplaceFormText("txt_EmailAccountUserName")
		Dim EmailAccountUserPWD:EmailAccountUserPWD = F.ReplaceFormText("txt_EmailAccountUserPWD")
		Dim prohibitIP:prohibitIP = F.ReplaceFormText("txt_prohibitIP")
		Dim IsprohibitIP:IsprohibitIP = F.ReplaceFormText("txt_IsprohibitIP")			
		
		Dim AdsCategoryID:AdsCategoryID = F.ReplaceFormText("txt_AdsCategoryID")	
		Dim FriendLinkCategoryID:FriendLinkCategoryID = F.ReplaceFormText("txt_FriendLinkCategoryID")	
		Dim JobDepartCategoryID:JobDepartCategoryID = F.ReplaceFormText("txt_JobDepartCategoryID")
		Dim UPloadFileAllowEXT:UPloadFileAllowEXT = F.ReplaceFormText("txt_UPloadFileAllowEXT")
		Dim DatabasePath:DatabasePath = F.ReplaceFormText("txt_DatabasePath")
		Dim DataBackForder:DataBackForder = F.ReplaceFormText("txt_DataBackForder")
		Dim isRecordUploadFile:isRecordUploadFile = F.ReplaceFormText("txt_isRecordUploadFile")
		Dim baiduguestbookJS:baiduguestbookJS = F.ReplaceFormText("txt_baiduguestbookJS")
		Dim isbaiduguestbookJS:isbaiduguestbookJS = F.ReplaceFormText("txt_isbaiduguestbookJS")
		Dim SinglePageEditor:SinglePageEditor = F.ReplaceFormText("txt_SinglePageEditor")
		Dim IsShowguestbookList:IsShowguestbookList = F.ReplaceFormText("txt_IsShowguestbookList")
		Dim IsShowOrderItem:IsShowOrderItem = F.ReplaceFormText("txt_IsShowOrderItem")
		'''''
		Dim copyright:copyright = replace(F.ReplaceFormText("txt_copyright"),chr(13),"#")
		Dim copyright_en:copyright_en = replace(F.ReplaceFormText("txt_copyright_en"),chr(13),"#")
		'''''
		Dim IsSiteTotal:IsSiteTotal = F.ReplaceFormText("txt_IsSiteTotal")
		Dim SiteTotalJS:SiteTotalJS = F.ReplaceFormText("txt_SiteTotalJS")		
		Dim IsIndexStatic:IsIndexStatic = F.ReplaceFormText("txt_IsIndexStatic")
		Dim IndexTemplate:IndexTemplate = F.ReplaceFormText("txt_IndexTemplate")
		Dim isShowCity:isShowCity = F.ReplaceFormText("txt_isShowCity")
		Dim isNeedUploadPassword:isNeedUploadPassword = F.ReplaceFormText("txt_isNeedUploadPassword")
		Dim UploadPassword:UploadPassword = F.ReplaceFormText("txt_UploadPassword")
		
		''''''写数据库保存
		dim sql:sql = "update Sp_Config set "
		sql = sql & "Verison='"&Verison&"'"
		sql = sql & ",IsRelativePath='"&IsRelativePath&"'"
		sql = sql & ",sitename='"&sitename&"'"
		sql = sql & ",seokeyword='"&seokeyword&"'"
		sql = sql & ",seocontent='"&seocontent&"'"
		sql = sql & ",sitename_en='"&sitename_en&"'"
		sql = sql & ",seokeyword_en='"&seokeyword_en&"'"
		sql = sql & ",seocontent_en='"&seocontent_en&"'"
		sql = sql & ",guestSplitWord='"&guestSplitWord&"'"
		sql = sql & ",guestbookCategoryID='"&guestbookCategoryID&"'"		
		sql = sql & ",guestbookStatus="&guestbookStatus&""	
		sql = sql & ",guestbookSplit="&guestbookSplit&""	
		sql = sql & ",guestbookAudit="&guestbookAudit&""	
		sql = sql & ",guestbookValidation="&guestbookValidation&""	
		sql = sql & ",MailSMtp='"&MailSMtp&"'"	
		sql = sql & ",EmailAccount='"&EmailAccount&"'"	
		sql = sql & ",EmailAccountUserName='"&EmailAccountUserName&"'"	
		sql = sql & ",EmailAccountUserPWD='"&EmailAccountUserPWD&"'"
		sql = sql & ",AdsCategoryID="&AdsCategoryID&""
		sql = sql & ",FriendLinkCategoryID="&FriendLinkCategoryID&""
		sql = sql & ",JobDepartCategoryID="&JobDepartCategoryID&""
		sql = sql & ",UPloadFileAllowEXT='"&UPloadFileAllowEXT&"'"
		sql = sql & ",DatabasePath='"&DatabasePath&"'"
		sql = sql & ",DataBackForder='"&DataBackForder&"'"
		sql = sql & ",prohibitIP='"&prohibitIP&"'"	
		sql = sql & ",IsprohibitIP="&IsprohibitIP&""
		sql = sql & ",isRecordUploadFile="&isRecordUploadFile&""
		sql = sql & ",baiduguestbookJS='"&baiduguestbookJS&"'"
		sql = sql & ",isbaiduguestbookJS="&isbaiduguestbookJS&""
		sql = sql & ",SinglePageEditor="&SinglePageEditor&""		
		sql = sql & ",IsShowguestbookList="&IsShowguestbookList&""
		sql = sql & ",IsShowOrderItem="&IsShowOrderItem&""	
		sql = sql & ",copyright='"&copyright&"'"	
		sql = sql & ",copyright_en='"&copyright_en&"'"	
		sql = sql & ",IsSiteTotal="&IsSiteTotal&""	
		sql = sql & ",SiteTotalJs='"&SiteTotalJs&"'"
		sql = sql & ",IsIndexStatic="&IsIndexStatic&""
		sql = sql & ",IndexTemplate='"&IndexTemplate&"'"
		sql = sql & ",isShowCity="&isShowCity&""		
		sql = sql & ",isNeedUploadPassword="&isNeedUploadPassword&""		
		sql = sql & ",UploadPassword="&UploadPassword&""		
		
		'response.Write sql
		'response.End()
		Con.execute(sql)
		''''''写入配置文件
		Dim Templatestream:Templatestream = F.ReadTextFile ("Base.Setting.Config.template")
		Templatestream = replace(Templatestream,"#Verison#",""&Verison&"")
		Templatestream = replace(Templatestream,"#IsRelativePath#",""&IsRelativePath&"")		
		Templatestream = replace(Templatestream,"#sitename#",""&sitename&"")
		Templatestream = replace(Templatestream,"#seokeyword#",""&seokeyword&"")
		Templatestream = replace(Templatestream,"#seocontent#",""&seocontent&"")
		Templatestream = replace(Templatestream,"#sitename_en#",""&sitename_en&"")
		Templatestream = replace(Templatestream,"#seokeyword_en#",""&seokeyword_en&"")
		Templatestream = replace(Templatestream,"#seocontent_en#",""&seocontent_en&"")
		Templatestream = replace(Templatestream,"#guestSplitWord#",""&guestSplitWord&"")
		Templatestream = replace(Templatestream,"#guestbookCategoryID#",""&guestbookCategoryID&"")		
		Templatestream = replace(Templatestream,"#guestbookStatus#",""&guestbookStatus&"")
		Templatestream = replace(Templatestream,"#guestbookSplit#",""&guestbookSplit&"")
		Templatestream = replace(Templatestream,"#guestbookAudit#",""&guestbookAudit&"")
		Templatestream = replace(Templatestream,"#guestbookValidation#",""&guestbookValidation&"")
		Templatestream = replace(Templatestream,"#MailSMtp#",""&MailSMtp&"")
		Templatestream = replace(Templatestream,"#EmailAccount#",""&EmailAccount&"")
		Templatestream = replace(Templatestream,"#EmailAccountUserName#",""&EmailAccountUserName&"")
		Templatestream = replace(Templatestream,"#EmailAccountUserPWD#",""&EmailAccountUserPWD&"")
		Templatestream = replace(Templatestream,"#AdsCategoryID#",""&AdsCategoryID&"")
		Templatestream = replace(Templatestream,"#FriendLinkCategoryID#",""&FriendLinkCategoryID&"")
		Templatestream = replace(Templatestream,"#JobDepartCategoryID#",""&JobDepartCategoryID&"")
		Templatestream = replace(Templatestream,"#UPloadFileAllowEXT#",""&UPloadFileAllowEXT&"")
		Templatestream = replace(Templatestream,"#DatabasePath#",""&DatabasePath&"")
		Templatestream = replace(Templatestream,"#DataBackForder#",""&DataBackForder&"")
		Templatestream = replace(Templatestream,"#prohibitIP#",""&prohibitIP&"")
		Templatestream = replace(Templatestream,"#IsprohibitIP#",""&IsprohibitIP&"")	
		Templatestream = replace(Templatestream,"#isRecordUploadFile#",""&isRecordUploadFile&"")		
		Templatestream = replace(Templatestream,"#baiduguestbookJS#",""&baiduguestbookJS&"")		
		Templatestream = replace(Templatestream,"#isbaiduguestbookJS#",""&isbaiduguestbookJS&"")		
		Templatestream = replace(Templatestream,"#SinglePageEditor#",""&SinglePageEditor&"")		
		Templatestream = replace(Templatestream,"#IsShowguestbookList#",""&IsShowguestbookList&"")		
		Templatestream = replace(Templatestream,"#IsShowOrderItem#",""&IsShowOrderItem&"")	
		Templatestream = replace(Templatestream,"#copyright#",""&copyright&"")
		Templatestream = replace(Templatestream,"#copyright_en#",""&copyright_en&"")
		Templatestream = replace(Templatestream,"#IsSiteTotal#",""&IsSiteTotal&"")
		Templatestream = replace(Templatestream,"#SiteTotalJs#",""&SiteTotalJs&"")
		Templatestream = replace(Templatestream,"#IsIndexStatic#",""&IsIndexStatic&"")
		Templatestream = replace(Templatestream,"#IndexTemplate#",""&IndexTemplate&"")
		Templatestream = replace(Templatestream,"#isShowCity#",""&isShowCity&"")
		Templatestream = replace(Templatestream,"#isNeedUploadPassword#",""&isNeedUploadPassword&"")
		Templatestream = replace(Templatestream,"#UploadPassword#",""&UploadPassword&"")
		
		F.WriteTextFile "../plugIn/Setting.Config.asp",Templatestream
		Alert "设置成功","setting.config.asp"
	end if
	''''''''
	set rs = Con.Query("select top 1 * from Sp_Config order by id desc")
	if rs.recordcount<>0 then	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
    <script language="javascript" type="text/javascript" src="Scripts/popup.js"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_SiteNmae").value=="")
			{
				alert("Sp_CMS提示\r\n\n网站名称必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">全局系统设置</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent1',this)" href="javascript:void(0)">邮件系统设置</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent2',this)" href="javascript:void(0)">留言板设置</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent3',this)" href="javascript:void(0)">模型字典类别设置</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent4',this)" href="javascript:void(0)">版权信息</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent5',this)" href="javascript:void(0)">网站静态设置</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent6',this)" href="javascript:void(0)">其他设置</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">网站版本：</TD>
				<TD>
				<input name="txt_Verison" type="radio" value="1" <%if rs("Verison")=1 then response.Write "checked"%>>中文版&nbsp;
				<input name="txt_Verison" type="radio" value="2" <%if rs("Verison")=2 then response.Write "checked"%>>中英文版&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>网站上传文件路径：</TD>
				<TD>
				<input name="txt_IsRelativePath" type="radio" value="0" <%if rs("IsRelativePath")=0 then response.Write "checked"%>>绝对/&nbsp;
				<input name="txt_IsRelativePath" type="radio" value="1" <%if rs("IsRelativePath")=1 then response.Write "checked"%>>相对../&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>网站名称：</TD>
				<TD><input name="txt_SiteNmae" type="text" class="input" value="<%=rs("sitename")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>网站Meta标题：</TD>
				<TD><input name="txt_seokeyword" type="text" class="input" value="<%=rs("seokeyword")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>网站Meta描述：</TD>
				<TD><input name="txt_seocontent" type="text" class="input" value="<%=rs("seocontent")%>" size="80">
			    <span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>网站名称(英文版)：</TD>
				<TD><input name="txt_SiteNmae_en" type="text" class="input" value="<%=rs("sitename_en")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>网站Meta标题(英文版)：</TD>
				<TD><input name="txt_seokeyword_en" type="text" class="input" value="<%=rs("seokeyword_en")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>网站Meta描述(英文版)：</TD>
				<TD><input name="txt_seocontent_en" type="text" class="input" value="<%=rs("seocontent_en")%>" size="80">
			    <span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			</TABLE>
		</DIV>
	</DIV>
	
	<DIV class="tagContent selectTag content" id="tagContent1" style="display:none;">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">邮件smtp服务器：</TD>
				<TD><input name="txt_MailSMtp" type="text" class="input" value="<%=rs("MailSMtp")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>邮件帐号：</TD>
				<TD><input name="txt_EmailAccount" type="text" class="input" value="<%=rs("EmailAccount")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>邮件帐号用户名：</TD>
				<TD><input name="txt_EmailAccountUserName" type="text" class="input" value="<%=rs("EmailAccountUserName")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>邮件帐号密码：</TD>
				<TD><input name="txt_EmailAccountUserPWD" type="text" class="input" value="<%=rs("EmailAccountUserPWD")%>" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			</TABLE>
		</DIV>
	</DIV>
	
	<DIV class="tagContent selectTag content" id="tagContent2" style="display:none;">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">是否显示留言板列表：</TD>
				<TD>
				<input type="radio" name="txt_IsShowguestbookList" value="1" <%if rs("IsShowguestbookList")=1 then response.Write "checked"%>> 显示 &nbsp;
				<input type="radio" name="txt_IsShowguestbookList" value="0" <%if rs("IsShowguestbookList")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>留言本状态：</TD>
				<TD>
				<input type="radio" name="txt_guestbookStatus" value="1" <%if rs("guestbookStatus")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_guestbookStatus" value="0" <%if rs("guestbookStatus")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>留言本脏话过滤：</TD>
				<TD>
				<input type="radio" name="txt_guestbookSplit" value="1" <%if rs("guestbookSplit")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_guestbookSplit" value="0" <%if rs("guestbookSplit")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>留言本审核功能：</TD>
				<TD>
				<input type="radio" name="txt_guestbookAudit" value="1" <%if rs("guestbookAudit")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_guestbookAudit" value="0" <%if rs("guestbookAudit")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>留言验证码：</TD>
				<TD>
				<input type="radio" name="txt_guestbookValidation" value="1" <%if rs("guestbookValidation")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_guestbookValidation" value="0" <%if rs("guestbookValidation")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>留言关键字过滤：</TD>
				<TD><textarea name="txt_guestSplitWord" cols="60" rows="5" class="input"><%=rs("guestSplitWord")%></textarea><span class="huitext">&nbsp;必填</span></TD>
			  </TR>			  
			</TABLE>
		</DIV>
	</DIV>
	
	<DIV class="tagContent selectTag content" id="tagContent3" style="display:none;">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">留言类别(数字)：</TD>
				<TD><input type="text" name="txt_guestbookCategoryID" class="input" value="<%=rs("guestbookCategoryID")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>	 
			  <TR class="content-td1">
				<TD>广告类别(数字)：</TD>
				<TD><input type="text" name="txt_AdsCategoryID" class="input" value="<%=rs("AdsCategoryID")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>	 
			  <TR class="content-td1">
				<TD>友情链接类别(数字)：</TD>
				<TD><input type="text" name="txt_FriendLinkCategoryID" class="input" value="<%=rs("FriendLinkCategoryID")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>	 
			  <TR class="content-td1">
				<TD>人才招聘部门类别(数字)：</TD>
				<TD><input type="text" name="txt_JobDepartCategoryID" class="input" value="<%=rs("JobDepartCategoryID")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>	 
			</TABLE>
		</DIV>
	</DIV>
	
	<DIV class="tagContent selectTag content" id="tagContent4" style="display:none;">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">版权信息：</TD>
				<TD><textarea name="txt_copyright" cols="80" rows="10" class="input"><%=rs("copyright")%></textarea><span class="huitext">&nbsp;必填</span><br>
				<span class="redtext">连续书写信息不需要回车，如需要用"#"代替</span>
				</TD>
			  </TR>	 
			  <TR class="content-td1">
				<TD width="15%">版权信息(英文版)：</TD>
				<TD><textarea name="txt_copyright_en" cols="80" rows="10" class="input"><%=rs("copyright_en")%></textarea><span class="huitext">&nbsp;必填</span></TD>
			  </TR>	 
			</TABLE>
		</DIV>
	</DIV>
	
	<DIV class="tagContent selectTag content" id="tagContent5" style="display:none;">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">首页静态化：</TD>
				<TD>
				<input type="radio" name="txt_IsIndexStatic" value="1" <%if rs("IsIndexStatic")=1 then response.Write "checked"%>> 是 &nbsp;
				<input type="radio" name="txt_IsIndexStatic" value="0" <%if rs("IsIndexStatic")=0 then response.Write "checked"%>> 否 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>	 
			  <TR class="content-td1">
				<TD>首页模板文件：</TD>
				<TD><input name="txt_IndexTemplate" type="text" class="input" value="<%=rs("IndexTemplate")%>" size="80">
				&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_IndexTemplate');">选择</a></TD>
			  </TR>	 
			</TABLE>
		</DIV>
	</DIV>
	
	
	<DIV class="tagContent selectTag content" id="tagContent6" style="display:none;">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			<TR class="content-td1">
				<TD width="15%">是否开启百度留言版：</TD>
				<TD>
				<input type="radio" name="txt_isbaiduguestbookJS" value="1" <%if rs("isbaiduguestbookJS")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_isbaiduguestbookJS" value="0" <%if rs("isbaiduguestbookJS")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>百度留言板JS：</TD>
				<TD><input name="txt_baiduguestbookJS" type="text" class="input" value="<%=rs("baiduguestbookJS")%>" size="80">
			    <span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否开启站点统计：</TD>
				<TD>
				<input type="radio" name="txt_isSiteTotal" value="1" <%if rs("isSiteTotal")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_isSiteTotal" value="0" <%if rs("isSiteTotal")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>站点统计JS：</TD>
				<TD><input name="txt_SiteTotalJS" type="text" class="input" value="<%=rs("SiteTotalJS")%>" size="80">
			    <span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>单页内容编辑器：</TD>
				<TD>
				<input type="radio" name="txt_SinglePageEditor" value="0" <%if rs("SinglePageEditor")=0 then response.Write "checked"%>> 
				简单js编辑器 &nbsp;
				<input type="radio" name="txt_SinglePageEditor" value="1" <%if rs("SinglePageEditor")=1 then response.Write "checked"%>> 
				ewebeditor编辑器 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>数据库路径：</TD>
				<TD><input name="txt_DatabasePath" size="60" class="input" value="<%=RS("DatabasePath")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>数据库备份文件夹：</TD>
				<TD><input name="txt_DataBackForder" size="60" class="input" value="<%=RS("DataBackForder")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>允许上传文件种类：</TD>
				<TD><input name="txt_UPloadFileAllowEXT" type="text" class="input" value="<%=rs("UPloadFileAllowEXT")%>" size="80">
			    <span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否开启记录上传文件：</TD>
				<TD>
				<input type="radio" name="txt_isRecordUploadFile" value="1" <%if rs("isRecordUploadFile")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_isRecordUploadFile" value="0" <%if rs("isRecordUploadFile")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否启用Ip禁止：</TD>
				<TD>
				<input type="radio" name="txt_IsprohibitIP" value="1" <%if rs("IsprohibitIP")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_IsprohibitIP" value="0" <%if rs("IsprohibitIP")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>		
			  <TR class="content-td1">
				<TD>交互性Ip禁止列：</TD>
				<TD><textarea name="txt_prohibitIP" cols="60" rows="5" class="input"><%=rs("prohibitIP")%></textarea>&nbsp;IP之间#隔开<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>订单类型：</TD>
				<TD>
				<input type="radio" name="txt_IsShowOrderItem" value="1" <%if rs("IsShowOrderItem")=1 then response.Write "checked"%>> 显示订单详细(需要用到购物车) &nbsp;
				<input type="radio" name="txt_IsShowOrderItem" value="0" <%if rs("IsShowOrderItem")=0 then response.Write "checked"%>> 直接订单 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否启用城市：</TD>
				<TD>
				<input type="radio" name="txt_isShowCity" value="1" <%if rs("isShowCity")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_isShowCity" value="0" <%if rs("isShowCity")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否启用上传密码：</TD>
				<TD>
				<input type="radio" name="txt_isNeedUploadPassword" value="1" <%if rs("isNeedUploadPassword")=1 then response.Write "checked"%>> 正常 &nbsp;
				<input type="radio" name="txt_isNeedUploadPassword" value="0" <%if rs("isNeedUploadPassword")=0 then response.Write "checked"%>> 关闭 &nbsp;
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>上传密码：</TD>
				<TD>
				<input name="txt_UploadPassword" size="60" class="input" value="<%=RS("UploadPassword")%>">
				<span class="huitext">&nbsp;必填</span>
				</TD>
			  </TR>
			</TABLE>
		</DIV>
	</DIV>
	<div class="divpadding">
	  <input name="btnsearch" value="设置变量" class="button" type="submit">
	</div>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
%>
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
