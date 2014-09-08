<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
		dim adsdisplayID:adsdisplayID = F.ReplaceFormText("txt_adsdisplayID")
		dim adsname:adsname = F.ReplaceFormText("txt_adsname")
		dim adspicurl:adspicurl = F.ReplaceFormText("txt_adspicurl")
		dim adsurl:adsurl = F.ReplaceFormText("txt_adsurl")
		dim adswidth:adswidth = F.ReplaceFormText("txt_adswidth")
		dim adsheight:adsheight = F.ReplaceFormText("txt_adsheight")
		dim adstitle:adstitle = F.ReplaceFormText("txt_adstitle")
		dim starttime:starttime = F.ReplaceFormText("txt_starttime")
		dim endtime:endtime = F.ReplaceFormText("txt_endtime")
		dim IsActive:IsActive = F.ReplaceFormText("txt_IsActive")
		dim memo:memo = F.ReplaceFormText("txt_memo")
		if IsActive="" then IsActive = 0
		'''''''增加记录
		Con.execute("insert into Sp_ads (adsdisplayID,adsname,adspicurl,adsurl,adswidth,adsheight,adstitle,starttime,endtime,IsActive,[memo]) values ("&adsdisplayID&",'"&adsname&"','"&adspicurl&"','"&adsurl&"',"&adswidth&","&adsheight&",'"&adstitle&"','"&starttime&"','"&endtime&"',"&IsActive&",'"&memo&"')")
		Alert "增加广告成功","pulg.ads.list.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/calender.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_adsname").value=="")
			{
				alert("Sp_CMS提示\r\n\n广告名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_adspicurl").value=="")
			{
				alert("Sp_CMS提示\r\n\n广告元素必须填写!");
				return false;			
			}
			if(document.getElementById("txt_adsurl").value=="")
			{
				alert("Sp_CMS提示\r\n\n站点链接必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">增加广告</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">显示类型：</TD>
				<TD>
				<select name="txt_adsdisplayID">
				<%
					set Rs1 = Con.Query("select id,categoryname from Sp_dictionary where categoryID = "&config_AdsCategoryID&"")
					if rs1.recordcount<>0 then
					do while not rs1.eof
					response.Write "<option value="&rs1("id")&">"&rs1("categoryname")&"</option>"
					rs1.movenext
					loop
					end if
				%>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>广告名称：</TD>
				<TD><input name="txt_adsname" type="text" class="input" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>广告元素<span class="huitext">(图片,flash,文字)</span>：</TD>
				<TD>
				<input name="txt_adspicurl" type="text" class="input" size="60"><span class="huitext">&nbsp;必填</span>
				<iframe id="uploadfile" name="uploadfile" width="100%" height="30px;" frameborder="0" scrolling="no" src="File_Upload.asp?obj=txt_adspicurl"></iframe>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>站点链接：</TD>
				<TD><input name="txt_adsurl" type="text" class="input" size="60"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>提示文字：</TD>
				<TD><input name="txt_adstitle" type="text" class="input" size="80"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>广告规格：</TD>
				<TD>
				<input name="txt_adswidth" type="text" class="input" size="5">
				&nbsp;×&nbsp;
				<input name="txt_adsheight" type="text" class="input" size="5"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>广告开始时间：</TD>
				<TD><input name="txt_starttime" type="text" class="input" onfocus="javascript:HS_setDate(this);"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>广告结束时间：</TD>
				<TD><input name="txt_endtime" type="text" class="input" onfocus="javascript:HS_setDate(this);"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>备注：</TD>
				<TD><textarea name="txt_memo" cols="80" rows="6"></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否激活：</TD>
				<TD><input name="txt_IsActive" class="input" type="checkbox" value="1"></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="增加广告" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>