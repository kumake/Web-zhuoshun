<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<TITLE>上传文件管理</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
	<script language="javascript" type="text/javascript">
		function inserturlimg(obj)
		{
			if(obj.value=="")
			{
				alert("SP_CMS系统提示:\t\n\n\n请输入要插入的图片路径!");
			}
			else
			{
				 var code = obj.value.split(".");
				 var ext = code[code.length-1].toLowerCase();
				 var allowstr = "<%=config_UPloadFileAllowEXT%>";
				 if(allowstr.indexOf(ext)==-1)
				 {
				 	alert("SP_CMS系统提示:\t\n\n\n格式不正确!\r\n\n");
				 }
				 else
				 {
					window.returnValue=obj.value;
					window.close();
				 }
			}
		}
	</script>
</head>
<body>
<form action="" method="post">
<DIV id=wrap>
	<UL id=tags>
	  <LI class="selectTag" style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="#">添加网络图片</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent1',this)" href="#">本地上传</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent2',this)" href="#">网络硬盘</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<DIV>
			<input name="imgUrl" size="50" class="input" type="text">&nbsp;<input name="btnInsertImgUrl" class="button" value="插入图片" onClick="javascript:inserturlimg(document.getElementById('imgUrl'));" type="button">
		</DIV>
	</DIV>
	<DIV class="tagContent selectTag content" id="tagContent1" style="display:none;">
		<DIV>
			<TABLE cellSpacing=1 width="100%">
			<TR class="content-td1">
				<TD width="100%"><input name="txt_imgUrl" class="input" size="60" type="text">&nbsp;<input name="btnInsertImgUrl" class="button" value="插入图片" onClick="javascript:inserturlimg(document.getElementById('txt_imgUrl'));" type="button"></TD>
			</TR>
			<TR class="content-td1">
				<TD width="100%"><iframe name="ad" frameborder="0"   width="100%" height="30px;" scrolling="no" src="File_Upload.asp?obj=txt_imgUrl"></iframe></TD>
			</TR>
			</TABLE>
		</DIV>
	</DIV>
	<DIV class="tagContent selectTag content" id="tagContent2" style="display:none;">
		<DIV>
			功能开发中......
		</DIV>
	</DIV>
	</DIV>
</DIV>
</form>
</body>
</html>
