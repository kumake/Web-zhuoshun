<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim obj:obj = VerificationUrlParam("obj","string","")
	Dim path:path = VerificationUrlParam("path","string","")
	Dim Forderpath:Forderpath = Mid(path,InStrRev(path, "/")+1)
	Dim CurrentPath
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">选择文件</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="AllCheck"></TH>
			    <TH width="25%" align="left">文件名称</TH>
			    <TH width="25%" align="left">文件大小</TH>
			    <TH width="15%" align="left">最后修改日期</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
			    CurrentPath = "../upload"
			    if path <>"" then CurrentPath = path
			  	response.Write ShowFolderList(server.MapPath(CurrentPath))				
				Function ShowFolderList(folderspec)
					Dim fso, f, f1, fc, s,fd
					Set fso = CreateObject("Scripting.FileSystemObject")
					Set f = fso.GetFolder(folderspec)
					Set fc = f.Files
					'''''''''''''''''
					Set fd = f.SubFolders
					For Each f1 in fd
					  response.Write "<TR class='content-td1'>"
					  response.Write "<TD><input type='checkbox' name='ItemID'></TD>"
					  response.Write "<TD><img src='images/folder.gif' align='absmiddle'><a href='?obj="&obj&"&path=../upload/"&f1.name&"'>"&f1.name &"</a></TD>"
					  response.Write "<TD><span class='huitext'>文件夹</span></TD>"
					  response.Write "<TD>"&year(f1.DateLastModified)&"-"&month(f1.DateLastModified)&"-"&day(f1.DateLastModified)&"</TD>"
					  response.Write "<TD></TD>"
				      response.Write "</TR>"
					Next
					'''''
					For Each f1 in fc
					  response.Write "<TR class='content-td1'>"
					  response.Write "<TD><input type='checkbox' name='ItemID'></TD>"
					  response.Write "<TD><img src='images/file.gif' align='absmiddle'><a href='/upload/"&Forderpath&"/"&f1.name&"'>"&f1.name &"</a></TD>"
					  response.Write "<TD>"&f1.Size&"</TD>"
					  response.Write "<TD>"&f1.DateLastModified&"</TD>"
					  if Forderpath<>"" then
					  response.Write "<TD align='center'><a href='javascript:selectTemplate(""/upload/"&Forderpath&"/"&f1.name&""");'>选择</a></TD>"
					  else
					  response.Write "<TD align='center'><a href='javascript:selectTemplate(""/upload/"&f1.name&""");'>选择</a></TD>"
				      end if
					  response.Write "</TR>"
					Next
					''''
					ShowFolderList = s
				End Function
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
<script language="javascript" type="text/javascript">
    function selectTemplate(strPath)
    {
        var obj = "<%=obj%>";
        parent.document.getElementById(obj).value = strPath;
        parent.ClosePop();
    }
</script>