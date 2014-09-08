<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim Db_path:Db_path = VerificationUrlParam("Db_path","string","")
	if action<>"" and action="del" then
		''删除信息
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set a = fso.GetFile(server.MapPath(config_DataBackForder&Db_path&""))
		a.Delete
		ALert "删除成功","plug.manage.database.back.asp"
	elseif action="backup" then
		'''''备份
		Dim filename:filename=date()
		filename=filename&"-"
		filename=filename&time()
		'filename=replace(filename,"-","")
		filename=replace(filename,":","-")
		filename=replace(filename," ","-")
		''''文件路径
		Set Fso=server.createobject("scripting.filesystemobject") 
		folderpath=Server.MapPath(config_DatabasePath)
		fso.CopyFile folderpath,Server.MapPath(config_DataBackForder)&"\"&filename&".bak"
		ALert "备份成功","plug.manage.database.back.asp"
	elseif action="restore" then
		Set Fso=server.createobject("scripting.filesystemobject") 
		folderpath=server.MapPath(config_DataBackForder&Db_path&"")
		fso.copyfile folderpath,Server.MapPath(config_DatabasePath)
		ALert "还原成功","plug.manage.database.back.asp"
	elseif action="compact" then
		CompactDB Server.MapPath(config_DatabasePath),true
		ALert "压缩成功","plug.manage.database.back.asp" 
	end if	
	'response.Write date()
	
	Function CompactDB(dbPath, boolIs97)
		Dim fso, Engine, strDBPath
		strDBPath = left(dbPath,instrrev(DBPath,"\"))
		Set fso = createObject("Scripting.FileSystemObject")		
		If fso.FileExists(dbPath) Then
		Set Engine = createObject("JRO.JetEngine")
				
		If boolIs97 = "True" Then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
			& "Jet OLEDB:Engine Type=" & JET_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
		End If
		fso.CopyFile strDBPath & "temp.mdb",dbpath
		fso.deleteFile(strDBPath & "temp.mdb")
		Set fso = nothing
		Set Engine = nothing
		CompactDB = "你的数据库, " & dbpath & ",已经被压缩" & vbCrLf
		Else
		CompactDB = "没有发现数据库或者路径不对. 再来了" & vbCrLf
		End If
	End Function

	''''''''''''''''''
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">数据库备份还原</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<div class="divpadding">
			  <input name="btbackup" onclick="javascript:if(confirm('SP_CMS系统提示\r\n\n\n确定备份?'))location.href='?action=backup';" value="备份数据库" class="button" type="button">&nbsp;&nbsp;<input name="btbackup" onclick="javascript:if(confirm('SP_CMS系统提示\r\n\n\n压缩数据库?'))location.href='?action=compact';" value="压缩数据库" class="button" type="button">
			</div>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="30%" align="left">备份数据库名称</TH>
			    <TH width="30%" align="left">数据库大小</TH>
			    <TH width="20%" align="left">备份日期</TH>
			    <TH width="15%" align="left"></TH>
			  </TR>
			  <%
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set f = fso.GetFolder(server.MapPath(config_DataBackForder))
				Set fc = f.Files
				For Each f1 in fc
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=f1.name%>"></TD>
				<TD><%=f1.name%></TD>
				<TD><%=f1.Size%>byte</TD>
				<TD><%=f1.DateCreated%></TD>
				<TD><a href="?action=restore&Db_path=<%=f1.name%>">还原</a>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&Db_path=<%=f1.name%>';">删除</a></TD>
			  </TR>
			  <%
				Next
			  %>
			</TABLE>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
