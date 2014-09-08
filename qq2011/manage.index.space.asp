<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
	Function ShowSize(objsize)
		dim sizeExt
		if objsize<1024 then 		   
 		   ShowSize=objsize&"&nbsp;Byte"
 		end if
 		if objsize>1024 then
 		   objsize=(objsize/1024)
 		   ShowSize=objsize&"&nbsp;KB"
 		end if
 		if objsize>1024 then
 		   objsize=(objsize/1024)
 		   ShowSize=formatnumber(objsize,2) & "&nbsp;MB"		
 		end if
 		if objsize>1024 then
 		   objsize=(objsize/1024)
 		   ShowSize=formatnumber(objsize,2) & "&nbsp;GB"	   
 		end if   
 		ShowSize="<font face=verdana>" & ShowSize & "</font>"		
 	end Function	
	Function GetTotalSize()
		dim fso,d,cmsRootPath
 		set fso=server.createobject("scripting.filesystemobject")
 		cmsRootPath=server.mappath("../")
 		set d=fso.getfolder(cmsRootPath)
 		GetTotalSize=d.size
		set d=nothing
		set fso=nothing
	end Function
	
	Function showFolderSize(folderpath)
		dim fso,d,foldersize,barsize
 		set fso=server.createobject("scripting.filesystemobject") 		
 		set d=fso.getfolder(server.mappath(folderpath))
 		foldersize=d.size 		
 		showFolderSize=ShowSize(foldersize)
		set d=nothing
		set fso=nothing
	end function 	
 	
 	Function Drawbar(drvpath)
		response.Write server.mappath(drvpath)
 	End Function 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">网站空间查看</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
		<table cellspacing="1" width="100%">			 	
			<tr>
				<td height="25" class="content-td1">·数据库占用空间：&nbsp;<img src="images/bar.gif" width=<%=drawbar("../tot_dabase")%> height=10>&nbsp;/Sp_dabase:<%=showFolderSize("../Sp_data")%></td>
			</tr>
			<tr>
				<td height="25" class="content-td1">·系统图片占用空间：&nbsp;<img src="images/bar.gif" width=<%=drawbar("../images")%> height=10>&nbsp;/images:<%=showFolderSize("../images")%></td>
			</tr>
			<tr>
				<td height="25" class="content-td1">·上传图片占用空间：&nbsp;<img src="images/bar.gif" width=<%=drawbar("../upload")%> height=10>&nbsp;/pic:<%=showFolderSize("../upload")%></td>
			</tr>
			<tr>
				<td height="25" class="content-td1">·系统占用空间总计：<img src="images/bar.gif" width=400 height=10> <%=ShowSize(GetTotalSize())%></td>
			</tr>
		</table>
	  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
