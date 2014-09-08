<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">google 百度地图</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
			<%
			'''''''''''''''''''''''''''''''''''''''''''''''
			'''''''''''''''''''''''''''''''''''''''''''''''
			Dim strBuffer
			strBuffer="<?xml version=""1.0"" encoding=""UTF-8"" ?> "&chr(10)
			strBuffer=strBuffer&"<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">"&chr(10)
			
			'if(isarray(classArr)) then
				'for i=0 to ubound(classArr,2)
					'if(classArr(5,i)=true) then
						'pri="1.0000"
						'linkstr="/index.htm"
					'else
						'pri="0.5000"
						'linkstr="/index_1.htm"
					'end if
					'strBuffer=strBuffer&"<url>"&chr(10)
					'strBuffer=strBuffer&"<loc>"&SiteUrl&classArr(4,i)&linkstr&"</loc>"&chr(10)
					'strBuffer=strBuffer&"<lastmod>"&classArr(6,i)&"</lastmod>"&chr(10)
					'strBuffer=strBuffer&"<priority>"&pri&"</priority>"&chr(10)
					'strBuffer=strBuffer&"<changefreq>daily</changefreq>"&chr(10)
					'strBuffer=strBuffer&"</url>"&chr(10)
				'next
			'end if
			'strBuffer=strBuffer&"</urlset>"&chr(10)
			'set classArr=nothing
			'生成sitemap.xml
			'Set fs=Server.CreateObject("Scripting.FileSystemObject")
			'Set file_w = fs.CreateTextFile(server.mappath("../sitemap.xml"),true)	
			'file_w.WriteLine(strBuffer)
			'file_w.close
			'set file_w=nothing
			'set fs=nothing
			response.Write("<br><p align=""center"">成功生成sitemap.xml</p>")
			%>
	  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
