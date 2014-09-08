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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">服务器配置</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
			<TABLE cellSpacing="1" width="100%">
			<tr>
			<th colspan="2">服务器信息</th>
			</tr>
			<tr class="content-td1">
			<td height="32">　服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
			<td height="32">　站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
			</tr>
			<tr class="content-td1">
			<td width="48%" height="32">　返回服务器的主机名，IP地址：<font color=#0076AE><%=Request.ServerVariables("SERVER_NAME")%></td>
			<td width="52%" height="32">　服务器操作系统：<font color=#0076AE><%=Request.ServerVariables("OS")%></td>
			</tr>
			<tr class="content-td1">
			<td width="48%" height="32">　脚本解释引擎<span class="small2">：</span><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>　</td>
			<td width="52%" height="37">　<span class="small2">WEB</span>服务器的名称和版本：<font color=#0076AE><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
			</tr>
			<tr class="content-td1">
			<td width="48%" height="32">　脚本超时时间<span class="small2">：</span><font color=#0076AE><%=Server.ScriptTimeout%> 秒</td>
			<td width="52%" height="32">　CDONTS组件支持<span class="small2">：</span>
			<%
			On Error Resume Next
			Server.CreateObject("CDONTS.NewMail")
			if err=0 then 
			response.write("<font color=#0076AE>√")
			else
			response.write("<font color=red>×")
			end if	 
			err=0
			%>
			</td>
			</tr>
			<tr class="content-td1">
			<td width="48%" height="32">　虚拟路径：<%=Request.ServerVariables("SCRIPT_NAME")%></td>
			<td width="52%" height="32">　<span class="small2">Jmail</span>邮箱组件支持<span class="small2">：</span>
			  <%If Not IsObjInstalled(theInstalledObjects(13)) Then%>
			  <font color="red">×
			  <%else%>
			  <font color="0076AE"> √
			  <%end if%>
			</td>
			</tr>
			<tr class="content-td1">
			<td height="32">　返回服务器处理请求的端口：<%=Request.ServerVariables("SERVER_PORT")%></td>
			<td height="32">　协议的名称和版本：<%=Request.ServerVariables("SERVER_PROTOCOL")%></td>
			</tr>
			<tr class="content-td1">
			<td height="32">　服务器 CPU 数量：<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个　</td>
			<td height="32">　客户端操作系统：
			<%
			dim thesoft,vOS
			thesoft=Request.ServerVariables("HTTP_USER_AGENT")
			if instr(thesoft,"Windows NT 5.0") then
			vOS="Windows 2000"
			elseif instr(thesoft,"Windows NT 5.2") then
			vOs="Windows 2003"
			elseif instr(thesoft,"Windows NT 5.1") then
			vOs="Windows XP"
			elseif instr(thesoft,"Windows NT") then
			vOs="Windows NT"
			elseif instr(thesoft,"Windows 9") then
			vOs="Windows 9x"
			elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
			vOs="类Unix"
			elseif instr(thesoft,"Mac") then
			vOs="Mac"
			else
			vOs="Other"
			end if
			response.Write(vOs)
			%>
			</td>
			<tr>
			<th colspan="2">本系统需要的软件环境：</th>
			</tr>
			</tr>
			<tr class="content-td1">
			<td width="48%" height="25">　JRO.JetEngine(ACCESS&nbsp;  数据库<span class="small2">)：</span>
			<%
			On Error Resume Next
			Server.CreateObject("JRO.JetEngine")
			if err=0 then 
			response.write("<font color=#0076AE>√")
			else
			response.write("<font color=red>×")
			end if	 
			err=0
			%>                  </td>
			<td width="52%" height="25">　数据库使用<span class="small2">：</span>
			<%
			On Error Resume Next
			Server.CreateObject("adodb.connection")
			if err=0 then 
			response.write("<font color=#0076AE>√,可以使用本系统")
			else
			response.write("<font color=red>×,不能使用本系统")
			end if	 
			err=0
			%>                  </td>
			</tr>
			<tr class="content-td1">
			<td height="25">　<span class="small2">FSO</span>文本文件读写<span class="small2">：</span>
			<%
			On Error Resume Next
			Server.CreateObject("Scripting.FileSystemObject")
			if err=0 then 
			response.write("<font color=#0076AE>√,可以使用本系统")
			else
			response.write("<font color=red>×，不能使用此系统")
			end if	 
			err=0
			%>                  </td>
			<td height="25">　Microsoft.XMLHTTP
			  <%If Not IsObjInstalled("Microsoft.XMLHTTP") Then%>
			  <font color="red">×，不能使用此系统
			  <%else%>
			  <font color="0076AE"> √,可以使用本系统
			  <%end if%>
			 　Adodb.Stream
			<%If Not IsObjInstalled("Adodb.Stream") Then%>
			<font color="red">×
			<%else%>
			<font color="0076AE"> √
			<%end if%>                  </td>
			</tr>                
			</table>			  
		  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
