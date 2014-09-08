<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
'modelID 	模型ID
'classid 	所属模型中类别的ID
'retype	 	生成类型 0：按时间区间 1：按ID区间
'start	 	开始标志页当start值小于endid 时候停止执行
'startday  	开始时间
'endday	 	开始日期
'startid  	开始ID
'endid		结束ID
'havehtml	是否已经生成Html
'Dim GenerateType:GenerateType = VerificationUrlParam("GenerateType","string","")
'Dim modelID:modelID = VerificationUrlParam("modelID","int","0")
'Dim modelTable:modelTable = VerificationUrlParam("modelTable","string","")
'Dim classid:classid = VerificationUrlParam("classid","int","0")
'Dim retype:retype = VerificationUrlParam("retype","int","1")
'Dim start:start = VerificationUrlParam("start","string","0")
'Dim startday:startday = VerificationUrlParam("startday","string","")
'Dim endday:endday = VerificationUrlParam("endday","string","")
'Dim startid:startid = VerificationUrlParam("startid","int","0")
'Dim endid:endid = VerificationUrlParam("endid","int","10")
'Dim havehtml:havehtml = VerificationUrlParam("havehtml","int","0")
%>

<HTML>
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
</HEAD>
<%
'''''''''''''''''''''''''''''''''''''''''''''''
		dim startId:startId=1
		dim endId:endId=100
		dim starts:starts =0 
		
		if not isempty(request("id")) then
			id=Clng(request("id"))
		else
			id=1
		end if
		
		if not isempty(request("start")) then
			starts=Cint(request("start"))
		else
			starts=0
		end if
		''''
		if(application("TOT_Reset_Robot")="") then
			application.Lock()
			application("TOT_Reset_Robot")=0
			application.UnLock()
		end if
		if(application("TOT_Reset_Total")="") then
			application.Lock()
			application("TOT_Reset_Total")=0
			application.UnLock()
		end if
		
		if Cint(starts)=1 then
			application.Lock()
			application("TOT_Reset_Robot")=0
			application.UnLock()
		end if
		
		''''结算总数
		application.Lock()
		application("TOT_Reset_Total")=100
		application.UnLock()


		if(Cint(id)>Cint(endId)) then
			response.Write("<br><br><br><br><br>成功发布"&application("TOT_Reset_Total")&"条新闻记录,id从"&startId&"到"&endId&"")
			response.Write("&nbsp;&nbsp;<a href='javascript:void(0);' onclick='javascript:parent.ClosePop();' style='color:red; font-size:12px;'>关闭本窗口,重新发布</a>")
			response.End()
		else
			id = id+1
			application.Lock()
			application("TOT_Reset_Robot")=application("TOT_Reset_Robot")+1
			application.UnLock()
		end if
		percentNum=Cint((application("TOT_Reset_Robot")*100)/application("TOT_Reset_Total"))
%>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">数据更新中</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV>
	  <DIV>
		<br>
		<br>
			<table width="400" height="13" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #999999;">
			  <tr>
				<td style="padding-left:1px; padding-right:1px;">
				  <table id="stat_w" width="1%"  border="0" align="left" cellpadding="0" cellspacing="0">
				  <tr>
					<td style="background-image:url(images/loadingbg.gif); height:13px; background-repeat:repeat-x;"></td>
				  </tr>
				</table>
				
			   </td>
			  </tr>
		  </table>
			<Br>正在发布新闻，共<%=application("TOT_Reset_Total")%>条&nbsp;&nbsp;第<%=application("TOT_Reset_Robot")%>条&nbsp;&nbsp;<span style="font-size:12px; color:#FF0000;" id="P_num"></span>
		<br>
		<br>
	  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
<script type="text/javascript" language="javascript">
var HTTP;
function CallServer(Url)
{
	HTTP = new ActiveXObject("Microsoft.XMLHTTP");
	HTTP.onreadystatechange = doAction;
	var ReturnValue=HTTP.open("POST", Url, false);
	HTTP.send("");
	return HTTP.responseText;
}
function doAction() {		
        if (HTTP.readyState == 4) // 收到完整的服务器响应
		{
			if (HTTP.status == 500) {//服务器错误				
				document.write("服务器出现内部错误。");				

            }
            else if (HTTP.status == 200) {//HTTP服务器响应的值OK
				//document.write(HTTP.responseText);//将服务器返回的字符串写到页面中ID为message的区域
				//document.write("id:"+<%=id%>+"<br>");
				document.write('<meta http-equiv="refresh" content="1;URL=manage.generate.asp?id=<%=id%>">');
				document.getElementById('stat_w').width ="<%=percentNum%>%";
				document.getElementById('P_num').innerHTML ="<%=percentNum%>%";
            } 
			else {
                alert(HTTP.status);
            }
        }
}
function resetHtml()
{
	var ReturnValue;	
	ReturnValue=CallServer('manage.generatehtml.asp?id=<%=id%>');
	document.write(ReturnValue);
}
resetHtml();
</script>
