<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
'modelID 	ģ��ID
'classid 	����ģ��������ID
'retype	 	�������� 0����ʱ������ 1����ID����
'start	 	��ʼ��־ҳ��startֵС��endid ʱ��ִֹͣ��
'startday  	��ʼʱ��
'endday	 	��ʼ����
'startid  	��ʼID
'endid		����ID
'havehtml	�Ƿ��Ѿ�����Html
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
	<TITLE>��̨��ҳ</TITLE>
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
		
		''''��������
		application.Lock()
		application("TOT_Reset_Total")=100
		application.UnLock()


		if(Cint(id)>Cint(endId)) then
			response.Write("<br><br><br><br><br>�ɹ�����"&application("TOT_Reset_Total")&"�����ż�¼,id��"&startId&"��"&endId&"")
			response.Write("&nbsp;&nbsp;<a href='javascript:void(0);' onclick='javascript:parent.ClosePop();' style='color:red; font-size:12px;'>�رձ�����,���·���</a>")
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">���ݸ�����</A></LI>
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
			<Br>���ڷ������ţ���<%=application("TOT_Reset_Total")%>��&nbsp;&nbsp;��<%=application("TOT_Reset_Robot")%>��&nbsp;&nbsp;<span style="font-size:12px; color:#FF0000;" id="P_num"></span>
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
        if (HTTP.readyState == 4) // �յ������ķ�������Ӧ
		{
			if (HTTP.status == 500) {//����������				
				document.write("�����������ڲ�����");				

            }
            else if (HTTP.status == 200) {//HTTP��������Ӧ��ֵOK
				//document.write(HTTP.responseText);//�����������ص��ַ���д��ҳ����IDΪmessage������
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
