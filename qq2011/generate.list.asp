<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
'start	 	��ʼ��־ҳ��startֵС��endid ʱ��ִֹͣ��
'startid  	��ʼID
'endid		����ID
'havehtml	�Ƿ��Ѿ�����Html
Dim start:start = VerificationUrlParam("start","string","0")
Dim startid:startid = VerificationUrlParam("startid","int","0")
Dim endid:endid = VerificationUrlParam("endid","int","10")
Dim havehtml:havehtml = VerificationUrlParam("havehtml","int","0")

function ReJsHtml(Byval start,Byval startid,Byval endid,Byval havehtml)
end function

function generateHtm(Byval modelID,Byval ModelTable,Byval JsFileTemplate,Byval JsItemID)

end function

'''''ִ�д���
ReJsHtml start,startid,endid,havehtml
%>
<HTML>
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<style>
		td{ font-size:12px;}
		A {
			COLOR: #333; TEXT-DECORATION: none
		}
		A:hover {
			COLOR: #0079b8; TEXT-DECORATION: underline
		}
	</style>
</HEAD>
<BODY>
</BODY>
</HTML>
