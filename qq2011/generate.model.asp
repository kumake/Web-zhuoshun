<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
'modelID 	ģ��ID
'modelTable ģ�ͱ�����
'classid 	����ģ��������ID
'retype	 	�������� 0����ʱ������ 1����ID����
'start	 	��ʼ��־ҳ��startֵС��endid ʱ��ִֹͣ��
'startday  	��ʼʱ��
'endday	 	��ʼ����
'startid  	��ʼID
'endid		����ID
'havehtml	�Ƿ��Ѿ�����Html
Dim modelID:modelID = VerificationUrlParam("modelID","int","0")
Dim modelTable:modelTable = VerificationUrlParam("modelTable","string","")
Dim classid:classid = VerificationUrlParam("classid","int","0")
Dim retype:retype = VerificationUrlParam("retype","int","1")
Dim start:start = VerificationUrlParam("start","string","0")
Dim startday:startday = VerificationUrlParam("startday","string","")
Dim endday:endday = VerificationUrlParam("endday","string","")
Dim startid:startid = VerificationUrlParam("startid","int","0")
Dim endid:endid = VerificationUrlParam("endid","int","10")
Dim havehtml:havehtml = VerificationUrlParam("havehtml","int","0")

function ReNewsHtml(Byval modelID,Byval modelTable,Byval classid,Byval retype,Byval start,Byval startday,Byval endday,Byval startid,Byval endid,Byval havehtml)
	if retype=1 then '��IDΪ��������HTML
		if start<>"" and isnumeric(start) then
			if Cint(start)<Cint(endid) then
				response.Write "���ڽ���IDΪ "& start &" ����Ϣ����HTML��"
				''''��ʼִ�к�������
				generateHtm modelID,classid,startid
				'response.End()
				response.Redirect "generate.asp?modelID="&modelID&"&classid="&classid&"&retype="&retype&"&start="&(start+1)&"&startday="&startday&"&endday="&endday&"&startid="&startid&"&endid="&endid&"&havehtml="& havehtml
			else
				response.Write "<span style='font-size:12px;color:#FF0000;'>���ɳɹ���</span>"
			end if
		end if
	end if
end function

function generateHtm(Byval modelID,Byval modelTable,Byval classid,Byval ItemID)

end function

'''''ִ�д���
ReNewsHtml modelID,modelTable,classid,retype,start,startday,endday,startid,endid,havehtml
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
