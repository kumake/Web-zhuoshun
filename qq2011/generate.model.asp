<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
'modelID 	模型ID
'modelTable 模型表名称
'classid 	所属模型中类别的ID
'retype	 	生成类型 0：按时间区间 1：按ID区间
'start	 	开始标志页当start值小于endid 时候停止执行
'startday  	开始时间
'endday	 	开始日期
'startid  	开始ID
'endid		结束ID
'havehtml	是否已经生成Html
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
	if retype=1 then '以ID为区间生成HTML
		if start<>"" and isnumeric(start) then
			if Cint(start)<Cint(endid) then
				response.Write "正在进行ID为 "& start &" 的信息生成HTML。"
				''''开始执行函数生成
				generateHtm modelID,classid,startid
				'response.End()
				response.Redirect "generate.asp?modelID="&modelID&"&classid="&classid&"&retype="&retype&"&start="&(start+1)&"&startday="&startday&"&endday="&endday&"&startid="&startid&"&endid="&endid&"&havehtml="& havehtml
			else
				response.Write "<span style='font-size:12px;color:#FF0000;'>生成成功！</span>"
			end if
		end if
	end if
end function

function generateHtm(Byval modelID,Byval modelTable,Byval classid,Byval ItemID)

end function

'''''执行代码
ReNewsHtml modelID,modelTable,classid,retype,start,startday,endday,startid,endid,havehtml
%>
<HTML>
<HEAD>
	<TITLE>后台首页</TITLE>
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
