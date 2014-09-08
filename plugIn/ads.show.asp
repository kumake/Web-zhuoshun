<!--#include file="../Sp_inc/conn.asp"-->

<%
	Dim AD_ID
	Dim AD_adsdisplayID
	Dim AD_adsname
	Dim AD_adspicurl
	Dim AD_adsurl
	Dim AD_adswidth
	Dim AD_adsheight
	Dim AD_adstitle
	Dim AD_starttime
	Dim AD_endtime
	Dim AD_IsActive
	Dim ScrCode
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if ItemID =0 then 
		response.write "document.write('没有找到您要浏览的广告');"
	else
		set rs = Con.Query("select * from Sp_Ads where ID="&ItemID&"")
		
		if rs.recordcount<>0 then
			AD_ID = rs("ID")
			AD_adsdisplayID = rs("adsdisplayID")
			AD_adsname = rs("adsname")
			AD_adspicurl = rs("adspicurl")
			AD_adsurl = rs("adsurl")
			AD_adswidth = rs("adswidth")
			AD_adsheight = rs("adsheight")
			AD_adstitle = rs("adstitle")
			AD_starttime = rs("starttime")
			AD_endtime = rs("endtime")
			AD_IsActive = rs("IsActive")
			''''连接
			ScrCode = "<a href='"&AD_adsurl&"'>"
			
			If InStr(1,AD_adspicurl,".swf",1)>0 Then
				ScrCode=ScrCode&"<EMBED src='"& AD_adspicurl &"' quality=high WIDTH='"& AD_adswidth &"' HEIGHT='"& AD_adsheight &"' TYPE='application/x-shockwave-flash' PLUGINSPAGE='http://www.macromedia.com/go/getflashplayer'></EMBED>"
			Else
				'response.Write rs.recordcount
				ScrCode=ScrCode&"<img src='"& AD_adspicurl &"' border='0' width="& AD_adswidth &" height="& AD_adsheight &" alt='"& AD_adstitle &"' align='top'>"
			End If
			'''''''''''''''''''''''''''''''''
			''''
			ScrCode=ScrCode&"</a>"
			'''''''''''''''''''''''''''''''''
			response.write "document.write("""& ScrCode &""");"&vbCrlf
		else
			response.write "document.write('没有找到您要浏览的广告');"
		end if
	end if
%>