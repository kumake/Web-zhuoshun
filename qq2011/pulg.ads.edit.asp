<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="save" then 
		dim adsdisplayID:adsdisplayID = F.ReplaceFormText("txt_adsdisplayID")
		dim adsname:adsname = F.ReplaceFormText("txt_adsname")
		dim adspicurl:adspicurl = F.ReplaceFormText("txt_adspicurl")
		dim adsurl:adsurl = F.ReplaceFormText("txt_adsurl")
		dim adswidth:adswidth = F.ReplaceFormText("txt_adswidth")
		dim adsheight:adsheight = F.ReplaceFormText("txt_adsheight")
		dim adstitle:adstitle = F.ReplaceFormText("txt_adstitle")
		dim starttime:starttime = F.ReplaceFormText("txt_starttime")
		dim endtime:endtime = F.ReplaceFormText("txt_endtime")
		dim IsActive:IsActive = F.ReplaceFormText("txt_IsActive")
		dim memo:memo = F.ReplaceFormText("txt_memo")
		if IsActive="" then IsActive = 0

		Con.execute("update Sp_ads set adsdisplayID="&adsdisplayID&",adsname='"&adsname&"',adspicurl='"&adspicurl&"',adsurl='"&adsurl&"',adswidth="&adswidth&",adsheight="&adsheight&",adstitle='"&adstitle&"',starttime='"&starttime&"',endtime='"&endtime&"',IsActive="&IsActive&",[memo]='"&memo&"' where ID="&ItemID&"")
		Alert "�޸Ĺ��ɹ�","pulg.ads.list.asp"
	end if
	if ItemID = 0 then
		Alert "��������ʧ��",""
	else
		set rs = Con.Query("select * from Sp_ads where ID="&ItemID&"")
		if rs.recordcount=0 then
			Alert "�Ҳ�����¼",""
		else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/calender.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_adsname").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n������Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_adspicurl").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n���Ԫ�ر�����д!");
				return false;			
			}
			if(document.getElementById("txt_adsurl").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\nվ�����ӱ�����д!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save&ItemID=<%=ItemID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�޸Ĺ��</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">��ʾ���ͣ�</TD>
				<TD>
				<select name="txt_adsdisplayID">
				<%
					set Rs1 = Con.Query("select id,categoryname from Sp_dictionary where categoryID = "&config_AdsCategoryID&"")
					if rs1.recordcount<>0 then
					do while not rs1.eof
					response.Write "<option value="&rs1("id")&""
					if rs1("id") = rs("adsdisplayID") then response.Write " selected"
					response.Write ">"&rs1("categoryname")&"</option>"
					rs1.movenext
					loop
					end if
				%>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>������ƣ�</TD>
				<TD><input name="txt_adsname" type="text" class="input" value="<%=rs("adsname")%>" size="60"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>���Ԫ��<span class="huitext">(ͼƬ,flash,����)</span>��</TD>
				<TD>
				<input name="txt_adspicurl" type="text" class="input" value="<%=rs("adspicurl")%>" size="60"><span class="huitext">&nbsp;����</span>
				<iframe id="uploadfile" name="uploadfile" width="100%" height="30px;" frameborder="0" scrolling="no" src="File_Upload.asp?obj=txt_adspicurl"></iframe>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>վ�����ӣ�</TD>
				<TD><input name="txt_adsurl" type="text" class="input" value="<%=rs("adsurl")%>" size="60"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ʾ���֣�</TD>
				<TD><input name="txt_adstitle" type="text" class="input" value="<%=rs("adstitle")%>" size="80"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�����</TD>
				<TD>
				<input name="txt_adswidth" type="text" class="input" value="<%=rs("adswidth")%>" size="5">
				&nbsp;��&nbsp;
				<input name="txt_adsheight" type="text" class="input" value="<%=rs("adsheight")%>" size="5"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��濪ʼʱ�䣺</TD>
				<TD><input name="txt_starttime" type="text" class="input" value="<%=rs("starttime")%>" onfocus="javascript:HS_setDate(this);"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>������ʱ�䣺</TD>
				<TD><input name="txt_endtime" type="text" class="input" value="<%=rs("endtime")%>" onfocus="javascript:HS_setDate(this);"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ע��</TD>
				<TD><textarea name="txt_memo" cols="80" rows="6"><%=rs("memo")%></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ񼤻</TD>
				<TD><input name="txt_IsActive" class="input" type="checkbox" <%if rs("IsActive")=1 then response.Write "checked"%> value="1"></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�޸�����" class="button" type="submit">
		</div>
		<br>
		<div>
			<TABLE cellSpacing="1" width="100%">
			<TR class="content-td1">
				<TH align="left" colspan="2">�������</TH>
			</TR>
			<TR class="content-td1">
				<TD>asp</TD>
				<TD>
				<textarea name="textarea" cols="120" rows="2"><script src="http://<%=request.ServerVariables("HTTP_HOST")%>/plugIn/ads.show.asp?ItemID=<%=ItemID%>"></script></textarea>
				</TD>
			</TR>
			</TABLE>
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
	end if
%>
