<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","")
	if action<>"" and action="del" then
		''ɾ����Ϣ
		set rs = con.query("select * from Sp_City where parentID="&ItemID&"")
		if rs.recordcount<>0 then
			ALert "����ɾ���¼�Ŀ¼","setting.city.list.asp"
		else
			Con.Execute("delete * from Sp_City where ID="&ItemID&"")
			ALert "ɾ���ֵ�ɹ�","setting.city.list.asp"
		end if
	end if	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">��������</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--��1����ǩ//-->
		<div class="divpadding">		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="85%" align="left">����</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
				'''ȡ���ݴ�Ž�ModelCategoryArray
				Dim CityCategoryArray:CityCategoryArray = Con.QueryData("select ID,Areaname,parentID,Depth,pathstr from Sp_city")
				function ResponseCity(CityCategoryArray,parentID)
					for arraystep=0 to ubound(CityCategoryArray,2) step 1
						'if spaceLen<>0 then spaceLen = spaceLen - 1
						if CityCategoryArray(2,arraystep)=parentID then
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=CityCategoryArray(0,arraystep)%>"></TD>
				<TD><%=string((CityCategoryArray(3,arraystep)-1),"��")%><%=CityCategoryArray(1,arraystep)%></TD>
				<TD><a href="setting.city.edit.asp?id=<%=CityCategoryArray(0,arraystep)%>">�޸�</a>&nbsp;<a href="javascript:if(confirm('Sp_CMSϵͳ��ʾ:\r\n\nȷ��ɾ��?ɾ���󲻿ɻָ�!'))location.href='setting.city.list.asp?action=del&ItemID=<%=CityCategoryArray(0,arraystep)%>';">ɾ��</a></TD>
			  </TR>
			  <%
							ResponseCity CityCategoryArray,CityCategoryArray(0,arraystep)
						end if
					next
				end function				
			  %>
			  <%ResponseCity CityCategoryArray,0%>
			</TABLE>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
