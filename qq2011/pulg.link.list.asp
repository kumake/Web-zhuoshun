<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","")
	if action<>"" and action="del" then
		''ɾ����Ϣ
		Con.Execute("delete * from Sp_FriendLink where ID="&ItemID&"")
		ALert "ɾ���ɹ�","pulg.link.list.asp"
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�������ӹ���</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="10%" align="left">�������</TH>
			    <TH width="15%" align="left">��������</TH>
			    <TH width="15%" align="left">վ������</TH>
			    <TH width="20%" align="left">վ���ַ</TH>
			    <TH width="20%" align="left">logo</TH>
			    <TH width="5%" align="left">��ʾ</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				Dim strwhere:strwhere = "F.linkCategory=D.id"
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_FriendLink F,Sp_dictionary D where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("F.*,D.Categoryname","Sp_FriendLink F,Sp_dictionary D", ""&strwhere&"", "F.ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><%=rs("Categoryname")%></TD>
				<TD><%if rs("urltype")=1 then:response.Write "logo":else:response.Write "�ı�":end if%></TD>
				<TD><%=rs("sitename")%></TD>
				<TD><%=rs("siteurl")%></TD>
				<TD><%if rs("urltype")=1 and rs("sitelogourl")<>"" then response.Write "<img src='"&rs("sitelogourl")&"' align='absmiddle'>"%></TD>
				<TD><%if rs("ISdisplay")=1 then:response.Write "��":else:response.Write "��":end if%></TD>
				<TD><a href="pulg.link.edit.asp?ItemID=<%=rs("ID")%>">�޸�</a>&nbsp;<a href="javascript:if(confirm('Sp_CMSϵͳ��ʾ:\r\n\nȷ��ɾ��?ɾ���󲻿ɻָ�!'))location.href='?action=del&ItemID=<%=rs("ID")%>';">ɾ��</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
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
				<textarea name="textarea" cols="120" rows="2"><script src="http://<%=request.ServerVariables("HTTP_HOST")%>/plugIn/link.show.asp"></script></textarea>
				</TD>
			</TR>
			</TABLE>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
