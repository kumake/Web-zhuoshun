<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<form action="?action=batchdel" method="post" onSubmit="javascript:return batchAction();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">��Ʒ����ͳ�ƹ���</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
			    <TH width="57%" align="left">��Ʒ����</TH>
			    <TH width="19%" align="left">��Ʒ�۸�</TH>
			    <TH width="20%" align="left">��Ʒ��������</TH>
		      </TR>
			  <%
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = "select top 100 P.ID,P.title,I.price,sum(I.num) as countnum from Sp_OrderItem I,user_pro P where I.proid=P.ID group by P.ID,P.title,I.price order by sum(I.num) desc"
				'response.Write sql
				'response.End()
				set rs = Con.Query(sql)
				if rs.recordcount<>0 then
				dim rsindex:rsindex = 1
				do while not rs.eof and rsindex <=100
			  %>
			  <TR class="content-td1">
				<TD><%=rs("title")%></TD>
				<TD><%=rs("price")%></TD>
				<TD><%=rs("countnum")%></TD>
			  </TR>
			  <%
			   rs.movenext
			   rsindex = rsindex + 1
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
