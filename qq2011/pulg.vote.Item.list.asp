<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim voteID:voteID = VerificationUrlParam("voteID","int","0")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="del" then
		''ɾ����Ϣ
		Con.Execute("delete * from Sp_VoteItem where id="&ItemID&"")		
		ALert "ɾ���ɹ�","pulg.vote.Item.list.asp?voteID="&voteID&""
	elseif action="save" then
		Dim voteItem:voteItem =F.ReplaceFormText("txt_voteItem")
		Dim voteNum:voteNum =F.ReplaceFormText("txt_voteNum")
		Dim voteMemo:voteMemo =F.ReplaceFormText("txt_voteMemo")
		Con.Execute("insert into Sp_VoteItem (VoteID,voteItem,voteNum,voteMemo) values ("&VoteID&",'"&voteItem&"',"&voteNum&",'"&voteMemo&"')")		
		ALert "���ӳɹ�","pulg.vote.Item.list.asp?voteID="&voteID&""
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
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_voteItem").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�������Ʊ�����д!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�����������</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent1',this)" href="javascript:void(0)">��������</A></LI>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent2',this)" href="javascript:void(0)">ͶƱͼʾ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="40%" align="left">��������</TH>
			    <TH width="15%" align="left">���ע</TH>
			    <TH width="15%" align="left">����Ʊ��</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				Dim strwhere:strwhere = "voteID = "&voteID&""
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_VoteItem where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("*","Sp_VoteItem", ""&strwhere&"", "ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><%=rs("voteItem")%></TD>
				<TD><%=rs("votememo")%></TD>
				<TD><%=rs("voteNum")%></TD>
				<TD><a href="javascript:if(confirm('Sp_CMSϵͳ��ʾ:\r\n\nȷ��ɾ��?ɾ���󲻿ɻָ�!'))location.href='?action=del&VoteID=<%=voteID%>&ItemID=<%=rs("ID")%>';">ɾ��</a></TD>
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
	</DIV>
	<DIV class="tagContent selectTag content" id="tagContent1" style="display:none;">
		<!--��1����ǩ//-->
		<form action="?action=save&VoteID=<%=VoteID%>" method="post" onSubmit="javascript:return checkform();">
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">�������ƣ�</TD>
				<TD><input name="txt_voteItem" class="input" type="text" value=""><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>����Ĭ��Ʊ����</TD>
				<TD><input name="txt_voteNum" class="input" type="text" value="0"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ע��</TD>
				<TD><textarea name="txt_voteMemo" cols="40" rows="5" class="input"></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="��������" class="button" type="submit">
		</div>
		</form>
	</DIV>
	<DIV class="tagContent selectTag content" id="tagContent2" style="display:none;">
		<DIV>
		�¹��ܿ�����...
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
