<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ExpModelID:ExpModelID = VerificationUrlParam("ExpModelID","int","-1")
	Dim id:id = VerificationUrlParam("id","int","0")
	if action<>"" and action="del" then
		''ȡģ�Ͷ�Ӧ����ϸ��Ϣ
		Dim ModelArray:ModelArray = Con.QueryRow("select * from Sp_Model where ID="&ID&"",0)
		Dim ModelTable:ModelTable = ModelArray(2)
		'response.Write ModelTable
		'response.End()
		''ɾ����Ϣ
		Con.Execute("delete * from Sp_Model where ID="&id&"")
		Con.Execute("delete * from Sp_ModelField where ModelID="&id&"")
		''���¹�ϵģ�͵�ID
		Con.Execute("update Sp_Model set ParentModelID=0 where ParentModelID="&id&"")
		Con.Execute("update Sp_Model set SubModelID=0 where SubModelID="&id&"")
		''ɾ�������ṹ		
		if ExpModelID=0 then		
			Con.Execute("drop table "&ModelTable&"")
		end if
		''''ɾ��ģ���û����ϵ
		Con.Execute("delete * from Sp_ModelGroup where ModelID="&id&"")
		'''ɾ��ģ����Ϣ�����¼
		Con.Execute("delete * from Sp_ModelPurchaseLog where modelID="&id&"")
		''''ɾ��ģ�Ͷ�Ӧ��Ȩ�޹�����еķ����¼
		Con.Execute("delete * from Sp_ManagePower where ModelID="&id&"")
		'''''
		ALert "ɾ��ģ�ͳɹ�","Setting.Model.List.asp"
	elseif action="clear" then
		Dim table:table = VerificationUrlParam("table","string","")
		if submodelID=0 then		
		Con.Execute("delete * from "&table&"")
		'''ɾ��ģ����Ϣ�����¼
		Con.Execute("delete * from Sp_ModelPurchaseLog where modelID="&id&"")
		end if
		ALert "��ձ�ɹ�","Setting.Model.List.asp"
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">ģ���б�</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="AllCheck"></TH>
			    <TH width="25%" align="left">ģ������</TH>
			    <TH width="25%" align="left">ģ�ͱ�����</TH>
			    <TH width="15%" align="left">ģ������Ӧ�����ֵ�</TH>
			    <TH width="7%" align="left">�Ƿ���ʾ</TH>
			    <TH width="23%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Model M,Sp_dictionaryCategory D where D.ID=M.modelCategoryID")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("M.*,D.dictionary","Sp_Model M,Sp_dictionaryCategory D", "D.ID=M.modelCategoryID", "M.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><%=rs("ModelName")%></TD>
				<TD><%=rs("ModelTable")%></TD>
				<TD><%=rs("dictionary")%></TD>
				<TD><%if rs("IsDisplay")=0 then:response.Write "����":else:response.Write "��ʾ":end if%></TD>
				<TD><a href="setting.Model.Field.list.asp?ModelID=<%=rs("ID")%>">�ֶ�</a>&nbsp;<a href="setting.Model.edit.asp?ModelID=<%=rs("ID")%>">�޸�</a>&nbsp;<%if rs("issystem")=0 then%><a href="javascript:if(confirm('Sp_CMSϵͳ��ʾ:\r\n\nȷ��ɾ��?ɾ���󲻿ɻָ�!'))location.href='?action=del&id=<%=rs("ID")%>&ExpModelID=<%=rs("ExpModelID")%>';">ɾ��</a><%end if%>&nbsp;<a href="javascript:if(confirm('Sp_CMSϵͳ��ʾ:\r\n\nȷ��ɾ��?ɾ���󲻿ɻָ�!'))location.href='?action=clear&table=<%=rs("Modeltable")%>&ExpModelID=<%=rs("ExpModelID")%>&id=<%=rs("ID")%>';">��ձ�</a>&nbsp;<%if rs("ExpModelID")=0 then response.Write "<a href='ExpandModel.asp?ModelID="&rs("id")&"'><span style='color:#999;'>�Ӵ���չ</span></a>"%></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,false,false%>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
