<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">ģ���ֶ�����</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<div class="divpadding">
		��Ϣ������
		<select name="selcategory" style="width:100px;" class="select" onchange="javascript:if(this.value!='') location.href='?modelID='+this.value;">
		<option value="">ѡ��ģ��</option>
		<%
			Dim modelID:modelID = VerificationUrlParam("modelID","int","0")
			set rs1 = Con.Query("select id,modelname from Sp_Model")
			if rs1.recordcount<>0 then
				do while not rs1.eof
		%>
		<option value="<%=rs1("id")%>" <%if Cint(modelID)=Cint(rs1("id")) then response.Write "selected"%>><%=rs1("modelname")%></option>
		<%
				rs1.movenext
				loop
			end if
		%>
		</select>
		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH align=left><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH align=left>ģ������</TH>
			    <TH align=left>�ֶ�����</TH>
			    <TH align=left>�ֶ�����</TH>
			    <TH align=left>�ֶ�ҳ������</TH>
			    <TH align=left>�Ƿ����</TH>
			    <TH align=left></TH>
		      </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Model M, Sp_ModelField F where M.ID=F.ModelID and M.ID="&modelID&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("F.ID,M.ID as MID","Sp_Model M,Sp_ModelField F", "M.ID=F.ModelID and M.ID="&modelID&"", "F.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD width="5%">&nbsp;</TD>
				<TD width="20%">&nbsp;</TD>
				<TD width="20%">&nbsp;</TD>
				<TD width="20%">&nbsp;</TD>
				<TD width="20%">&nbsp;</TD>
				<TD width="5%">&nbsp;</TD>
				<TD width="10%" align="center"><a href="field_add.asp?id=<%=rs("id")%>">�޸�</a>&nbsp;<a href="javascript:if(confirm('ȷ��ɾ��?ɾ������Ϣ���ɻָ�!')) location.href='field_del.asp?action=del&id=<%=rs("id")%>';">ɾ��</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<input name="btnsearch" value="����ɾ��" class="button" type="button">
		</div>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
