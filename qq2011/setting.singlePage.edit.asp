<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="save" then 
  		dim ParentID:ParentID = F.ReplaceFormText("txt_ParentID")
		dim isneedpic:isneedpic = F.ReplaceFormText("txt_isneedpic")
		dim IndexID:IndexID = F.ReplaceFormText("txt_IndexID")
		dim isneedinstr:isneedinstr = F.ReplaceFormText("txt_isneedinstr")		
		dim singlePagename:singlePagename = F.ReplaceFormText("txt_singlePagename")
		dim singlePagename_En:singlePagename_En = F.ReplaceFormText("txt_singlePagename_En")
		dim contentTemplate:contentTemplate = F.ReplaceFormText("txt_contentTempalte")
		if isneedinstr = "" then isneedinstr = 0
		if isneedpic = "" then isneedpic = 0
		if IndexID="" then IndexID=0
		Dim sql:sql = "update Sp_SinglePage set IndexID="&IndexID&",isneedinstr="&isneedinstr&",singlePagename='"&singlePagename&"',singlePagename_En='"&singlePagename_En&"',isneedpic="&isneedpic&",parentID="&ParentID&",contentTemplate='"&contentTemplate&"' where ID="&ItemID&""
		Con.Execute(sql)
		Alert "�޸ĵ�ҳ�ɹ�","setting.singlePage.list.asp"
	end if
	if ItemID=0 then
		Alert "��������ʧ��",""
	else
		set rs = con.Query("select * from Sp_SinglePage where ID="&ItemID&"")
		if rs.recordcount=0 then
			Alert "������Ϣʧ��",""
		else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
    <script language="javascript" type="text/javascript" src="Scripts/popup.js"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_singlePagename").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n��ҳ���Ʊ�����д!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�޸ĵ�ҳ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">��ҳ���ࣺ</TD>
				<TD>
				<select name="txt_ParentID" class="input">
				<option value="0">ѡ�����</option>
				<%
					set Rs1 = Con.Query("select id,singlePagename from Sp_SinglePage where parentID=0 and singlePagename<>'����ϵͳ'")
					if rs1.recordcount<>0 then
					do while not rs1.eof
					response.Write "<option value="&rs1("id")&""
					if rs("parentID")=rs1("id") then response.Write " selected"
					response.Write ">"&rs1("singlePagename")&"</option>"
					rs1.movenext
					loop
					end if
				%>
				</select>
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ҳ���ƣ�</TD>
				<TD><input name="txt_singlePagename" class="input" type="text" value="<%=rs("singlePagename")%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ҳ����(Ӣ��)��</TD>
				<TD><input name="txt_singlePagename_En" class="input" type="text" value="<%=rs("singlePagename_En")%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ʾ˳��</TD>
				<TD><input name="txt_IndexID" class="input" type="text" value="<%=rs("IndexID")%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>�Ƿ���Ҫ˵����ͼƬ</TD>
		        <TD><input type="checkbox" name="txt_isneedpic" value="1" <%if rs("isneedpic")=1 then response.Write "checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>����ҳģ�壺</TD>
				<TD><input name="txt_contentTempalte" value="<%=rs("contentTemplate")%>" type="text" class="input" size="70">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_contentTempalte');">ѡ��</a></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>�Ƿ���Ҫ���</TD>
		        <TD><input type="checkbox" name="txt_isneedinstr" value="1"  <%if rs("isneedinstr")=1 then response.Write "checked"%>></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�޸ĵ�ҳ" class="button" type="submit">
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
<script language=javascript>
    //ѡ��ģ��
	var g_pop;
    function SelectTemplate(obj)
    {
        g_pop=new Popup({ contentType:1, isReloadOnClose:false,scrollType:'no', width:700, height:320 });
        g_pop.setContent("title","ѡ��ģ��");
        g_pop.setContent("contentUrl","Template.asp?obj="+obj);
    	
        g_pop.build();
        g_pop.show();
        return false;
    }
    
    //�رմ򿪵Ĵ���
    function ClosePop()
    {
        g_pop.close();
    }
</script>
