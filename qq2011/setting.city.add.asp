<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("id","int","0")
	if action<>"" and action="save" then 
  		dim parentID:parentID = F.ReplaceFormText("txt_parentID")
  		dim categoryname:categoryname = F.ReplaceFormText("txt_cityname")
		dim IndexID:IndexID = F.ReplaceFormText("txt_IndexID")
		if IndexID="" then IndexID = 0
		
		dim Pathstr
		dim Depth:Depth = 1
		'''''ȡ�������ID��Pathstr
		if parentID<>0 then
			Dim ItemArray:ItemArray = Con.QueryRow("select top 1 Pathstr,Depth from Sp_City where ID="&parentID&"",0)
			Pathstr = ItemArray(0)
			Depth = (ItemArray(1)+1)
		else
			Pathstr = 0
		end if
		Con.execute("insert into Sp_City (IndexID,Areaname,Pathstr,Depth,parentID) values ("&IndexID&",'"&categoryname&"','0',"&Depth&","&parentID&")")
		'''''ȡ���µ�ID
		ItemArray = Con.QueryRow("select top 1 id from Sp_City order by id desc",0)
		if parentID<>0 then
			Pathstr = Pathstr &","&ItemArray(0)
		else
			Pathstr = ItemArray(0)
		end if
		'response.Write Pathstr
		'response.End()
		''''
		Con.execute("update Sp_City set Pathstr='"&Pathstr&"' where ID="&ItemArray(0)&"")
		Alert "���ӵ����ɹ�","setting.city.list.asp?CategoryID="&categoryID&""
	end if
	'''''
	Dim dictionaryArray:dictionaryArray = Con.QueryData("select ID,Areaname,pathstr,parentID,Depth from Sp_city")
	''''''
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_categoryID").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n������Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_categoryname").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�ֵ����Ʊ�����д!");
				return false;			
			}			
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�޸ĳ���</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">&nbsp;</TD>
				<TD>
				<select name="txt_parentID">
				<option value="0">------</option>
				<%filldictionaryDropDownList dictionaryArray,0,""%>
				</select>
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>���ƣ�</TD>
				<TD><input name="txt_cityname" class="input" type="text" value=""><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ʾ˳��</TD>
				<TD><input name="txt_IndexID" class="input" value="0" type="text"><span class="huitext">&nbsp;��ֵԽСԽ��ǰ,��ͬʱ����������������.</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="���ӵ���" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
