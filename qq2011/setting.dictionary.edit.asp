<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim location:location = VerificationUrlParam("location","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","")
	Dim CategoryID:CategoryID = VerificationUrlParam("CategoryID","int","")
	if action<>"" and action="save" then 
  		dim parentID:parentID = F.ReplaceFormText("txt_parentID")
  		dim categoryname:categoryname = F.ReplaceFormText("txt_categoryname")
		dim categoryname_en:categoryname_en = F.ReplaceFormText("txt_categoryname_en")
		''''
		dim Exp_img:Exp_img = F.ReplaceFormText("txt_Exp_img")
		dim Exp_field1:Exp_field1 = F.ReplaceFormText("txt_Exp_field1")
		dim Exp_field2:Exp_field2 = F.ReplaceFormText("txt_Exp_field2")
		dim Exp_field3:Exp_field3 = F.ReplaceFormText("txt_Exp_field3")
		''''
		Dim IsCommand:IsCommand = F.ReplaceFormText("IsCommand")
		Dim IsTop:IsTop= F.ReplaceFormText("IsTop")
		Dim IsHot:IsHot = F.ReplaceFormText("IsHots")
		if IsCommand="" then IsCommand = 0
		if IsTop="" then IsTop = 0
		if IsHot="" then IsHot = 0
		'''''
		dim IndexID:IndexID = F.ReplaceFormText("txt_IndexID")
		if IndexID = "" then IndexID = 0
		dim Pathstr
		Dim Depth
		'''''ȡ�������ID��Pathstr
		if parentID<>0 then
			Dim ItemArray:ItemArray = Con.QueryRow("select top 1 Pathstr,Depth from Sp_dictionary where ID="&parentID&"",0)
			Pathstr = ItemArray(0)
			Depth = ItemArray(1)
		else
			Pathstr = 0
			Depth = 1
		end if
		Depth = Depth + 1
		Con.execute("update Sp_dictionary set IndexID="&IndexID&",categoryID="&categoryID&",categoryname='"&categoryname&"',categoryname_en='"&categoryname_en&"',Pathstr='0',Exp_img = '"&Exp_img&"',Exp_field1='"&Exp_field1&"',Exp_field2='"&Exp_field2&"',Exp_field3='"&Exp_field3&"',parentID="&parentID&",Depth="&Depth&",IsCommand="&IsCommand&",IsTop="&IsTop&",IsHot="&IsHot&" where ID="&ItemID&"")

		if parentID<>0 then
			Pathstr = Pathstr &","&ItemID
		else
			Pathstr = ItemID
		end if
		'response.Write Pathstr
		'response.End()
		''''
		Con.execute("update Sp_dictionary set Pathstr='"&Pathstr&"' where ID="&ItemID&"")
		Alert "�����ֵ�ɹ�","setting.dictionary.asp?CategoryID="&categoryID&"&location="&location&""
	end if
	'''''
	Dim dictionaryArray:dictionaryArray = Con.QueryData("select id,categoryname,Pathstr,parentID from Sp_dictionary where categoryID="&CategoryID&"")
	''''''
	Dim dictionaryItemArray:dictionaryItemArray = Con.QueryRow("select ID,categoryname,Pathstr,parentID,categoryname_en,IndexID,Exp_img,Exp_field1,Exp_field2,Exp_field3,iscommand,ishot,istop from Sp_dictionary where ID="&ItemID&"",0)
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
<form action="?action=save&CategoryID=<%=CategoryID%>&ItemID=<%=ItemID%>&location=<%=location%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�����ֵ�</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <%
			  ''''''�Ƿ��������ݹ�������������
			  if location="Model" then
			  	response.Write "<input name=""CategoryID"" type=""hidden"" value="""&CategoryID&""">"
			  else
			  %>			  
			  <TR class="content-td1">
				<TD width="15%">������ƣ�</TD>
				<TD>
				<select name="txt_categoryID" onChange="javascript:location.href='?CategoryID='+this.value+'&ItemID=<%=ItemID%>';">
				<%
					set Rs = Con.Query("select id,dictionary from Sp_dictionaryCategory")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&""
					if Cint(CategoryID)=Cint(rs("id")) then
					response.Write " selected"
					end if
					response.Write ">"&rs("dictionary")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <%
			  end if
			  %>
			  <TR class="content-td1">
				<TD width="15%">&nbsp;</TD>
				<TD>
				<select name="txt_parentID">
				<option value="0">------</option>
				<%filldictionaryDropDownList dictionaryArray,0,dictionaryItemArray(3)%>
				</select>
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֵ����ƣ�</TD>
				<TD><input name="txt_categoryname" class="input" type="text" value="<%=dictionaryItemArray(1)%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֵ�����(Ӣ��)��</TD>
				<TD><input name="txt_categoryname_en" class="input" type="text" value="<%=dictionaryItemArray(4)%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1" id="PicUpload">
				<TD>��չͼƬ��ַ��</TD>
				<TD>
				<input name="txt_Exp_img" class="input" size="60" type="text" value="<%=dictionaryItemArray(6)%>"><br>
				<iframe name="ad" frameborder="0"   width="100%" height="30px;" scrolling="no" src="File_Upload.asp?obj=txt_Exp_img"></iframe>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��չ�ֶ�һ��</TD>
				<TD><input name="txt_Exp_field1" class="input" type="text" value="<%=dictionaryItemArray(7)%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��չ�ֶζ���</TD>
				<TD><input name="txt_Exp_field2" class="input" type="text" value="<%=dictionaryItemArray(8)%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��չ�ֶ�����</TD>
				<TD><input name="txt_Exp_field3" class="input" type="text" value="<%=dictionaryItemArray(9)%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>����:</TD>
				<TD>
				<input name="IsCommand" type="checkbox" value="1" <%if dictionaryItemArray(10)=1 then response.Write "checked"%>>�Ƽ�&nbsp;&nbsp;
				<input name="IsTop" type="checkbox" value="1" <%if dictionaryItemArray(12)=1 then response.Write "checked"%>>�ö�&nbsp;&nbsp;
				<input name="IsHots" type="checkbox" value="1" <%if dictionaryItemArray(11)=1 then response.Write "checked"%>>�ȵ�&nbsp;&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ʾ˳��</TD>
				<TD><input name="txt_IndexID" class="input" value="<%=dictionaryItemArray(5)%>" type="text"><span class="huitext">&nbsp;��ֵԽСԽ��ǰ,��ͬʱ����������������.</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�����ֵ�" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
