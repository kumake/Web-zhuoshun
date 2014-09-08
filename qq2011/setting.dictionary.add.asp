<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim location:location = VerificationUrlParam("location","string","")
	Dim CategoryID:CategoryID = VerificationUrlParam("CategoryID","int","1")
	if action<>"" and action="save" then 
  		dim parentID:parentID = F.ReplaceFormText("txt_parentID")
  		dim categoryname:categoryname = F.ReplaceFormText("txt_categoryname")
		dim categoryname_en:categoryname_en = F.ReplaceFormText("txt_categoryname_en")
		''''
		dim Exp_img:Exp_img = F.ReplaceFormText("txt_Exp_img")
		dim Exp_field1:Exp_field1 = F.ReplaceFormText("txt_Exp_field1")
		dim Exp_field2:Exp_field2 = F.ReplaceFormText("txt_Exp_field2")
		dim Exp_field3:Exp_field3 = F.ReplaceFormText("txt_Exp_field3")
		'''''
		Dim IsCommand:IsCommand = F.ReplaceFormText("IsCommand")
		Dim IsTop:IsTop= F.ReplaceFormText("IsTop")
		Dim IsHot:IsHot = F.ReplaceFormText("IsHots")
		if IsCommand="" then IsCommand = 0
		if IsTop="" then IsTop = 0
		if IsHot="" then IsHot = 0
		'''''
		dim IndexID:IndexID = F.ReplaceFormText("txt_IndexID")
		if IndexID="" then IndexID = 0
		dim Pathstr
		Dim Depth
		'''''取付父类别ID的Pathstr
		if parentID<>0 then
			Dim ItemArray:ItemArray = Con.QueryRow("select top 1 Pathstr,Depth from Sp_dictionary where ID="&parentID&"",0)
			Pathstr = ItemArray(0)
			Depth = ItemArray(1)
		else
			Pathstr = 0
			Depth = 1
		end if
		Depth = Depth + 1
		Con.execute("insert into Sp_dictionary (IndexID,categoryID,categoryname,categoryname_en,Pathstr,Exp_img,Exp_field1,Exp_field2,Exp_field3,parentID,Depth,IsCommand,ishot,istop) values ("&IndexID&","&categoryID&",'"&categoryname&"','"&categoryname_en&"','0','"&Exp_img&"','"&Exp_field1&"','"&Exp_field2&"','"&Exp_field3&"',"&parentID&","&Depth&","&IsCommand&","&ishot&","&istop&")")
		'''''取最新的ID
		ItemArray = Con.QueryRow("select top 1 id from Sp_dictionary order by id desc",0)
		if parentID<>0 then
			Pathstr = Pathstr &","&ItemArray(0)
		else
			Pathstr = ItemArray(0)
		end if
		'response.Write Pathstr
		'response.End()
		''''
		Con.execute("update Sp_dictionary set Pathstr='"&Pathstr&"' where ID="&ItemArray(0)&"")
		Alert "增加字典成功","setting.dictionary.asp?CategoryID="&categoryID&"&location="&location&""
	end if
	'''''
	Dim dictionaryArray:dictionaryArray = Con.QueryData("select id,categoryname,Pathstr,parentID from Sp_dictionary where categoryID="&CategoryID&"")
	''''''
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_categoryID").value=="")
			{
				alert("Sp_CMS提示\r\n\n类别名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_categoryname").value=="")
			{
				alert("Sp_CMS提示\r\n\n字典名称必须填写!");
				return false;			
			}			
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save&CategoryID=<%=CategoryID%>&location=<%=location%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onFocus="this.blur()" onClick="selectTag('tagContent0',this)" href="javascript:void(0)">增加字典</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <%
			  ''''''是否是在内容管理处进行添加类别
			  if location="Model" then
			  	response.Write "<input name=""CategoryID"" type=""hidden"" value="""&CategoryID&""">"
			  else
			  %>			  
			  <TR class="content-td1">
				<TD width="15%">类别名称：</TD>
				<TD>
				<select name="txt_categoryID" onChange="javascript:location.href='?CategoryID='+this.value;">
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
				<span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <%
			  end if
			  %>
			  <TR class="content-td1">
				<TD width="15%">&nbsp;</TD>
				<TD>
				<select name="txt_parentID">
				<option value="0">------</option>
                
                <%set Rs = Con.Query("select id,dictionary from Sp_dictionaryCategory")
				if rs.recordcount<>0 then%>
				<%filldictionaryDropDownList dictionaryArray,0,""%>
                <%end if%>
				</select>
				<span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字典名称：</TD>
				<TD><input name="txt_categoryname" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字典名称(英文)：</TD>
				<TD><input name="txt_categoryname_en" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1" id="PicUpload">
				<TD>扩展图片地址：</TD>
				<TD>
				<input name="txt_Exp_img" class="input" size="60" type="text"><br>
				<iframe name="ad" frameborder="0"   width="100%" height="30px;" scrolling="no" src="File_Upload.asp?obj=txt_Exp_img"></iframe>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>扩展字段一：</TD>
				<TD><input name="txt_Exp_field1" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>扩展字段二：</TD>
				<TD><input name="txt_Exp_field2" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>扩展字段三：</TD>
				<TD><input name="txt_Exp_field3" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>属性：</TD>
				<TD>
				<input name="IsCommand" type="checkbox" value="1">推荐&nbsp;&nbsp;
				<input name="IsTop" type="checkbox" value="1">置顶&nbsp;&nbsp;
				<input name="IsHots" type="checkbox" value="1">热点&nbsp;&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>显示顺序</TD>
				<TD><input name="txt_IndexID" class="input" value="0" type="text"><span class="huitext">&nbsp;数值越小越考前,相同时候其他参数其作用.</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="增加字典" class="button" <%if CategoryID = 1 then response.Write "disabled"%> type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
