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
		Alert "修改单页成功","setting.singlePage.list.asp"
	end if
	if ItemID=0 then
		Alert "参数传递失败",""
	else
		set rs = con.Query("select * from Sp_SinglePage where ID="&ItemID&"")
		if rs.recordcount=0 then
			Alert "查找信息失败",""
		else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
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
				alert("Sp_CMS提示\r\n\n单页名称必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">修改单页</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">单页归类：</TD>
				<TD>
				<select name="txt_ParentID" class="input">
				<option value="0">选择类别</option>
				<%
					set Rs1 = Con.Query("select id,singlePagename from Sp_SinglePage where parentID=0 and singlePagename<>'关于系统'")
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
				<span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>单页名称：</TD>
				<TD><input name="txt_singlePagename" class="input" type="text" value="<%=rs("singlePagename")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>单页名称(英文)：</TD>
				<TD><input name="txt_singlePagename_En" class="input" type="text" value="<%=rs("singlePagename_En")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>显示顺序：</TD>
				<TD><input name="txt_IndexID" class="input" type="text" value="<%=rs("IndexID")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>是否需要说明性图片</TD>
		        <TD><input type="checkbox" name="txt_isneedpic" value="1" <%if rs("isneedpic")=1 then response.Write "checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>内容页模板：</TD>
				<TD><input name="txt_contentTempalte" value="<%=rs("contentTemplate")%>" type="text" class="input" size="70">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('txt_contentTempalte');">选择</a></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>是否需要简介</TD>
		        <TD><input type="checkbox" name="txt_isneedinstr" value="1"  <%if rs("isneedinstr")=1 then response.Write "checked"%>></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="修改单页" class="button" type="submit">
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
    //选择模板
	var g_pop;
    function SelectTemplate(obj)
    {
        g_pop=new Popup({ contentType:1, isReloadOnClose:false,scrollType:'no', width:700, height:320 });
        g_pop.setContent("title","选择模板");
        g_pop.setContent("contentUrl","Template.asp?obj="+obj);
    	
        g_pop.build();
        g_pop.show();
        return false;
    }
    
    //关闭打开的窗口
    function ClosePop()
    {
        g_pop.close();
    }
</script>
