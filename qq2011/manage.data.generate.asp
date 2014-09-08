<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
Dim modelID:modelID = VerificationUrlParam("modelID","int","0")
Dim ModelCategoryArray
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<script language="javascript" type="text/javascript" src="Scripts/calender.js"></script>
    <script language="javascript" type="text/javascript" src="Scripts/popup.js"></script>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">数据更新中心</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
			<TABLE cellSpacing=1 width="100%">
			<TR class=content-td1>
				<TH colspan="2" align=left scope=col>页面生成管理 </TH>
			</TR>
			<TR class=content-td1>
				<TD width="36%" height="40"><input type="button" name="btnchangePwd" class="buttonlen" value="生成首页" onClick="if(confirm('确定需要生成页面!')) javascript:SelectTemplate('manage.index.generate.asp');"></TD>
			    <TD width="64%" rowspan="4">&nbsp;</TD>
			</TR>
			<TR class=content-td1>
			  <TD height="40"><input type="button" name="btnchangePwd" class="buttonlen" value="刷新所有模型的列表页" onClick="javascript:if(confirm('确定需要生成页面!')) generateHtml('List');"></TD>
			  </TR>
			<TR class=content-td1>
			  <TD height="40"><input type="button" name="btnchangePwd" class="buttonlen" value="刷新所有js调用页面" onClick="javascript:if(confirm('确定需要生成页面!')) generateHtml('Js');"></TD>
			  </TR>
			<TR class=content-td1>
			  <TD valign="top"><input type="button" name="btnchangePwd" class="buttonlen" value="刷新所有单页系统" onClick="javascript:if(confirm('确定需要生成页面!')) generateHtml('SinglePage');"></TD>
			  </TR>
			</TABLE>
	  </DIV>
	  <br>
	  <DIV>
			<TABLE cellSpacing=1 width="100%">
			<TR class=content-td1>
				<TH colspan="2" align=left scope=col>按条件生成信息内容页面 </TH>
			</TR>
			<TR class=content-td1>
				<TD width="10%" height="20">刷新模型表</TD>
				<TD width="90%">
				<select name="txtModelID" onChange="javascript:location.href='?modelID='+this.value;">
				<%
				set rs = Con.Query("select ID,modelname,modeltable from Sp_Model where IsDisplay=1")
				if rs.recordcount<>0 then
				if modelID=0 then modelID = rs("id")
					do while not rs.eof
						response.Write "<option value='"&rs("id")&"#"&rs("modeltable")&"'"
						if Cint(modelID)=rs("id") then response.Write " selected"
						response.Write ">"&rs("modelname")&"</option>"
					rs.movenext
					loop
					rs.close()
					set rs = nothing
				end if
				%>				
				</select>
				</TD>
			</TR>
			<TR class=content-td1>
				<TD height="20">刷新栏目</TD>
				<TD>
				<select name="txtcategoryID">
				<%
				'''取数据存放进ModelCategoryArray
				ModelCategoryArray = Con.QueryData("select D.ID,D.categoryID,D.categoryname,D.Pathstr,D.ParentID from Sp_dictionaryCategory C,Sp_dictionary D,SP_Model M where C.id=D.categoryID and D.categoryID=M.modelCategoryID and M.ID="&ModelID&" order by D.id desc")
				''填充下拉菜单结束
				fillDropDownList ModelCategoryArray,0,categoryID
				%>
				</select>
				</TD>
			</TR>
			<TR class=content-td1>
				<TD height="20"><input type="radio" id="txtredate" name="txtretype" value="0">按时间刷新</TD>
				<TD>从<input name="txtstarttime" type="text" onfocus="javascript:HS_setDate(this);" class="input"> &nbsp;到 &nbsp;<input name="txtendtime" onfocus="javascript:HS_setDate(this);" type="text" class="input"> &nbsp;的数据 </TD>
			</TR>
			<TR class=content-td1>
				<TD height="20"><input name="txtretype" id="txtreId" type="radio" value="1" checked>
			    按ID刷新</TD>
				<TD>从<input name="txtstartID" type="text" class="input" value="0">
				&nbsp; 到 
				<input name="textendid" type="text" class="input" value="0">&nbsp;的数据 </TD>
			</TR>
			</TABLE>
	  </DIV>
	  <div class="divpadding">
	    <input name="btnsearch" value="按条件刷新" class="buttonlen" type="button" onClick="javascript:if(confirm('确定需要生成页面!')) generateHtml('Model');">
	  </div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
<script language=javascript>
	function generateHtml(GenerateType)
	{
		//GenerateType 表示生成静态Html页面类型
		//GenerateType种类
		//Model			模型
		//SinglePage	单页图文系统
		//List			列表页
		//Js			生成Javascript
		var tempval = document.getElementById("txtModelID").value;
		var tempArray = tempval.split("#");
		var modelID=tempArray[0];
		var modelTable = tempArray[1];
		var classid = document.getElementById("txtcategoryID").value;
		var retype = 0;
		if(document.getElementById("txtreId").checked) retype=1
		var start = 0;
		var startday = document.getElementById("txtstarttime").value;
		var endday = document.getElementById("txtendtime").value;		
		var startid=document.getElementById("txtstartID").value;
		var endid=document.getElementById("textendid").value;
		var havehtml=0;
		//SelectTemplate("manage.generate.asp?start=1&GenerateType="+GenerateType+"&modelID="+modelID+"&modelTable="+modelTable+"&classid="+classid+"&retype="+retype+"&start="+start+"&startday="+startday+"&endday="+endday+"&startid="+startid+"&endid="+endid+"&havehtml="+havehtml+"");
		SelectTemplate("manage.generate.asp?start=1");
	}
    //选择模板
	var g_pop;
    function SelectTemplate(url)
    {
        g_pop=new Popup({ contentType:1, isReloadOnClose:false,scrollType:'no', width:500, height:200 });
        g_pop.setContent("title","静态页面生成系统");
        g_pop.setContent("contentUrl",url);
    	
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