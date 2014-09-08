<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	'action
	Dim action:action = VerificationUrlParam("action","string","")
	'新闻关键字筛选
	Dim keyword:keyword = VerificationUrlParam("keyword","string","")
	'存放该模型数据的一维数组
	Dim id:id = VerificationUrlParam("id","int","0")
	if action<>"" and action="remove" then
		Con.Execute("delete * from Sp_EmailFeed where ID="&id&"")
		ALert "删除成功","pulg.email.list.asp"
	end if	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript">
		function ModelSearch()
		{
			var obj2 = document.getElementById("txt_keyword").value;
			var url = "?FormID=<%=FormID%>";
			if(obj2!="")
			{
				url = url + "&keyword="+obj2;
			}
			location.href=url;
		}
		//////批量操作信息
		function SelectALLCheck(obj)
		{
			for (var i=0;i<document.forms[0].elements.length;i++) 
			{ 
				var e = document.forms[0].elements[i]; 
				if (e.name != obj.name && e.disabled==false && e.type=="checkbox") 
				{
					e.checked = obj.checked;
				}				
			} 			
		}
		/////
		function checkedItem()
		{
			var Item = 0;
			for (var i=0;i<document.forms[0].elements.length;i++) 
			{ 
				var e = document.forms[0].elements[i]; 
				if (e.name == "ItemUID" && e.disabled==false && e.type=="checkbox") 
				{
					if(e.checked)
					{
						Item = Item + 1
					}
				}				
			} 
			return Item;
		}
		//////
		function batchReMove()
		{
			if(checkedItem()==0)
			{
				alert("Sp_CMS系统提示:\r\n\n请选择要批量操作的信息!");
			}
			else
			{
				document.forms[0].action = "?FormID=<%=ID%>&action=BatchreMove";
				document.forms[0].submit();
			}		
		}
	</script>
</HEAD>
<BODY>
<form action="" method="post">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">Email信息订阅</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<div class="divpadding">
		信息检索：		
		<input name="txt_keyword" type="text" class="input" value="<%=keyword%>">
		<input name="btnserch" type="button" onClick="javascript:ModelSearch();" value="检索信息" class="button">
		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align=left><input type="checkbox" name="allCheck" onclick="javascript:SelectALLCheck(this);"></TH>
			    <TH width="50%" align=left>邮件名称</TH>
			    <TH width="20%" align=left>是否订阅</TH>
			    <TH width="15%" align=left>最后发送邮件时间</TH>
			    <TH width="10%" align=left></TH>
		      </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strWhere		'查询条件
				
				strWhere = "1=1"
				''''查询条件
				if keyword<>"" then strWhere = strWhere &" and email like '%"&keyword&"%'"

				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_EmailFeed where "&strWhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("email,isfeed,lastsendmail,id","Sp_EmailFeed", ""&strWhere&"", "ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemUID" value="<%=rs("id")%>"></TD>
				<TD><a href="pulg.email.send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></TD>
				<TD><%if rs("isfeed")=1 then:response.Write "是":else:response.Write "否":end if%></TD>
				<TD><%=Year(rs("lastsendmail"))%>-<%=month(rs("lastsendmail"))%>-<%=day(rs("lastsendmail"))%></TD>
				<TD align="center"><a href="javascript:if(confirm('确定删除?删除后信息不可恢复!')) location.href='?action=remove&id=<%=rs("id")%>';">删除</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<input name="btnsearch" onClick="javascript:batchReMove();" value="批量删除" class="button" type="button">
		</div>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
