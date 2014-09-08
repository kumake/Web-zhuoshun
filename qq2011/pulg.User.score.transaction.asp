<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim UseruID:UseruID = VerificationUrlParam("UseruID","int","0")
	Dim moneyType:moneyType = VerificationUrlParam("Type","string","")
	Dim action:action = VerificationUrlParam("action","string","")
	''''
	if action<>"" and action="delete" then
		Dim BatchremoveItemID:BatchremoveItemID = F.ReplaceFormText("ItemID")
		Con.Execute("delete * from Sp_TransactionLog where ID in ("&BatchremoveItemID&")")
		'''''''
		WriteLog CK("username"),"action","批量删除用户ID"&UseruID&"的交易记录"
		'''''''
		Alert "批量删除成功","?type="&moneyType&"&useruid="&useruid&""
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
	<script language="javascript" type="text/javascript">
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
				if (e.name == "ItemID" && e.disabled==false && e.type=="checkbox") 
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
		function batchAction()
		{
			if(checkedItem()==0)
			{
				alert("Sp_CMS系统提示:\r\n\n请选择要批量操作的信息!");
				return false;
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=delete&type=<%=moneyType%>&useruid=<%=useruid%>" method="post" onSubmit="javascript:return batchAction();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">用户账户交易记录</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="allCheck" onclick="javascript:SelectALLCheck(this);"></TH>
			    <TH width="10%" align="left">用户名称</TH>
			    <TH width="10%" align="left">账户类型</TH>
			    <TH width="10%" align="left">交易类型</TH>
			    <TH width="10%" align="left">交易金额</TH>
			    <TH width="35%" align="left">备注</TH>
			    <TH width="10%" align="left">交易日期</TH>
			  </TR>
			  <%
				Dim total_record   	'总记录数
				Dim current_page	'当前页面
				Dim PCount			'循环页显示数目
				Dim pagesize		'每页显示记录数
				Dim showpageNum		'是否显示数字循环页
				Dim strwhere:strwhere = "U.id=G.useruid"
				if useruid<>0 then strwhere =strwhere &" and u.ID="&UseruID&""
				if moneyType<>"" then  strwhere =strwhere &" and G.moneyType='"&moneyType&"'"
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_User U, Sp_TransactionLog G where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
				Dim sql:sql = Con.QueryPageByNum("U.username,U.GroupID,G.*","Sp_User U, Sp_TransactionLog G", ""&strwhere&"", "U.ID", true,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemID" value="<%=rs("ID")%>"></TD>
				<TD><a href="pulg.user.view.asp?GroupID=<%=rs("GroupID")%>&ItemID=<%=rs("useruid")%>"><%=rs("username")%></a></TD>
				<TD><%if rs("moneyType")="M" then:response.Write "账户金额":else:response.Write "网站积分":end if%></TD>
				<TD><%if rs("plusminus")=0 then:response.Write "消费":else:response.Write "充值":end if%></TD>
				<TD><%=rs("score")%></TD>
				<TD><%=rs("tranmemo")%></TD>
				<TD><%=year(rs("trantime"))%>-<%=month(rs("trantime"))%>-<%=day(rs("trantime"))%></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<input name="btnrestore" value="批量删除" class="button" type="submit">		
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
