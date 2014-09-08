<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	Dim OrderItemID:OrderItemID = VerificationUrlParam("OrderItemID","int","0")
	if action<>"" and action="replay" then
		Dim pstatus:pstatus = F.ReplaceFormText("txt_status")
		Dim moneystatus:moneystatus = F.ReplaceFormText("txt_moneystatus")
		Dim hiduseruid:hiduseruid = F.ReplaceFormText("hid_useruid")
		Dim hidreturnMoney:hidreturnMoney = F.ReplaceFormText("hid_returnMoney")
		Dim hidgenerateScore:hidgenerateScore = F.ReplaceFormText("hid_generateScore")
		Dim hidisbalance:hidisbalance = F.ReplaceFormText("hid_isbalance")		
		'''''''当订单确认成功要处理积分
		if cint(pstatus)=2 and hidisbalance=0 then
			Con.Execute("update Sp_User set score=score+"&hidgenerateScore&",[money]=[money]+"&hidreturnMoney&" where ID="&hiduseruid&"")
			'''增加积分变化记录
			Con.Execute("insert into Sp_TransactionLog (useruid,moneyType,plusminus,score,tranmemo) values ("&hiduseruid&",'S',1,"&hidgenerateScore&",'订单返积分')")
			Con.Execute("insert into Sp_TransactionLog (useruid,moneyType,plusminus,score,tranmemo) values ("&hiduseruid&",'M',1,"&hidreturnMoney&",'订单返金额')")
			Con.Execute("update Sp_OrderForm set isbalance=1 where ID="&ItemID&"")
		end if
		'''''''		
		Con.Execute("update Sp_OrderForm set moneystatus="&moneystatus&",status="&pstatus&" where ID="&ItemID&"")
		Alert "订单处理成功！","plug.order.view.asp?ItemID="&ItemID&""
	elseif action="del" then
		if OrderItemID = 0 then 
			Alert "参数传递错误",""
		else
			Con.Execute("delete * from Sp_OrderItem where ID="&OrderItemID&"")
			Alert "订单处理成功！","plug.order.view.asp?ItemID="&ItemID&""
		end if
	end if
	''''''
	if ItemID=0 then
		Alert "出现错误",""
	else
		set rs = Con.Query("select F.*,D.deliverytype,P.payType as paymentType from Sp_OrderForm F,Sp_orderDeliver D,Sp_orderPay P where F.DeliverType=D.id and F.paytype=P.ID and F.ID="&ItemID&"")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/calender.js" type="text/javascript"></script>
	<!--有些字段必须填写的javascript的验证//-->
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("replay").value=="")
			{
				alert("Sp_CMS提示\r\n\n回复内容必须填写!");
				return false;			
			}
			return true;
		}
	</script>
	<style media=print>      
		 .Noprint {    
			display:none;    
		 }    
		  
		 .NextPage {    
			page-break-after:always;    
		 }    
	</style>  
</HEAD>
<BODY>
<form action="?action=replay&ItemID=<%=ItemID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">订单详细展示</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<div class="divpadding">
		  <input name="btnsearch" value="订单打印" class="button Noprint" type="button" onClick="javascript:window.print();">
		</div>
		<DIV>
		 <!---//-->
		 <input name="hid_useruid" type="hidden" value="<%=rs("useruid")%>">
		 <input name="hid_returnMoney" type="hidden" value="<%=rs("returnMoney")%>">
		 <input name="hid_generateScore" type="hidden" value="<%=rs("generateScore")%>">
		 <input name="hid_isbalance" type="hidden" value="<%=rs("isbalance")%>">
		 <!---//-->
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="15%" align=left>&nbsp;</TH>
				<TH align=left>&nbsp;</TH>
		      </TR>
			  <TR class="content-td1">
				<TD>用户名：</TD>
				<TD><%=rs("username")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>联系电话：</TD>
				<TD><%=rs("tel")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>电子邮件：</TD>
				<TD><%=rs("email")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>配送地址：</TD>
				<TD><%=rs("address")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>留言内容：</TD>
				<TD><%=rs("content")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>提交时间：</TD>
				<TD><%=rs("addtime")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>提交IP：</TD>
				<TD><%=rs("ip")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>支付方式：</TD>
				<TD><a href="pulg.order.pay.list.asp"><u><%=rs("paymentType")%></u></a></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>配送方式：</TD>
				<TD><a href="pulg.order.deliver.list.asp"><u><%=rs("deliverytype")%></u></a></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>返点金额：</TD>
				<TD>￥<%=rs("returnMoney")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>本订单折算积分：</TD>
				<TD><%=rs("generateScore")%></TD>
			  </TR>			  			  
			  <TR class="content-td1">
				<TD>货款状态：</TD>
				<TD>
				<input type="radio" name="txt_moneystatus" value="0" <%if rs("moneystatus")=0 then response.Write "checked"%>>未到款&nbsp;
				<input type="radio" name="txt_moneystatus" value="1" <%if rs("moneystatus")=1 then response.Write "checked"%>>款到&nbsp;
				</TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>订单状态：</TD>
				<TD>
				<input type="radio" name="txt_status" value="-1" <%if rs("status")=-1 then response.Write "checked"%>>无效订单&nbsp;
				<input type="radio" name="txt_status" value="-2" <%if rs("status")=-2 then response.Write "checked"%>>缺货订单&nbsp;
				<input type="radio" name="txt_status" value="0" <%if rs("status")=0 then response.Write "checked"%>>新单处理中&nbsp;
				<input type="radio" name="txt_status" value="1" <%if rs("status")=1 then response.Write "checked"%>>款到发货&nbsp;
				<input type="radio" name="txt_status" value="2" <%if rs("status")=2 then response.Write "checked"%>>确认收货&nbsp;
				</TD>
			  </TR>			  
			  
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="订单处理" class="button Noprint" type="submit">
		</div>
		
		<DIV>
			<TABLE cellSpacing="1" width="60%">
			  <TR class="content-td2">
			    <TH width="25%" align="left">产品名称</TH>
				<TH width="25%" align="left">产品数目</TH>
			    <TH width="20%" align="left">产品价格</TH>
			    <TH width="20%" align="left">促销手段</TH>
				<TH width="10%" align="left"></TH>
			  </TR>
			  <%
			  	'response.Write config_IsShowOrderItem
			  	Dim sql:sql = ""
			  	if Cint(config_IsShowOrderItem) = 1 then
					sql = "select O.*,P.title from user_pro P,Sp_OrderItem O where P.id=O.proID and O.OrderID="&ItemID&""
				else
					sql = "select 1 as num,'未知' as price, P.title from user_pro P,Sp_OrderForm O where P.id=O.proID and O.id="&ItemID&""
				end if
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><%=rs("title")%></TD>
				<td><%=rs("num")%>&nbsp;</td>
				<TD><%=rs("price")%>&nbsp;</TD>
				<TD></TD>
				<TD><%if config_IsShowOrderItem = 1 then%><a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&ItemID=<%=rs("orderID")%>&OrderItemID=<%=rs("id")%>';">删除</a><%end if%></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
%>
