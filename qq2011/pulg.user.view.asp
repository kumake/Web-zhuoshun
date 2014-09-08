<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim GroupID:GroupID = VerificationUrlParam("GroupID","int","0")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	Dim action:action = VerificationUrlParam("action","string","")
	if ItemID=0 then
		Alert "出现错误",""
	else
	    ''''修改用户资料
		if action<>"" and action="update" then 
			Dim money:money = F.ReplaceFormText("txt_money")
			Dim score:score = F.ReplaceFormText("txt_score")
			Dim isactive:isactive = F.ReplaceFormText("txt_isactive")
			Dim isVip:isVip = F.ReplaceFormText("txt_isVip")
			dim sql:sql = "update sp_user set [money] = "&money&",score="&score&",isVip="&isVip&",isactive="&isactive&" where ID="&ItemID&""
			'response.Write sql
			'response.End()
			Con.execute(sql)
			Alert "修改资料成功!","pulg.user.view.asp?GroupID="&GroupID&"&ItemID="&ItemID&""
		end if
		
		'取对应模型数据和对应模型的自定义字段
		Dim ExpandArray:ExpandArray = Con.QueryData("select fieldname,FieldDescription from Sp_UserExpField where GroupID="&GroupID&"") 

		set rs = Con.Query("select U.*,G.UserGroup from Sp_User U,Sp_UserGroup G where U.GroupID=G.ID and U.ID="&ItemID&"")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript">
		function checkform()
		{
			if(!confirm("确定修改该用户资料?"))
			{
				return false;
			}
			else
			{
				return true;
			}
		}
	</script>
</HEAD>
<BODY>
<form action="?action=update&GroupID=<%=GroupID%>&ItemID=<%=ItemID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">用户管理:详细信息</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">用户组：</TD>
				<TD><%=rs("usergroup")%></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD width="15%">用户形象：</TD>
			    <TD>
				<%if isnull(rs("image")) or rs("image")="" then%>
				<img src="images/nopic.gif" width="128" style="border:1px solid #efefef;" height="128">
				<%else%>
				<img src="<%=rs("image")%>" width="128" style="border:1px solid #efefef;" height="128">
				<%end if%>
				</TD>
		      </TR>
			  <TR class="content-td1">
				<TD>用户名：</TD>
				<TD><%=rs("username")%>&nbsp;<span class="redtext"><%if rs("isvip")=1 then response.Write "Vip"%></span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>电子邮件：</TD>
				<TD><%=rs("useremail")%></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>注册时间：</TD>
				<TD><%=year(rs("regtime"))%>-<%=month(rs("regtime"))%>-<%=day(rs("regtime"))%></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>账户余额：</TD>
			    <TD><input name="txt_money" class="input" type="text" value="<%=rs("money")%>">&nbsp;&nbsp;<a href="pulg.user.score.transaction.asp?type=m&useruid=<%=ItemID%>">金额交易记录</a></TD>
		      </TR>
			  <TR class="content-td1">
			    <TD>网站积分：</TD>
			    <TD><input name="txt_score" class="input" type="text" value="<%=rs("score")%>">&nbsp;&nbsp;<a href="pulg.user.score.transaction.asp?type=s&useruid=<%=ItemID%>">积分交易记录</a></TD>
		      </TR>
			  <TR class="content-td1">
			    <TD>登录积分：</TD>
			    <TD class="redtext"><%=rs("logincount")%></TD>
		      </TR>
			  
			  
			  
			  <TR class="content-td1">
				<TD>最后登录时间：</TD>
				<TD><%=year(rs("lastlogintime"))%>-<%=month(rs("lastlogintime"))%>-<%=day(rs("lastlogintime"))%></TD>
			  </TR>			  
			  <TR class="content-td1">
				<TD>当前状态：</TD>
				<TD>
				<select name="txt_isactive">
				<option value="1" <%if rs("isactive")=1 then:response.Write "selected"%>>激活</option>
				<option value="0" <%if rs("isactive")=0 then:response.Write "selected"%>>禁止</option>
		        </select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>认证状态：</TD>
				<TD>
				<select name="txt_isVip">
				<option value="1" <%if rs("IsVip")=1 then:response.Write "selected"%>>Vip</option>
				<option value="0" <%if rs("IsVip")=0 then:response.Write "selected"%>>普通</option>
		        </select>
				<%if isnull(rs("isvipimg")) or rs("isvipimg")="" then%>
				<img src="images/nopic.gif" width="128" style="border:1px solid #efefef;" height="128">
				<%else%>
				<img src="<%=rs("isvipimg")%>" width="128" style="border:1px solid #efefef;" height="128">
				<%end if%>
				</TD>
			  </TR>
			  <TR class="content-td2">
				<th colspan="2"><span class="huitext">所属组自定义表单</span></th>
		      </TR>
			  <%
			  for i=0 to ubound(ExpandArray,2)
			  %>
			  <TR class="content-td1">
				<TD><%=ExpandArray(1,i)%>：</TD>
				<TD><%=rs(""&ExpandArray(0,i)&"")%></TD>
			  </TR>
			  <%
			  next
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="修改资料" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
%>
