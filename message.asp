<!--#include file="conn.asp"-->
<!--#include file="sp_inc/class_page.asp"-->
<!--#include file="plugIn/Setting.Config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=config_sitename%></title>
<meta name="keywords" content="<%=config_seokeyword%>">
<meta name="description" content="<%=config_seocontent%>">
<link href="css/public.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.proLi{ width:160px; line-height:30px; border-bottom:#CCCCCC solid 1px; display:block; background:url(images/li.jpg) no-repeat 30px 50%; padding-left:50px; margin-left:32px;}
 -->
</style>
</head>
<body>
<div id="container">
<table id="__01" width="984"  border="0" cellpadding="0" cellspacing="0">
	<tr>
		
		<td width="984" height="123">
		<!--#include file="head.asp" -->
		</td>
	</tr>
	<tr>
		<td>
		<table id="__01" width="984" height="575" border="0" cellpadding="0" cellspacing="0">
			<tr>
			  <td>
							<!--#include file="left.asp" --></td>
				<td>			
					<table id="__01" width="772" height="575" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td background="images/neiye_d_02_01.jpg" width="772" height="61">
							<h3 style="color:red;margin-left:20px;line-height:40px;padding-top:15px;">在线留言</h3></td>
						</tr>
						<tr>
							<td width="772" height="514">
							<div id="neiye_main">

								<!--内容开始 -->
									<div style="margin-left:20px; margin-top:20px; line-height:30px;">
												<%
									''''''''''''''''
									''以下代码必须包含
									''''''''''''''''
									Dim action:action = VerificationUrlParam("action","string","")
									Dim bookid:bookid = VerificationUrlParam("bookid","int","9")
									if action<>"" and action="save" then 
										if config_guestbookStatus=0 then 
											Alert "留言版已经关闭",""
											response.End()
										end if
										dim guestCategoryID:guestCategoryID = bookid
										dim username:username = request.Form("txt_username")
										dim tel:tel = request.Form("txt_tel")
										dim email:email = request.Form("txt_email")
										dim content:content = request.Form("txt_content")
										dim isaudit:isaudit = 1
										dim userip:userip=request.ServerVariables("REMOTE_ADDR")
										''''过滤
										if config_guestbookSplit = 1 then 
											Dim splitWord:splitWord = split(config_guestSplitWord,",")
											for i=0 to ubound(splitWord)
												content = replace(content,splitWord(i),"***脏话(危险字符)过滤***")
											next
										end if
										''''验证码
										if config_guestbookValidation = 1 then
											Dim veriCode:veriCode = request.Form("txt_veriCode")
											if veriCode<>session("GetCode") then
												Alert "验证码输入错误!",""
												response.End()
											end if
										end if 
										''''审核
										if config_guestbookAudit=1 then isaudit = 0
										
										''''IP禁止
										if config_IsprohibitIP = 1 then
											if instr(config_prohibitIP,userip)>0 then
												Alert "对不起你的Ip在禁止对列，不能留言!",""
												response.End()
											end if
										end if
										
										UICon.execute("insert into Sp_GuestBook (guestCategoryID,username,tel,email,content,ip,isaudit) values ("&guestCategoryID&",'"&username&"','"&tel&"','"&email&"','"&content&"','"&userip&"',"&isaudit&")")
										Alert "留言成功","message.asp"
									end if
									''''''''''''''''
									''以上代码必须包含
									''''''''''''''''
								%>
								
								<script language="javascript" type="text/javascript">
										function checkform()
										{
											if(document.forms[1].txt_username.value=='')
											{
												alert('Sp_CMS提示\r\n\n用户名必须填写!');
												return false;
											}
											if(document.forms[1].txt_tel.value=='')
											{
												alert('Sp_CMS提示\r\n\n联系方式必须填写!');
												return false;
											}
											if(document.forms[1].txt_content.value=='')
											{
												alert('Sp_CMS提示\r\n\n内容必须填写!');
												return false;
											}
											return true;
										}
									</script>
									<table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
										  <tr>
											<td height="28" style="padding-bottom:20px;">
											  <form action='?action=save&amp;bookid=<%=bookid%>' method='post'  onsubmit='javascript:return checkform();'>
												<table width="90%" align="center" cellpadding="1" cellspacing="1">
												  <tr class="content-td1">
													<td colspan="2"></td>
													</tr>
												  <tr class="content-td1">
													<td height="20">用户名</td>
													<td height="30"><input name="txt_username" type="text" class="input" value="" size="25" />
													  <span class='huitext STYLE1'><span class='huitext  STYLE3'>*</span></span></td>
													</tr>
												  <tr class="content-td1">
													<td height="20">联系方式</td>
													<td height="30"><input name='txt_tel' class='input' value='' size="25" />
													  <span class='huitext STYLE1'><span class='huitext  STYLE3'>*</span></span></td>
													</tr>
												  <tr class="content-td1">
													<td height="20">E-mail</td>
													<td height="30"><input name='txt_email' class='input' value='' size="25" /></td>
													</tr>
												  <tr class="content-td1">
													<td height="20">内容</td>
													<td><textarea name='txt_content' cols="50" rows="5" class='textarea'></textarea>
													  <span class='huitext STYLE1'><span class='huitext  STYLE3'>*</span></span></td>
													</tr>
												  <%if config_guestbookValidation = 1 then%>
												  <tr class="content-td1">
													<td height="20">验证码</td>
													<td><input name='txt_veriCode' class='input' value='' size="8" maxlength="4" />
													  <img align="absmiddle" onClick="javascript:this.src='Sp_inc/class_Checkcode.asp';" style="border:1px solid #999999; cursor:hand;" src="Sp_inc/class_Checkcode.asp" /></td>
													</tr>
												  <%end if%>
												  <tr>
													<td colspan='2'><input type="submit" name="button" id="button" value="提交" /></td>
													</tr>
												  </table>
												</form></td>
										  </tr>
										</table> <%
											''''''是否显示留言板列表
											if config_IsShowguestbookList=1 then
											%>
										  <table width="633" border="0" align="center" cellpadding="0" cellspacing="0">
											<tr>
											  <td height="22" bgcolor="#D4E5EF" class="left_topBg">&nbsp;&nbsp;留言列表</td>
											</tr>
											<tr>
											  <td height="28" style="padding-bottom:20px;"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
												<%
												Dim total_record   	'总记录数
												Dim current_page	'当前页面
												Dim PCount			'循环页显示数目
												Dim pagesize		'每页显示记录数
												Dim showpageNum		'是否显示数字循环页
												Dim strwhere:strwhere = " guestCategoryID="&bookid&""
												if config_guestbookAudit=1 then strwhere = strwhere &" and isaudit=1"
												if total = 0 then 
													Dim Tarray:Tarray = UICon.QUeryData("select count(*) as total from Sp_GuestBook where "&strwhere&"")
													total_record = Tarray(0,0)
												else
													total_record = total
												end if
												current_page = VerificationUrlParam("page","int","1")
												PCount = 6
												pagesize = 11
												showpageNum = true
												
												'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
												Dim sql:sql = UICon.QueryPageByNum("*","Sp_GuestBook", ""&strwhere&"", "ID", false,current_page,pagesize)
												'response.Write sql
												'response.End()
												set rs = UICon.QUery(sql)
												if rs.recordcount<>0 then
												do while not rs.eof
											  %>
												<tr>
												  <td height="30" style="padding-left:8px;">&nbsp;&nbsp;留言人:<u><%=rs("username")%></u>&nbsp;&nbsp;&nbsp;<span class="huitext">留言时间:<%=rs("addtime")%></span></td>
												</tr>
												<tr>
												  <td style="padding:8px; font-size:13px;"><%=rs("content")%></td>
												</tr>
												<tr>
												  <td><%if rs("replay")<>"" then response.Write "<span style='color:red;'>管理员回复:"&server.HTMLEncode(rs("replay"))&"</span>"%></td>
												</tr>
												<tr>
												  <td background="images/dotline.gif" height="1"></td>
												</tr>
												<%
												   rs.movenext
												   loop
												   end if
												  %>
												<tr>
												  <td><%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,false%></td>
												</tr>
											  </table></td>
											</tr>
										  </table>
										  <%end if%>
												</div>
								<!--内容结束 -->

							</div>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<!--#include file="footer.asp" -->
		</td>
	</tr>
</table>
</div>
</body>
</html>
