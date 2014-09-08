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
							<h3 style="color:red;margin-left:20px; line-height:40px;padding-top:15px;">企业新闻</h3></td>
						</tr>
						<tr>
							<td width="772" height="514">
							<div id="neiye_main">
								<!--内容开始 -->
										<div style="line-height:30px; margin-top:20px ">
										<%  
										CategoryID = VerificationUrlParam("CategoryID","int","0")%>
										<%
										''''''采用该分页类型的好处：
										''常规分页方式采用一次性将数据读入内存通过游标显示每页记录
										''当数据表记录大于10万甚至1000万时？
										''采用一支笔该分页方式的好处就是需要多少条记录就从表中读多少条记录
										''根据测试数据记录越多效果越明显。测试：asp+sql2000+数据量500万
										'''''''
										Dim total_record   	'总记录数
										Dim current_page	'当前页面
										Dim PCount			'循环页显示数目
										Dim pagesize		'每页显示记录数
										Dim showpageNum:showpageNum = true		'是否显示数字循环页
										Dim showpagetotal:showpagetotal = true	'是否显示记录总数
										Dim IsEnglish:IsEnglish = false			'是否显示英文分页格式		
										Dim strwhere:strwhere = "isdel=0"		'查询条件
										'''获取查询条件
										''''总数记录只取一次节省数据库压力
										if CategoryID<>0 then 
										strwhere = strwhere & " and categoryID="&CategoryID&""
										
										end if
										Dim total:total = VerificationUrlParam("total","int","0")
										if total = 0 then 
										Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from user_news where "&strwhere&"")
										total_record = Tarray(0,0)
										else
										total_record = total
										end if
										'''''
										
										current_page = VerificationUrlParam("page","int","1")
										PCount = 6			'''分页循环显示记录数
										pagesize = 14		'''每页显示记录数
										'这种方式为根据ID为关键字排序
										'该中分页方式效果最好。建议使用,但是排序效果受到限制
										'Dim sql:sql = UICon.QueryPageByNum("categoryID,id,title,posttime","user_newss", ""&strwhere&"", "ID", true,current_page,pagesize)
										'这种方式为根据IndexID排序，IndexID值越大越靠前
										Dim sql:sql = UICon.QueryPageByNotIn("*","user_news", ""&strwhere&"", "ID", "indexID desc,ID",false,current_page,pagesize)
										'response.Write sql
										'response.End()
										set rsn = UICon.Query(sql)
										if rsn.recordcount<>0 then
										do while not rsn.eof
										%>
										<table width="90%" border="0" cellspacing="0" cellpadding="0" style="margin-left:20px;">
										<tr style="border-bottom:#E4E4E4 solid 1px; line-height:30px;">
										<td height="30" align="left" valign="middle" class="news_title" style="border-bottom:#E4E4E4 solid 1px; line-height:30px;">					<span style="float:right"><%=year(rsn("PostTime"))%>-<%=month(rsn("PostTime"))%>-<%=day(rsn("PostTime"))%></span>
										<span style="font-family:'宋体'">·</span><a href="news_in.asp?categoryid=<%=rsn("categoryID")%>&amp;itemid=<%=rsn("id")%>" class="y_0"><%=rsn("title")%></a> </td>
										</tr>
										<tr>
										<td height="3" align="left" valign="middle" background="images/line778.gif"></td>
										</tr>
										</table>
										<%
										rsn.movenext
										loop
										
										%>
									
										<% end if %>
										 <div style="line-height:30px; text-align:center; width:100%;"> <%PageIndexUrl total_record,current_page,PCount,pagesize,showpageNum,showpagetotal,IsEnglish%></div>
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
