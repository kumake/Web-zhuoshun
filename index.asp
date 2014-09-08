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
		<table id="__01" width="984" height="791" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td>
				<table id="__01" width="984" height="575" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<!--#include file="left.asp" -->
						</td>
						<td>
						<table id="__01" width="772" height="575" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td>
									<img src="images/index_main_01.jpg" width="56" height="40" alt=""></td>
								<td>
									<img src="images/index_main_02.jpg" width="523" height="40" alt=""></td>
								<td>
									<img src="images/index_main_03.jpg" width="193" height="40" alt=""></td>
							</tr>
							<tr>
								<td>
									<img src="images/index_main_04.jpg" width="56" height="295" alt=""></td>
								<td>
									<img src="images/index_main_05.jpg" width="523" height="295" alt=""></td>
								<td background="images/index_main_06.jpg" width="193" height="295">
								<div style="margin-left:9px;margin-right:3px; height:295px;line-height:20px;">
<p>　　南京卓顺包装箱厂是南京地区专业的木质包装材料加工企业。我厂专业生产销售传统木质包装箱、新型夹板包装箱、围板木箱、免熏蒸托盘包装箱、各类木托盘和金属托盘及轻钢托盘等。产品广泛用于机械、医疗器械、化工产品、五金制品、建筑材料的包装。我厂产品中专门用于出口包装的有热处理和免熏蒸二个大类。 并可根据客户需要提供其他包装材料(如纸箱、塑料膜、泡沫护角、真空铝塑膜)和现场施工服务。我厂建立了健全的质量管理体系，严格执行一整套关</p>			
								</div>	
									</td>
							</tr>
							<tr>
								<td>
									<img src="images/index_main_07.jpg" width="56" height="56" alt=""></td>
								<td>
									<img src="images/index_main_08.jpg" width="523" height="56" alt=""></td>
								<td>
									<img src="images/index_main_09.jpg" width="193" height="56" alt=""></td>
							</tr>
							<tr>
								<td>
									<img src="images/index_main_10.jpg" width="56" height="162" alt=""></td>
								<td colspan="2"  width="716" height="162">
									<!--内容开始 -->
								<script src="JS/MSClass.js"></script>
                        <div id="marqueediv6" style=" text-align:center;width:700px;height:159px;margin-left:13px;">
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
						  
                           <%

				set rs = UICon.QUery("select top 20 * from user_pro order by hots desc ")
				if rs.recordcount<>0 then
				do while not rs.eof
				'''''''''怎么分多列''''''
				''该页面采用DIV+css。所以分列实现起来非常方便不需要改页面
				''只需要改变css中 procontent 标签下的li的宽度即可
				''一列 只要procontent 的宽度和li的宽度一致
				''需要几列 只要将li的宽度设置为 procontent的几分之几
			   %>
                            <td width="122"><table width="122" border="0" align="center" cellpadding="0" cellspacing="0"  height="122">
                                <tr>
                                  <td width="122"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>"><img src="<%=rs("Field_picture")%>"  height="120" ; width="140"   border="0" style="margin-top:5px"/></a>
								  <a href="product__in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" style="display:block; text-align:center; line-height:20px; color:#000; margin-top:5px;"><%=rs("title")%></a>								  </td>
                                </tr>
                            </table></td>
                            <td width="10">&nbsp;</td>
              <%
				rs.movenext
				loop
				end if
				%>              
                          </tr>
          </table>
        </div>
                        <script>new Marquee("marqueediv6",2,1,710,150,20,0,0)</script>
						<!--内容结束 -->
								</td>
							</tr>
							<tr>
								<td colspan="3">
									<img src="images/index_main_12.jpg" width="772" height="22" alt=""></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td>
					<table id="__01" width="984" height="216" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td><img src="images/index_02_02_01.jpg" width="224" height="216" alt="" /></td>
							<td>
								<table id="__01" width="464" height="216" border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td>
												<img src="images/index_gg_01.jpg" width="464" height="59" alt=""></td>
										</tr>
										<tr>
											<td  bgcolor="#F7F7F7" width="464" height="149">
											<ul id="xwzx">

											<%
												set rs = UICon.Query("select top 6 * from user_news order by id desc")
												do while not rs.eof
											
											%>
												<li><em>[<%=year(rs("PostTime"))%>-<%=month(rs("PostTime"))%>-<%=day(rs("PostTime"))%>]</em>　　<a  href="news_in.asp?categoryid=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" title="<%=rs("title")%>"  target="_blank" ><%=rs("title")%></a></li>	
											<%
												rs.movenext
												loop
												rs.close
												set rs=nothing
											%>	
											
											</ul>	</td>
										</tr>
										<tr>
											<td>
												<img src="images/index_gg_03.jpg" width="464" height="8" alt=""></td>
										</tr>
							  </table>							</td>
							<td>
								<table id="__01" width="296" height="216" border="0" cellpadding="0" cellspacing="0">
									<tr>
										<td>
											<img src="images/index_ss_01.jpg" width="296" height="59" alt=""></td>
									</tr>
									<tr>
										<td  background="images/index_ss_02.jpg" width="296" height="139">
										<ul id="rmsj">
											<li>[出口木箱] 南京卓顺包装箱厂是南京地区专</li>
											<li><p>2012-02-13</p></li>
											<li>[出口木箱] 南京卓顺包装箱厂是南京地区专</li>
											<li><p>2012-02-13</p></li>
											<li>[出口木箱] 南京卓顺包装箱厂是南京地区专</li>
											<li><p>2012-02-13</p></li>
											<li>[出口木箱] 南京卓顺包装箱厂是南京地区专</li>
											<li><p>2012-02-13</p></li>
										</ul>										</td>
									</tr>
									<tr>
										<td>
											<img src="images/index_ss_03.jpg" width="296" height="18" alt=""></td>
									</tr>
								</table>							</td>
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
