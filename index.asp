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
<p>�����Ͼ�׿˳��װ�䳧���Ͼ�����רҵ��ľ�ʰ�װ���ϼӹ���ҵ���ҳ�רҵ�������۴�ͳľ�ʰ�װ�䡢���ͼа��װ�䡢Χ��ľ�䡢��Ѭ�����̰�װ�䡢����ľ���̺ͽ������̼�������̵ȡ���Ʒ�㷺���ڻ�е��ҽ����е��������Ʒ�������Ʒ���������ϵİ�װ���ҳ���Ʒ��ר�����ڳ��ڰ�װ�����ȴ������Ѭ���������ࡣ ���ɸ��ݿͻ���Ҫ�ṩ������װ����(��ֽ�䡢����Ĥ����ĭ���ǡ��������Ĥ)���ֳ�ʩ�������ҳ������˽�ȫ������������ϵ���ϸ�ִ��һ���׹�</p>			
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
									<!--���ݿ�ʼ -->
								<script src="JS/MSClass.js"></script>
                        <div id="marqueediv6" style=" text-align:center;width:700px;height:159px;margin-left:13px;">
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
						  
                           <%

				set rs = UICon.QUery("select top 20 * from user_pro order by hots desc ")
				if rs.recordcount<>0 then
				do while not rs.eof
				'''''''''��ô�ֶ���''''''
				''��ҳ�����DIV+css�����Է���ʵ�������ǳ����㲻��Ҫ��ҳ��
				''ֻ��Ҫ�ı�css�� procontent ��ǩ�µ�li�Ŀ�ȼ���
				''һ�� ֻҪprocontent �Ŀ�Ⱥ�li�Ŀ��һ��
				''��Ҫ���� ֻҪ��li�Ŀ������Ϊ procontent�ļ���֮��
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
						<!--���ݽ��� -->
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
												<li><em>[<%=year(rs("PostTime"))%>-<%=month(rs("PostTime"))%>-<%=day(rs("PostTime"))%>]</em>����<a  href="news_in.asp?categoryid=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" title="<%=rs("title")%>"  target="_blank" ><%=rs("title")%></a></li>	
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
											<li>[����ľ��] �Ͼ�׿˳��װ�䳧���Ͼ�����ר</li>
											<li><p>2012-02-13</p></li>
											<li>[����ľ��] �Ͼ�׿˳��װ�䳧���Ͼ�����ר</li>
											<li><p>2012-02-13</p></li>
											<li>[����ľ��] �Ͼ�׿˳��װ�䳧���Ͼ�����ר</li>
											<li><p>2012-02-13</p></li>
											<li>[����ľ��] �Ͼ�׿˳��װ�䳧���Ͼ�����ר</li>
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
