							<table id="__01" width="212" height="575" border="0" cellpadding="0" cellspacing="0">
								<tr>
									<td>
										<img src="images/left_01.jpg" width="212" height="43" alt=""></td>
								</tr>
								<tr>
									<td background="images/left_02.jpg" width="212" height="286" >
									<div id="left_fl">
											
											<ul class="lei_01">
													<%
													set rsn = UICon.Query("select  * from Sp_dictionary where  categoryID = 12 and parentID=0 order by  IndexID") 
													do while not rsn.eof	
													%>
														<li>������<a><%=rsn("categoryname")%></a></li>
														<ul class="lei_02">
														<!--��ʼ -->
															<%
															dim cateid:cateid=rsn("id")
															dim tmpsql:tmpsql="select  * from Sp_dictionary where  categoryID = 12 and parentID="&cateid&" order by  IndexID"
															set rs = UICon.Query(tmpsql)
															do while not rs.eof			
															%>
															<li>��������<a  href="product.asp?categoryid=<%=rs("id")%>&amp;itemid=&lei=<%=rs("categoryname")%>" title="<%=rs("categoryname")%>"   ><%=rs("categoryname")%></a></li>	
															<%
																rs.movenext
																loop
																rs.close
																set rs=nothing
															%>
														<!--���� -->
														</ul>										
													<%
														rsn.movenext
														loop
														rsn.close
														set rsn=nothing
													%>
											</ul>											
											
											
											
									</div>
									</td>
								</tr>
								<tr>
									<td>
										<img src="images/left_03.jpg" width="212" height="20" alt=""></td>
								</tr>
								<tr>
									<td>
										<img src="images/left_04.jpg" width="212" height="54" alt=""></td>
								</tr>
								<tr>
									<td background="images/left_05.jpg" width="212" height="160">
									<div id="left_lxwm">
									<p><strong>�� ϵ �ˣ�</strong>������</p>
									<p><strong>�ƶ��绰��</strong>13951766516</p>
									<p><strong>��    ����</strong>+86-025-85575619</p>
									<p><strong>�������棺</strong>+86-025-85575619</p>
									<p><strong>�����ʼ���</strong>sunxiaoming1978@yahoo.cn</p>
									<p><strong>�������룺</strong>210000</p>
									<p><strong>�ء���ַ��</strong>�Ͼ��н����������ֵ���ׯ����ׯ������</p>
										
									</div>
									</td>
								</tr>
								<tr>
									<td>
										<img src="images/left_06.jpg" width="212" height="12" alt=""></td>
								</tr>
							</table>