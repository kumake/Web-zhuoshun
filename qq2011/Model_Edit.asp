<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="UserAuthority.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<%
	'action
	Dim action:action = VerificationUrlParam("action","string","")
	'��ϢID��ʶ
	Dim ItemID:ItemID = VerificationUrlParam("ID","int","0")
	'ģ��ID
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","0")
	''''��ģ��ID
	Dim ParentModelID
	''''��ģ��ID
	Dim SubModelID
	'�����ʱ���ݵĶ�ά����
	Dim Temparray
	'��Ÿ�ģ�����ݵĶ�ά����
	Dim Modelarray
	'ģ������
	Dim ModelName
	'��Ÿ�ģ���������ݱ�
	Dim ModelTableName
	'���ģ������ID
	Dim ModelCategoryID
	'������Ķ�ά����
	Dim ModelCategoryArray
	'��ŵ����Ķ�ά����
	Dim CityCategoryArray
	'ģ���Ƿ�����û���Ȩ�޻���
	Dim IsUserGroup:IsUserGroup = 0
	''�鿴ģ����Ϣ�Ƿ���Ҫ����
	Dim IsModelNeedScore
	'ģ�ʹ����û���Ȩ��
	Dim ModelUserGroup
	'ģ��Ȩ������
	Dim ModelUserGroupArray
	'''''''''''''''''''''''
	if ModelID=0 then
		response.Write "ģ�ͽ��ͳ��ִ���!"
	else
		'ȡ��Ӧģ�����ݺͶ�Ӧģ�͵��Զ����ֶ�
		Modelarray = Con.QueryData("select M.ID as MID,M.modelname,M.modeltable,M.modelCategoryID,F.ID,F.ModelID,F.fieldname,F.FieldUI,F.FieldDescription,F.fieldtype,F.fieldlength,F.isNull,F.Attribute,F.defaultValue,M.ParentModelID,M.SubModelID,M.isusergroup,M.usergroup,M.IsModelNeedScore from Sp_Model M,Sp_ModelField F where M.ID=F.ModelID and M.Id="&ModelID&" order by F.indexID asc") 
		'response.Write ubound(Modelarray,1)
		'response.End()
		if ubound(Modelarray,1)=0 then 
			'response.Write "���ִ���"
			'response.End()	
			TempModelarray = Con.QueryRow("select ID as MID,modelname,modeltable,modelCategoryID,ParentModelID,SubModelID,isusergroup,usergroup,IsModelNeedScore from Sp_Model where Id="&ModelID&"",0) 
			ModelName = TempModelarray(1)		
			ModelTableName = TempModelarray(2)		
			ModelCategoryID = TempModelarray(3)
			ParentModelID = TempModelarray(4)
			SubModelID = TempModelarray(5)
			IsUserGroup = TempModelarray(6)
			ModelUserGroup = TempModelarray(7)
			IsModelNeedScore = TempModelarray(8)
		else
			ModelName = Modelarray(1,0)		
			ModelTableName = Modelarray(2,0)		
			ModelCategoryID = Modelarray(3,0)
			ParentModelID = Modelarray(14,0)
			SubModelID = Modelarray(15,0)
			IsUserGroup = Modelarray(16,0)
			ModelUserGroup = Modelarray(17,0)
			IsModelNeedScore = Modelarray(18,0)
		end if
		'''response.Write ModelCategoryID
		'''ȡ���ݴ�Ž�ModelCategoryArray
		ModelCategoryArray = Con.QueryData("select D.ID,D.categoryID,D.categoryname,D.Pathstr,D.ParentID from Sp_dictionaryCategory C,Sp_dictionary D where C.id=D.categoryID and D.categoryID="&ModelCategoryID&" order by D.id desc")
		CityCategoryArray = Con.QueryData("select ID,Areaname,pathstr,parentID,Depth from Sp_city")
		'response.Write ubound(ModelCategoryArray,2)
		''''''������ϢID��ʶȡ��Ϣ������
		set rs = COn.Query("select MT.*,D.Pathstr from "&ModelTableName&" MT, Sp_dictionary D where D.ID=MT.categoryID and MT.ID="&ItemID&"")
		if rs.recordcount=0 then
			Alert "�Ҳ�����ӦItemID="&ItemID&"Ϊ��ʶ��������Ϣ!",""
		else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script src="Scripts/calender.js" type="text/javascript"></script>
    <script language="javascript" type="text/javascript" src="Scripts/popup.js"></script>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
	<!--��Щ�ֶα�����д��javascript����֤//-->
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			<%
			if ParentModelID<>0 then
				response.Write "if(document.getElementById('txt_SubModelItemID').value=='0')"
				response.Write "{"
				response.Write "	alert('Sp_CMS��ʾ\r\n\n��ѡ��Ҫ����ϵģ�������Ϣ�ı���!');"
				response.Write "	return false;"			
				response.Write "}"
			end if
			%>
			if(document.getElementById("txt_title").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n��Ϣ���������д!");
				return false;			
			}
			if (document.getElementById("ChangesLink").checked!=true)
			{
				<%
				''ѭ��
				if ubound(Modelarray,1)=0 then 
				else
					for i=0 to ubound(Modelarray,2) step 1
						if Modelarray(11,i)=1 then
							response.Write "if(document.getElementById('txt_"&Modelarray(6,i)&"').value=='')"&vbCrlf
							response.Write "{"&vbCrlf
							response.Write "	alert('Sp_CMS��ʾ\r\n\n"&Modelarray(8,i)&"������д!');"&vbCrlf
							response.Write "	return false;"&vbCrlf		
							response.Write "}"&vbCrlf
						end if
					next
				end if
				%>
			}
			else
			{
				return true;
			}
			return true;
		}
		///
		function ArticlePics(AC)
		{   
			var obj = document.getElementById("PicUpload");
			if (AC==1)
			{
				obj.style.display="";
			}
			else
			{
				obj.style.display="none";
			}
		}
		///
		function ChangesLinks()
		{ 
			var obj = document.getElementById("JumpUrl");
			if (document.getElementById("ChangesLink").checked==true)
			{
				obj.style.display="";
			}
			else
			{
				obj.style.display="none";
			}
		}
	</script>
</HEAD>
<BODY>
<form action="Model_Save.asp?ModelID=<%=ModelID%>&action=edit" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)"><%=ModelName%>�޸���Ϣ</A></LI>
	  <%if IsUserGroup<>0 then%>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent1',this)" href="javascript:void(0)">Ȩ������</A></LI>
	  <%end if%>
	</UL>
	<DIV id=tagContent>
	<%if IsUserGroup<>0 then%>
	<DIV class="tagContent selectTag content" id="tagContent1" style="display:none;">
		<div>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">�û���Ȩ�ޣ�</TD>
				<TD>
				<%				
				ModelUserGroupArray = split(rs("usergroup"),",")
				set rsg = con.query("select id,UserGroup from Sp_UserGroup")
				if rsg.recordcount<>0 then
				do while not rsg.eof
					response.Write "<input type='checkbox' name='txt_usergroup' value='"&rsg("id")&"'"
					for rspi=0 to ubound(ModelUserGroupArray)
					if trim(ModelUserGroupArray(rspi))=trim(rsg("id")) then 
					response.Write " checked"
					exit for
					end if
					next
					response.Write ">"&rsg("UserGroup")&"<br>"
				rsg.movenext
				loop
				end if
				%>
				</TD>
			  </TR>	
			  </table>		  
		</div>
	</div>
	<%end if%>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--��1����ǩ//-->
		<DIV>
			<input name="hid_ModelTable" type="hidden" value="<%=ModelTableName%>">
			<input name="hid_ModelItemID" type="hidden" value="<%=ItemID%>">
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD>�������ͣ�</TD>
				<TD>
				<input type="radio" Checked  name="ArticlePic" value="0" <%if rs("infoType")=0 then response.Write "checked"%> onclick="javascript:ArticlePics(this.value);"/>
				<label for="ArticlePic1">��ͨ����</label>
				<input   type="radio" name="ArticlePic" value="1" <%if rs("infoType")=1 then response.Write "checked"%> onclick="javascript:ArticlePics(this.value);"/>
				<label for="ArticlePic2">ͼƬ��õ�Ƭ����</label>
				<input  name="ChangesLink" type="checkbox" id="ChangesLink" value="1" <%if rs("IsJumpUrl")=1 then response.Write "checked"%> onclick="javascript:ChangesLinks();"/>
				<label for="ChangesLink"><font color="#FF0000"><b>ʹ��ת������</b></font></label>
	     		</TD>
			  </TR>			  
			  <TR class="content-td1" id="PicUpload" style="display:<%if rs("infoType")=1 then:response.Write "":else:response.Write "none":end if%>;">
				<TD>ͼƬ��ַ��</TD>
				<TD>
				<input name="txt_imgUrl" class="input" size="60" type="text" value="<%=rs("imgUrl")%>"><br>
				<iframe name="ad" frameborder="0"   width="100%" height="30px;" scrolling="no" src="File_Upload.asp?obj=txt_imgUrl"></iframe>
				</TD>
			  </TR>
			  <%
			  ''''�Ƿ���ʾ����
			  if config_isShowCity=1 then
			  %>
			  <TR class="content-td1">
				<TD width="15%">�������</TD>
				<TD>
				<select name="txt_CityID">
				<%filldictionaryDropDownList CityCategoryArray,0,rs("cityid")%>
				</select>
				</TD>
			  </TR>
			  <%end if%>
			  <TR class="content-td1">
				<TD width="15%">��Ϣ���</TD>
				<TD>
				<select name="CategoryID">
				<%fillDropDownList ModelCategoryArray,0,rs("CategoryID")%>
				</select>
				</TD>
			  </TR>
			  <!--���¹�����ҪΪ����ģ�ͷ���//-->
			  <%
			  if ParentModelID<>0 then
			  %>
			  <TR class="content-td1">
				<TD>����ģ����Ϣ��</TD>
				<TD>
				<input type="hidden" name="txt_SubModelID" value="<%=ParentModelID%>">
				<input type="hidden" name="txt_SubModelItemID" value="<%=rs("SubModelItemID")%>">
				<input name="txt_RelationTitle" disabled class="input" size="80" type="text" value="�Ѿ�ѡ��">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('Model_Parent.asp?ModelID=<%=ParentModelID%>&value=txt_SubModelItemID&text=txt_RelationTitle');" class="huitext">����ѡ��</a>
				</TD>
			  </TR>
			  <%
			  end if
			  %>
			  <%if IsModelNeedScore=1 then%>
			  <TR class="content-td1">
			    <TD>�鿴��Ҫ������Ŀ��</TD>
			    <TD><input name="txt_LookNeedScore" type="text" class="input" value="<%=rs("LookNeedScore")%>" size="6"></TD>
		      </TR>
			  <%end if%>
			  <!--���Ϲ��ܽ���//-->
			  <TR class="content-td1">
				<TD>���⣺</TD>
				<TD><input name="txt_title" class="input" size="80" type="text" value="<%=rs("title")%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <%
			  if config_Verison=2 then
			  %>
			  <TR class="content-td1">
				<TD><span class="huitext">����(Ӣ�İ�)��</span></TD>
				<TD><input name="txt_title_En" class="input" size="80" type="text" value="<%=rs("title_En")%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <%
			  end if
			  %>
			  <TR class="content-td1" id="JumpUrl" style="display:<%if rs("IsJumpUrl")=1 then:response.Write "":else:response.Write "none":end if%>;">
				<TD>ת�����ӵ�ַ��</TD>
				<TD><input name="txt_jumpUrl" class="input" size="80" type="text" value="<%=rs("jumpUrl")%>">&nbsp;<font color="#ff0000">���������ת�����ӣ���ô���¸�����д��Ч,�벻Ҫ��д!</font></TD>
			  </TR>
			  <%
			  if ubound(Modelarray,1)=0 then 
			  else
			  for i=0 to ubound(Modelarray,2)
			  %>
			  <TR class="content-td1">
				<TD><%=Modelarray(8,i)%>��</TD>
				<TD><%=FieldUIType(Modelarray(6,i),Modelarray(7,i),Modelarray(13,i),rs(""&Modelarray(6,i)&""),Modelarray(11,i),Modelarray(12,i))%></TD>
			  </TR>
			  <%
			  if config_Verison=2 then
			  %>
			  <TR class="content-td1">
				<TD><span class="huitext"><%=Modelarray(8,i)%>(Ӣ�İ�)��</span></TD>
				<TD><%=FieldUIType(Modelarray(6,i)&"_en",Modelarray(7,i),Modelarray(13,i),rs(""&Modelarray(6,i)&"_en"&""),Modelarray(11,i),Modelarray(12,i))%></TD>
			  </TR>
			  <%
			  end if
			  next
			  end if
			  %>
			  <TR class="content-td1">
				<TD>����:</TD>
				<TD>
				<input name="IsNew" type="checkbox" value="1" <%if rs("IsNew")=1 then response.Write "checked"%>>����&nbsp;&nbsp;
				<input name="IsCommand" type="checkbox" value="1" <%if rs("IsCommand")=1 then response.Write "checked"%>>�Ƽ�&nbsp;&nbsp;
				<input name="IsTop" type="checkbox" value="1" <%if rs("IsTop")=1 then response.Write "checked"%>>�ö�&nbsp;&nbsp;
				<input name="IsHots" type="checkbox" value="1" <%if rs("IsHot")=1 then response.Write "checked"%>>�ȵ�&nbsp;&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ʾ˳��</TD>
				<TD><input name="txt_IndexID" class="input" type="text" value="<%=rs("IndexID")%>"><span class="huitext">&nbsp;��ֵԽ��Խ��ǰ,��ͬʱ����������������.</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>������ʱ��</TD>
				<TD><input name="txt_lastUpdatetime" class="input" value="<%=year(rs("lastUpdatetime"))%>-<%=month(rs("lastUpdatetime"))%>-<%=day(rs("lastUpdatetime"))%>" type="text" onfocus='javascript:HS_setDate(this);'></TD>
			  </TR>
			</TABLE>
		</DIV>
	  </DIV>
	  <div class="divpadding">
	    <input name="btnsearch" value="�༭��Ϣ" class="button" type="submit">
	  </div>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
		end if
	end if
%>
<script language=javascript>
    //ѡ��ģ��
	var g_pop;
    function SelectTemplate(url)
    {
        g_pop=new Popup({ contentType:1, isReloadOnClose:false,scrollType:'no', width:700, height:420 });
        g_pop.setContent("title","ѡ�񸸽ڵ���Ϣ");
        g_pop.setContent("contentUrl",url);
    	
        g_pop.build();
        g_pop.show();
        return false;
    }
    
    //�رմ򿪵Ĵ���
    function ClosePop()
    {
        g_pop.close();
    }
</script>
