<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="UserAuthority.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<%
	'action
	Dim action:action = VerificationUrlParam("action","string","")
	'信息ID标识
	Dim ItemID:ItemID = VerificationUrlParam("ID","int","0")
	'模型ID
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","0")
	''''父模型ID
	Dim ParentModelID
	''''子模型ID
	Dim SubModelID
	'存放临时数据的二维数组
	Dim Temparray
	'存放该模型数据的二维数组
	Dim Modelarray
	'模型名称
	Dim ModelName
	'存放该模型数据数据表
	Dim ModelTableName
	'存放模型类别的ID
	Dim ModelCategoryID
	'存放类别的二维数组
	Dim ModelCategoryArray
	'存放地区的二维数组
	Dim CityCategoryArray
	'模型是否带入用户组权限机制
	Dim IsUserGroup:IsUserGroup = 0
	''查看模型信息是否需要积分
	Dim IsModelNeedScore
	'模型带入用户组权限
	Dim ModelUserGroup
	'模型权限数组
	Dim ModelUserGroupArray
	'''''''''''''''''''''''
	if ModelID=0 then
		response.Write "模型解释出现错误!"
	else
		'取对应模型数据和对应模型的自定义字段
		Modelarray = Con.QueryData("select M.ID as MID,M.modelname,M.modeltable,M.modelCategoryID,F.ID,F.ModelID,F.fieldname,F.FieldUI,F.FieldDescription,F.fieldtype,F.fieldlength,F.isNull,F.Attribute,F.defaultValue,M.ParentModelID,M.SubModelID,M.isusergroup,M.usergroup,M.IsModelNeedScore from Sp_Model M,Sp_ModelField F where M.ID=F.ModelID and M.Id="&ModelID&" order by F.indexID asc") 
		'response.Write ubound(Modelarray,1)
		'response.End()
		if ubound(Modelarray,1)=0 then 
			'response.Write "出现错误"
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
		'''取数据存放进ModelCategoryArray
		ModelCategoryArray = Con.QueryData("select D.ID,D.categoryID,D.categoryname,D.Pathstr,D.ParentID from Sp_dictionaryCategory C,Sp_dictionary D where C.id=D.categoryID and D.categoryID="&ModelCategoryID&" order by D.id desc")
		CityCategoryArray = Con.QueryData("select ID,Areaname,pathstr,parentID,Depth from Sp_city")
		'response.Write ubound(ModelCategoryArray,2)
		''''''根据信息ID标识取信息的内容
		set rs = COn.Query("select MT.*,D.Pathstr from "&ModelTableName&" MT, Sp_dictionary D where D.ID=MT.categoryID and MT.ID="&ItemID&"")
		if rs.recordcount=0 then
			Alert "找不到对应ItemID="&ItemID&"为标识的数据信息!",""
		else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script src="Scripts/calender.js" type="text/javascript"></script>
    <script language="javascript" type="text/javascript" src="Scripts/popup.js"></script>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
	<!--有些字段必须填写的javascript的验证//-->
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			<%
			if ParentModelID<>0 then
				response.Write "if(document.getElementById('txt_SubModelItemID').value=='0')"
				response.Write "{"
				response.Write "	alert('Sp_CMS提示\r\n\n请选择要给关系模型添加信息的标题!');"
				response.Write "	return false;"			
				response.Write "}"
			end if
			%>
			if(document.getElementById("txt_title").value=="")
			{
				alert("Sp_CMS提示\r\n\n信息标题必须填写!");
				return false;			
			}
			if (document.getElementById("ChangesLink").checked!=true)
			{
				<%
				''循环
				if ubound(Modelarray,1)=0 then 
				else
					for i=0 to ubound(Modelarray,2) step 1
						if Modelarray(11,i)=1 then
							response.Write "if(document.getElementById('txt_"&Modelarray(6,i)&"').value=='')"&vbCrlf
							response.Write "{"&vbCrlf
							response.Write "	alert('Sp_CMS提示\r\n\n"&Modelarray(8,i)&"必须填写!');"&vbCrlf
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)"><%=ModelName%>修改信息</A></LI>
	  <%if IsUserGroup<>0 then%>
	  <LI><A onfocus="this.blur()" onclick="selectTag('tagContent1',this)" href="javascript:void(0)">权限设置</A></LI>
	  <%end if%>
	</UL>
	<DIV id=tagContent>
	<%if IsUserGroup<>0 then%>
	<DIV class="tagContent selectTag content" id="tagContent1" style="display:none;">
		<div>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">用户组权限：</TD>
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
		<!--第1个标签//-->
		<DIV>
			<input name="hid_ModelTable" type="hidden" value="<%=ModelTableName%>">
			<input name="hid_ModelItemID" type="hidden" value="<%=ItemID%>">
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD>文章类型：</TD>
				<TD>
				<input type="radio" Checked  name="ArticlePic" value="0" <%if rs("infoType")=0 then response.Write "checked"%> onclick="javascript:ArticlePics(this.value);"/>
				<label for="ArticlePic1">普通文章</label>
				<input   type="radio" name="ArticlePic" value="1" <%if rs("infoType")=1 then response.Write "checked"%> onclick="javascript:ArticlePics(this.value);"/>
				<label for="ArticlePic2">图片或幻灯片文章</label>
				<input  name="ChangesLink" type="checkbox" id="ChangesLink" value="1" <%if rs("IsJumpUrl")=1 then response.Write "checked"%> onclick="javascript:ChangesLinks();"/>
				<label for="ChangesLink"><font color="#FF0000"><b>使用转向链接</b></font></label>
	     		</TD>
			  </TR>			  
			  <TR class="content-td1" id="PicUpload" style="display:<%if rs("infoType")=1 then:response.Write "":else:response.Write "none":end if%>;">
				<TD>图片地址：</TD>
				<TD>
				<input name="txt_imgUrl" class="input" size="60" type="text" value="<%=rs("imgUrl")%>"><br>
				<iframe name="ad" frameborder="0"   width="100%" height="30px;" scrolling="no" src="File_Upload.asp?obj=txt_imgUrl"></iframe>
				</TD>
			  </TR>
			  <%
			  ''''是否显示地区
			  if config_isShowCity=1 then
			  %>
			  <TR class="content-td1">
				<TD width="15%">地区类别：</TD>
				<TD>
				<select name="txt_CityID">
				<%filldictionaryDropDownList CityCategoryArray,0,rs("cityid")%>
				</select>
				</TD>
			  </TR>
			  <%end if%>
			  <TR class="content-td1">
				<TD width="15%">信息类别：</TD>
				<TD>
				<select name="CategoryID">
				<%fillDropDownList ModelCategoryArray,0,rs("CategoryID")%>
				</select>
				</TD>
			  </TR>
			  <!--以下功能主要为关联模型服务//-->
			  <%
			  if ParentModelID<>0 then
			  %>
			  <TR class="content-td1">
				<TD>关联模型信息：</TD>
				<TD>
				<input type="hidden" name="txt_SubModelID" value="<%=ParentModelID%>">
				<input type="hidden" name="txt_SubModelItemID" value="<%=rs("SubModelItemID")%>">
				<input name="txt_RelationTitle" disabled class="input" size="80" type="text" value="已经选择">&nbsp;&nbsp;<a href="#" onClick="javascript:SelectTemplate('Model_Parent.asp?ModelID=<%=ParentModelID%>&value=txt_SubModelItemID&text=txt_RelationTitle');" class="huitext">重新选择</a>
				</TD>
			  </TR>
			  <%
			  end if
			  %>
			  <%if IsModelNeedScore=1 then%>
			  <TR class="content-td1">
			    <TD>查看需要积分数目：</TD>
			    <TD><input name="txt_LookNeedScore" type="text" class="input" value="<%=rs("LookNeedScore")%>" size="6"></TD>
		      </TR>
			  <%end if%>
			  <!--以上功能结束//-->
			  <TR class="content-td1">
				<TD>标题：</TD>
				<TD><input name="txt_title" class="input" size="80" type="text" value="<%=rs("title")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <%
			  if config_Verison=2 then
			  %>
			  <TR class="content-td1">
				<TD><span class="huitext">标题(英文版)：</span></TD>
				<TD><input name="txt_title_En" class="input" size="80" type="text" value="<%=rs("title_En")%>"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <%
			  end if
			  %>
			  <TR class="content-td1" id="JumpUrl" style="display:<%if rs("IsJumpUrl")=1 then:response.Write "":else:response.Write "none":end if%>;">
				<TD>转向链接地址：</TD>
				<TD><input name="txt_jumpUrl" class="input" size="80" type="text" value="<%=rs("jumpUrl")%>">&nbsp;<font color="#ff0000">如果文章是转向链接，那么以下各项填写无效,请不要填写!</font></TD>
			  </TR>
			  <%
			  if ubound(Modelarray,1)=0 then 
			  else
			  for i=0 to ubound(Modelarray,2)
			  %>
			  <TR class="content-td1">
				<TD><%=Modelarray(8,i)%>：</TD>
				<TD><%=FieldUIType(Modelarray(6,i),Modelarray(7,i),Modelarray(13,i),rs(""&Modelarray(6,i)&""),Modelarray(11,i),Modelarray(12,i))%></TD>
			  </TR>
			  <%
			  if config_Verison=2 then
			  %>
			  <TR class="content-td1">
				<TD><span class="huitext"><%=Modelarray(8,i)%>(英文版)：</span></TD>
				<TD><%=FieldUIType(Modelarray(6,i)&"_en",Modelarray(7,i),Modelarray(13,i),rs(""&Modelarray(6,i)&"_en"&""),Modelarray(11,i),Modelarray(12,i))%></TD>
			  </TR>
			  <%
			  end if
			  next
			  end if
			  %>
			  <TR class="content-td1">
				<TD>属性:</TD>
				<TD>
				<input name="IsNew" type="checkbox" value="1" <%if rs("IsNew")=1 then response.Write "checked"%>>最新&nbsp;&nbsp;
				<input name="IsCommand" type="checkbox" value="1" <%if rs("IsCommand")=1 then response.Write "checked"%>>推荐&nbsp;&nbsp;
				<input name="IsTop" type="checkbox" value="1" <%if rs("IsTop")=1 then response.Write "checked"%>>置顶&nbsp;&nbsp;
				<input name="IsHots" type="checkbox" value="1" <%if rs("IsHot")=1 then response.Write "checked"%>>热点&nbsp;&nbsp;
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>显示顺序</TD>
				<TD><input name="txt_IndexID" class="input" type="text" value="<%=rs("IndexID")%>"><span class="huitext">&nbsp;数值越大越靠前,相同时候其他参数其作用.</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>最后更新时间</TD>
				<TD><input name="txt_lastUpdatetime" class="input" value="<%=year(rs("lastUpdatetime"))%>-<%=month(rs("lastUpdatetime"))%>-<%=day(rs("lastUpdatetime"))%>" type="text" onfocus='javascript:HS_setDate(this);'></TD>
			  </TR>
			</TABLE>
		</DIV>
	  </DIV>
	  <div class="divpadding">
	    <input name="btnsearch" value="编辑信息" class="button" type="submit">
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
    //选择模板
	var g_pop;
    function SelectTemplate(url)
    {
        g_pop=new Popup({ contentType:1, isReloadOnClose:false,scrollType:'no', width:700, height:420 });
        g_pop.setContent("title","选择父节点信息");
        g_pop.setContent("contentUrl",url);
    	
        g_pop.build();
        g_pop.show();
        return false;
    }
    
    //关闭打开的窗口
    function ClosePop()
    {
        g_pop.close();
    }
</script>
