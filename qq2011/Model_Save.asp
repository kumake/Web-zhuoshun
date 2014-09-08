<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="UserAuthority.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","0")
	''''
	if ModelID=0 then
		response.Write "模型解释出现错误!"
	else
		'取对应模型数据和对应模型的自定义字段
		Modelarray = Con.QueryData("select M.ID as MID,M.modelname,M.modeltable,M.modelCategoryID,F.ID,F.ModelID,F.fieldname,F.FieldUI,F.FieldDescription,F.fieldtype,F.fieldlength,F.isNull,F.Attribute,F.defaultValue from Sp_Model M left join Sp_ModelField F on M.ID=F.ModelID where M.Id="&ModelID&"") 
		'response.Write ubound(Modelarray,1)
		'response.End()
		if ubound(Modelarray,1)=0 then 
			response.Write "出现错误"
			response.End()	
		end if
		ModelName = Modelarray(1,0)		
		ModelTableName = Modelarray(2,0)		
		ModelCategoryID = Modelarray(3,0)
	end if
	
	''''获取控件值
	Dim CityID:CityID = 1
	if config_isShowCity=1 then CityID = F.ReplaceFormText("txt_cityID")
	Dim ItemID:ItemID = F.ReplaceFormText("hid_ModelItemID")
	Dim InfoType:InfoType = F.ReplaceFormText("ArticlePic")
	Dim ImgUrl:ImgUrl = F.ReplaceFormText("txt_imgUrl")
	DIm OrderIndex:OrderIndex = F.ReplaceFormText("txt_IndexID")
	if InfoType = 0 then ImgUrl = ""
	Dim IsJumpUrl:IsJumpUrl = F.ReplaceFormText("ChangesLink")
	if IsJumpUrl="" then IsJumpUrl = 0
	Dim JumpUrl:JumpUrl = F.ReplaceFormText("txt_jumpUrl")
	if IsJumpUrl = 0 then JumpUrl = ""
	Dim CategoryID:CategoryID = F.ReplaceFormText("CategoryID")
	Dim title:title = F.ReplaceFormText("txt_title")
	Dim title_en:title_en = F.ReplaceFormText("txt_title_en")
	Dim UserUID:UserUID = 0
	if CK("userid")<>"" and isnumeric(CK("userid")) then UserUID = CK("userid")	
	Dim UserName:UserName = "游客"
	if CK("username")<>"" then UserName = CK("username")	
	Dim IsNew:IsNew = F.ReplaceFormText("IsNew")
	Dim IsCommand:IsCommand = F.ReplaceFormText("IsCommand")
	Dim IsTop:IsTop= F.ReplaceFormText("IsTop")
	Dim IsHot:IsHot = F.ReplaceFormText("IsHots")
	if IsNew="" then IsNew = 0
	if IsCommand="" then IsCommand = 0
	if IsTop="" then IsTop = 0
	if IsHot="" then IsHot = 0
	Dim LastUpdateTime:LastUpdateTime = F.ReplaceFormText("txt_lastUpdatetime")
	DIm LookNeedScore:LookNeedScore = F.ReplaceFormText("txt_LookNeedScore")
	if LookNeedScore="" then LookNeedScore = 0
	'''
	Dim SubModelID:SubModelID = F.ReplaceFormText("txt_SubModelID")
	Dim SubModelItemID:SubModelItemID = F.ReplaceFormText("txt_SubModelItemID")
	if SubModelID="" then SubModelID = 0
	if SubModelItemID="" then SubModelItemID = 0
	'response.Write IsHots
	'''''
	Dim isusergroup:isusergroup = 0
	Dim ModelItemUserGroup:ModelItemUserGroup = F.ReplaceFormText("txt_usergroup")
	if ModelItemUserGroup<>"" then isusergroup = 1
	''''
	if action<>"" and action="add" then
		'''增加信息
		'自定义字段
		'自定义字段组合sql
		'自定义字段组合sql
		dim fsql:fsql="ModelID,SubModelID,SubModelItemID,LookNeedScore,CityID,IndexID,categoryID,IsJumpUrl,InfoType,Title,title_en,JumpUrl,ImgUrl,UserUID,UserName,IsNew,IsHot,IsTop,IsCommand,isusergroup,usergroup,LastUpdateTime"
		dim vsql:vsql=""&ModelID&","&SubModelID&","&SubModelItemID&","&LookNeedScore&","&CityID&","&OrderIndex&","&categoryID&","&IsJumpUrl&","&InfoType&",'"&Title&"','"&title_en&"','"&JumpUrl&"','"&ImgUrl&"',"&UserUID&",'"&UserName&"',"&IsNew&","&IsHot&","&IsTop&","&IsCommand&","&isusergroup&",'"&ModelItemUserGroup&"','"&LastUpdateTime&"'"
		dim tempvalue:tempvalue =""
		dim tempvalue_en:tempvalue_en =""
		for i=0 to ubound(Modelarray,2)
		   if Modelarray(6,i)<>"" then
			if fsql<>"" then 
				fsql = fsql &","&Modelarray(6,i)
				if config_Verison =2 then fsql = fsql &","&Modelarray(6,i)&"_en"
			else
				fsql = Modelarray(6,i)
				if config_Verison =2 then fsql = Modelarray(6,i)&"_en"
			end if
			if Modelarray(9,i)="int" then
				tempvalue = replace(F.ReplaceFormText("txt_"&Modelarray(6,i)),"'","")
				if config_Verison =2 then 
					tempvalue_en = replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_en"),"'","")
				else
					tempvalue_en = ""
				end if
			elseif Modelarray(7,i)="textaddSelect" then
				tempvalue = "'"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)),"'","")&"#"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_add"),"'","")&"'"
				if config_Verison =2 then 
					tempvalue_en ="'"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_en"),"'","")&"#"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_en_add"),"'","")&"'"
				else
					tempvalue_en = ""
				end if
			else
				tempvalue = "'"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)),"'","")&"'"
				if config_Verison =2 then 
					tempvalue_en = "'"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_en"),"'","")&"'"
				else
					tempvalue_en = ""
				end if
			end if
			if vsql<>"" then 
				vsql = vsql &","&tempvalue
				if config_Verison =2 then vsql = vsql &","&tempvalue_en
			else
				vsql = tempvalue
				if config_Verison =2 then vsql = tempvalue_en
			end if
		   end if
		next	 		 
		'自定义字段组合sql
		'自定义字段组合sql
		''''' 
		dim sql:sql="insert into "&ModelTableName&" ("&fsql&") values ("&vsql&")"
		'response.Write sql
		'response.End()
		Con.execute(sql)		
		Alert "增加信息成功!","Model_List.asp?ModelID="&ModelID&"&ParentModelID="&SubModelID&"&ModelItemID="&SubModelItemID&""
	elseif action="edit" then
		''修改信息
		'自定义字段组合sql
		'自定义字段组合sql
		'UserUID="&UserUID&",UserName='"&UserName&"'
		dim tempsql:tempsql="update "&ModelTableName&" set ModelID="&ModelID&",SubModelID="&SubModelID&",SubModelItemID="&SubModelItemID&",LookNeedScore="&LookNeedScore&",CityID="&CityID&",IndexID="&OrderIndex&",categoryID="&categoryID&",IsJumpUrl="&IsJumpUrl&",InfoType="&InfoType&",Title='"&Title&"',title_en='"&title_en&"',JumpUrl='"&JumpUrl&"',ImgUrl='"&ImgUrl&"',IsNew="&IsNew&",IsHot="&IsHot&",IsTop="&IsTop&",IsCommand="&IsCommand&",isusergroup="&isusergroup&",usergroup='"&ModelItemUserGroup&"',LastUpdateTime='"&LastUpdateTime&"'"
		dim sqlvalue:sqlvalue =""		
		dim sqlvalue_en:sqlvalue_en =""
		for i=0 to ubound(Modelarray,2)
		  if Modelarray(6,i)<>"" then
			if Modelarray(9,i)="int" then
				sqlvalue = replace(F.ReplaceFormText("txt_"&Modelarray(6,i)),"'","")
				if config_Verison=2 then 
					sqlvalue_en = replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_en"),"'","")
				else
					sqlvalue_en = ""
				end if
			elseif Modelarray(7,i)="textaddSelect" then
				sqlvalue = "'"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)),"'","")&"#"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_add"),"'","")&"'"
				if config_Verison=2 then 
					sqlvalue_en = "'"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_en"),"'","")&"#"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_en_add"),"'","")&"'"
				else
					sqlvalue_en = ""
				end if
			else
				sqlvalue = "'"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)),"'","")&"'"
				if config_Verison=2 then 
					sqlvalue_en = "'"&replace(F.ReplaceFormText("txt_"&Modelarray(6,i)&"_en"),"'","")&"'"
				else
					sqlvalue_en = "''"
				end if
			end if

		 	if tempsql<>"" then 
				tempsql = tempsql &","&Modelarray(6,i)&"="&sqlvalue
			else
				tempsql = Modelarray(6,i)&"="&sqlvalue
			end if
			if config_Verison=2 then
				tempsql = tempsql &","&Modelarray(6,i)&"_en"&"="&sqlvalue_en
			end if
		   end if
		 next 
		 '''加入条件
		 tempsql = tempsql &" where ID="&ItemID&""
		'response.Write tempsql
		'response.End()
		'自定义字段组合sql
		'自定义字段组合sql
		Con.execute(tempsql)
		Alert "修改信息成功!","Model_List.asp?ModelID="&ModelID&"&ParentModelID="&SubModelID&"&ModelItemID="&SubModelItemID&""
	end if
%>