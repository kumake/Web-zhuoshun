<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'函数名称：模型数据(列表,信息详细,推荐)					'	
'作    者：一支笔(QQ:317340145)							'
'开发日期：2008年8月12日								'
'备    注：原创代码免费授权使用。请保留该版权声明		'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class SpCMS_Tempalte
	Dim Reg
	Dim total_record   	'总记录数
	Dim current_page 	'当前页数
	Dim PCount			'循环页显示数目
	Dim page_pageSize	'每页显示记录数
	
	'''''初始化类
	Private Sub Class_Initialize()    
		Set Reg = New RegExp    
		Reg.Ignorecase = True   
		Reg.Global = True
		'全局变量
		current_page = 1 
		PCount=6	
		page_pageSize = 10 
	End Sub  
	 
	'类释放       
	Private Sub Class_Terminate()       
	End Sub  	
	'声明属性
	'Public Property Get ID       
	'End Property
	'Public Property Let ID(intId)       
	'End Property  
	'取URL路径
	public Function GetUrl(total) 
		'response.Write total
		Dim ScriptAddress, M_ItemUrl, M_item
		ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME"))&"?"'取得当前地址 
		If (Request.QueryString <> "") Then 
			M_ItemUrl = "" 
		For Each M_item In Request.QueryString 
			If InStr("page",M_Item)=0 and InStr("total",M_Item)=0 Then 
				M_ItemUrl = M_ItemUrl & M_Item &"="& Server.URLEncode(Request.QueryString(""&M_Item&"")) & "&" 
			End If 
		Next 
		ScriptAddress = ScriptAddress & M_ItemUrl'取得带参数地址 
		End If 
		GetUrl = ScriptAddress & "total="&total&"&page=" 
	End Function 
	
	public Function Parser_IF(Byval Content)
	  On Error Resume Next   
	  Dim TestIF    
	  Reg.Pattern = "{If:(.+?)}([\s\S]*?){Else}([\s\S]*?){End}"
	  Set Matches = Reg.Execute(Content)    
	  For Each Match In Matches    
	   Execute("If" & Match.SubMatches(0) & " Then TestIf = True Else TestIf = False")  
	   If TestIf Then 
	   Content = Replace(Content, Match.Value, Match.SubMatches(1)) 
	   Else 
	   Content = Replace(Content, Match.Value, Match.SubMatches(2)) '# 替换   
	   end if 
	   If Err Then 
	   Err.Clear
	   Response.Write "<font color=red>执行IF标签失败[" & Match.SubMatches(0) & "]</font>" 
	   Response.End 
	   end if  
	  Next
	  Parser_IF = Content
	End Function   
		
	''替换标签函数
	public Function Tags_Value(Byval Pattern,Byval Temp,Byval Dat,Byval TitleLen)    
		  Dim Matche    
		  Dim Tagsstr,Tagsval
		  Reg.Pattern = Pattern    
		  Set Matches = Reg.Execute(Temp)    
		  For Each Match In Matches    
		   Tagsstr = Match.SubMatches(0)   
		   Tagsval= Split(Tagsstr," ")(0)  
		   Tagsval= Dat(Tagsval)    
		   if isnumeric(TitleLen) and TitleLen>0 then  Tagsval = left(Tagsval,TitleLen)
		   Tagsval= Replace(Replace(Replace(Replace(Tagsval," "," "),"""",chr(34)),">",">"),"<","<")
		   Temp = Replace(Temp, Match.Value, Tagsval)    
		  next
		  Tags_Value = Temp
	end function 
	
	'singlepage 单页系统显示可以为ID也可以为字符
	'distype	显示类型 Instr:显示简介 content:显示内容  picname:显示图片
	public function SinglepageView(Byval singlepage,Byval distype)
		Dim BackValue:BackValue = ""
		Dim sql:sql = "select U.Instr,u.content,U.picname from user_Singlepage U,Sp_SinglePage S where U.singlepageID = S.ID"
		if isnumeric(singlepage) then
		sql = sql &" and S.id="&singlepage&""
		else
		sql = sql &" and S.singlePagename='"&singlepage&"'"
		end if
		'''执行sql取出数据
		set rs = UICon.Query(sql)
		if rs.recordcount<>0 then
			'''循环取得信息
			BackValue = rs(distype)
		rs.close
		set rs=nothing
		end if			
		''''打印
		SinglepageView = BackValue	
	end function
	
	'ModelID 	模型ID
	'ItemID		信息标题
	'DisplyStr	显示Html
	public function InfoView(Byval ModelID,Byval ItemID,Byval DisplyStr)
		'''取表名称
		Dim Modelarray:Modelarray = UICon.QueryRow("select modeltable from Sp_Model where Id="&ModelID&"",0) 
		Dim ModelTableName:ModelTableName = Modelarray(0)		
		'''查询sql
		if ItemID="" then ItemID = VerificationUrlParam("ItemID","int","0")
		Dim sql:sql = "select * from "&ModelTableName&" where isdel=0 and Id="&ItemID&""
		'''执行sql取出数据
		set rs = UICon.Query(sql)
		if rs.recordcount<>0 then
			'''循环取得信息
			BackValue = Tags_Value("\[field:(.+?)\]",DisplyStr,rs,0)
		rs.close
		set rs=nothing
		end if			
		''''打印
		InfoView = BackValue	
	end function
	
	'ModelID 	模型ID
	'CategoryID 模型类别 默认为0
	'GetField	需要的字段 默认为*全部多个字段用,隔开
	'OrderBy	排序条件 默认为 ID desc 在调用时候可为空
	'searchField	和下列
	'KeyWord	只显示和该关键词相关的信息 可以为空
	'TopNum 	置顶数量
	'TitleLen	标题字数 默认为0
	'LoopStr	循环元素
	public function CommandTopInfo(Byval ModelID,Byval CategoryID,Byval GetField,Byval OrderBy,Byval searchField,Byval KeyWord,Byval TopNum,Byval TitleLen,Byval LoopStr)
		'''定义变量
		Dim BackValue
		'''取表名称
		Dim Modelarray:Modelarray = UICon.QueryRow("select modeltable from Sp_Model where Id="&ModelID&"",0) 
		Dim ModelTableName:ModelTableName = Modelarray(0)		
		'''组合查询sql
		Dim sql:sql = "select top "&TopNum&" "&GetField&" from "&ModelTableName&" where isdel=0 and ModelID="&ModelID&""
		Dim strWhere
		if CategoryID<>0 then strWhere = strWhere &" and CategoryID="&CategoryID&""
		if searchField<>"" and keyword<>"" then strWhere = strWhere &"  and "&searchField&" like '%"&keyword&"%'"
		'''排列条件	
		if OrderBy<>"" then strWhere = strWhere &" order by "&OrderBy&""
		'''执行组合后的sql取出数据
		'''response.Write sql
		set rs = UICon.Query(sql)
		Dim Rstemp:Rstemp = 1
		if rs.recordcount<>0 then
		do while not rs.eof and Rstemp<=TopNum
			'''循环取得信息
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rs,TitleLen)
		rs.movenext
		Rstemp = Rstemp + 1
		loop
		rs.close
		set rs=nothing
		end if			
		''''打印
		CommandTopInfo = BackValue	
	end function
	
	
	'ModelID 	模型ID
	'CategoryID 模型类别 默认为0
	'pagesize	每页显示记录数
	'GetField	需要的字段 默认为*全部多个字段用,隔开
	'OrderBy	排序方式 0：ByNotIn 1：ByNum
	'searchField	和下列
	'KeyWord	只显示和该关键词相关的信息 可以为空
	'TitleLen	标题字数 默认为0
	'LoopStr	循环元素
	public function InfoList(Byval ModelID,Byval CategoryID,Byval pagesize,Byval GetField,Byval OrderBy,Byval searchField,Byval KeyWord,Byval TitleLen,Byval LoopStr)
		''''''采用该分页类型的好处：
		''常规分页方式采用一次性将数据读入内存通过游标显示每页记录
		''当数据表记录大于10万甚至1000万时？
		''采用一支笔该分页方式的好处就是需要多少条记录就从表中读多少条记录
		''根据测试数据记录越多效果越明显。测试：asp+sql2000+数据量500万
		page_pageSize = pagesize
		'''''''
		'''取表名称
		Dim Modelarray:Modelarray = UICon.QueryRow("select modeltable from Sp_Model where Id="&ModelID&"",0) 
		Dim ModelTableName:ModelTableName = Modelarray(0)
		'''
		Dim strwhere:strwhere = "isdel=0"		'查询条件
		'''获取查询条件
		if CategoryID=0 then CategoryID = VerificationUrlParam("CategoryID","int","0")
		if CategoryID<>0 then strwhere = strwhere & " and categoryID="&CategoryID&""
		if searchField<>"" and keyword<>"" then strWhere = strWhere &"  and "&searchField&" like '%"&keyword&"%'"
		''''总数记录只取一次节省数据库压力
		Dim total:total = VerificationUrlParam("total","int","0")
		if total = 0 then 
			Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from "&ModelTableName&" where "&strwhere&"")
			total_record = Tarray(0,0)
		else
			total_record = total
		end if
		'''''
		current_page = VerificationUrlParam("page","int","1")
		'这种方式为根据ID为关键字排序
		'该中分页方式效果最好。建议使用,但是排序效果受到限制
		Dim sql
		if isnumeric(OrderBy) and OrderBy=1 then
		sql = UICon.QueryPageByNum(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", true,current_page,page_pageSize)
		else
		'这种方式为根据IndexID排序，IndexID值越大越靠前
		sql = UICon.QueryPageByNotIn(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", "indexID desc,ID",false,current_page,page_pageSize)
		end if
		'response.Write sql
		'response.End()
		set rsn = UIcon.Query(sql)
		if rsn.recordcount<>0 then
		dim Rstemp:Rstemp = 1
		do while not rsn.eof and Rstemp<=page_pageSize
			'''循环取得信息
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rsn,TitleLen)
			'''
		rsn.movenext
		Rstemp = Rstemp + 1
		loop
		rsn.close
		set rsn=nothing
		end if			
		''''打印
		InfoList = BackValue	
	end function
	
	'ModelID 	模型ID
	'ItemID		模型信息ID
	'pagesize	每页显示记录数
	'GetField	需要的字段 默认为*全部多个字段用,隔开
	'OrderBy	排序方式 0：升 1：降
	'LoopStr	循环元素
	public function InfoCommentList(Byval ModelID,Byval ItemID,Byval pagesize,Byval GetField,Byval OrderBy,Byval LoopStr)
		page_pageSize = pagesize
		'''''''
		'''表名称
		Dim ModelTableName:ModelTableName = "Sp_Comment"
		'''
		Dim strwhere:strwhere = "isdel=0"		'查询条件
		'''获取查询条件
		if ModelID<>0 then strwhere = strwhere & " and ModelID="&ModelID&""
		if ItemID<>0 then strwhere = strwhere & " and ModelItemID="&ItemID&""
		''''总数记录只取一次节省数据库压力
		Dim total:total = VerificationUrlParam("total","int","0")
		if total = 0 then 
			Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from "&ModelTableName&" where "&strwhere&"")
			total_record = Tarray(0,0)
		else
			total_record = total
		end if
		'''''
		current_page = VerificationUrlParam("page","int","1")
		'这种方式为根据ID为关键字排序
		'该中分页方式效果最好。建议使用,但是排序效果受到限制
		Dim sql
		Dim orderFlag:orderFlag = false
		sql = UICon.QueryPageByNum(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", orderFlag,current_page,page_pageSize)
		'response.Write sql
		'response.End()
		set rsn = UIcon.Query(sql)
		if rsn.recordcount<>0 then
		dim Rstemp:Rstemp = 1
		do while not rsn.eof and Rstemp<=page_pageSize
			'''循环取得信息
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rsn,0)
			'''
		rsn.movenext
		Rstemp = Rstemp + 1
		loop
		rsn.close
		set rsn=nothing
		end if			
		''''打印
		InfoCommentList = BackValue	
	end function
	
	
	'TableName 	表名称
	'pagesize	每页显示记录数
	'GetField	需要的字段 默认为*全部多个字段用,隔开
	'OrderBy	排序方式 0：ByNotIn 1：ByNum
	'searchField	和下列
	'KeyWord	只显示和该关键词相关的信息 可以为空
	'TitleLen	标题字数 默认为0
	'LoopStr	循环元素
	public function IrregularInfoList(Byval tableName,Byval pagesize,Byval GetField,Byval OrderBy,Byval searchField,Byval KeyWord,Byval TitleLen,Byval LoopStr)
		''''''采用该分页类型的好处：
		''常规分页方式采用一次性将数据读入内存通过游标显示每页记录
		''当数据表记录大于10万甚至1000万时？
		''采用一支笔该分页方式的好处就是需要多少条记录就从表中读多少条记录
		''根据测试数据记录越多效果越明显。测试：asp+sql2000+数据量500万
		page_pageSize = pagesize
		'''''''
		Dim ModelTableName:ModelTableName = tableName
		Dim strwhere:strwhere = " 1=1"		'查询条件
		'''获取查询条件
		if searchField<>"" and keyword<>"" then strWhere = strWhere &"  and "&searchField&" like '%"&keyword&"%'"
		'if CategoryID=0 then CategoryID = VerificationUrlParam("CategoryID","int","0")
		'if CategoryID<>0 then strwhere = strwhere & " and categoryID="&CategoryID&""
		''''总数记录只取一次节省数据库压力
		Dim total:total = VerificationUrlParam("total","int","0")
		if total = 0 then 
			Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from "&ModelTableName&" where "&strwhere&"")
			total_record = Tarray(0,0)
		else
			total_record = total
		end if
		'''''
		current_page = VerificationUrlParam("page","int","1")
		'这种方式为根据ID为关键字排序
		'该中分页方式效果最好。建议使用,但是排序效果受到限制
		Dim sql
		'if isnumeric(OrderBy) and OrderBy=1 then
		sql = UICon.QueryPageByNum(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", true,current_page,page_pageSize)
		'else
		'这种方式为根据IndexID排序，IndexID值越大越靠前
		'sql = UICon.QueryPageByNotIn(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", "indexID desc,ID",false,current_page,page_pageSize)
		'end if
		'response.Write sql
		'response.End()
		set rsn = UIcon.Query(sql)
		if rsn.recordcount<>0 then
		dim Rstemp:Rstemp = 1
		do while not rsn.eof and Rstemp<=page_pageSize
			'''循环取得信息
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rsn,TitleLen)
			'''
		rsn.movenext
		Rstemp = Rstemp + 1
		loop
		rsn.close
		set rsn=nothing
		end if			
		''''打印
		IrregularInfoList = BackValue	
	end function
	
	
	'DisplyType  	1:logo 0:test
	'DisplyNum    	显示数量
	'Loopstr		循环条件
	public function FriendLink(Byval DisplyType,Byval DisplyNum,Byval Loopstr)
		'''定义变量
		Dim BackValue
		'''组合查询sql
		Dim sql:sql = "select top "&DisplyNum&" urltype,sitename,siteurl,sitelogourl from Sp_FriendLink where ISdisplay=1"
		if DisplyType<>"" and isnumeric(DisplyType) then sql = sql &" and urltype="&DisplyType&""
		'''执行sql取出数据
		set rs = UICon.Query(sql)
		Dim Rstemp:Rstemp = 1
		if rs.recordcount<>0 then
		do while not rs.eof and Rstemp<=DisplyNum
			'''循环取得信息
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rs,TitleLen)
		rs.movenext
		Rstemp = Rstemp + 1
		loop
		rs.close
		set rs=nothing
		end if			
		''''打印
		BackValue = Parser_IF(BackValue)
		FriendLink = BackValue		
	end function
	
	'dictionaryID		显示类别树的ID
	'Depth				显示几级深度
	'Separate			受程序限制最好是图片
	'Loopstr			循环条件
	public function ListCategoryTree(Byval dictionaryID,Byval DisplyDepth,Byval Separate,Byval Loopstr)
		''
		Dim sql:sql = "select ID,categoryID,categoryname,Pathstr,ParentID,Depth from Sp_dictionary where categoryID = "&dictionaryID&" and Depth<="&DisplyDepth&" order by Pathstr asc"
		set rs = UICon.Query(sql)
		Dim spaceLen:spaceLen = 0
		''''
		if rs.recordcount<>0 then
		do while not rs.eof
		spaceLen = Cint(rs("Depth"))-1
			'''循环取得信息
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rs,0)
			'''替换类别前的符号
			BackValue = replace(BackValue,"[Separate]",string(spaceLen,Separate))
		rs.movenext
		loop
		rs.close
		set rs=nothing
		end if			
		''''打印
		ListCategoryTree = BackValue		
	end function
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''分页函数													'
	''total_record  总记录数									'
	''current_page  当前页										'
	''PCount		   页码排列									'
	''pagesize	   每页记录数									'
	''showpageNum   是否显示循环页码							'
	''showpagetotal	是否显示总记录								'
	''IsEnglish		是否显示英文分页							'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	public function Infopage(Byval showpageNum,Byval showpagetotal,Byval IsEnglish)
		'''''''''''定义文字表达'''''''''''''''
		Dim Total_Text
		Dim First_Text 
		Dim Next_Text
		Dim Pre_Text
		Dim Last_Text
		Dim Record_Unit
		Dim Page_Unit
		
		if IsEnglish then:Total_Text = "Total Record:":else:Total_Text = "总记录:":end if
		if IsEnglish then:First_Text= "Frist":else:First_Text= "首页":end if
		if IsEnglish then:Next_Text= "next":else:Next_Text= "下一页":end if
		if IsEnglish then:Pre_Text= "pre":else:Pre_Text= "上一页":end if
		if IsEnglish then:Last_Text= "Last":else:Last_Text= "末页":end if
		if IsEnglish then:Record_Unit= "record":else:Record_Unit= "条":end if
		if IsEnglish then:Page_Unit= "page":else:Page_Unit= "页":end if
		''''''''''''''''''''''''''
		Dim PageIndex,pagestart,PageEnd
		Dim total_page:total_page = total_record \ page_pageSize   '总页数
		if total_record mod page_pageSize <>0 then
		  total_page = total_page + 1
		end if
	%>
	<div id="Divpage">
	<%
	if showpagetotal then
	%>
		<li><%=Total_Text%><%=total_record%><%=Record_Unit%></li>
		<li><%=current_page%>/<%=total_page%><%=Page_Unit%></li>
	<%
	end if
	%>
	<li><%if current_page>1 then:Response.Write "<a href='"& GetUrl(total_record) &"1'>"&First_Text&"</a>":else:Response.Write ""&First_Text&"":end if%></li>
	<li><%if current_page<=1 then:response.Write ""&Pre_Text&"":else:Response.Write "<a href='"& GetUrl(total_record) & (current_page-1) &"'>"&Pre_Text&"</a>":end if%></li>
	<%
	if showpageNum then
		Dim Temp_X:Temp_X = (current_page \ PCount)  '取整数
		Dim Temp_Y:Temp_Y = (current_page mod PCount)  '取余数
		if Temp_X<1 then
			pagestart = 0
			PageEnd = PCount+1
		elseif Temp_X=1 and Temp_Y=0 then
			pagestart = 0
			PageEnd = PCount+1
		elseif Temp_X=0 and Temp_Y=1 then	
			pagestart = 1
			PageEnd = 2*PCount+1
		elseif Temp_X>1 and Temp_Y=0 then
			pagestart = (Temp_X)*PCount-1
			PageEnd = (Temp_X+1)*PCount
		else
			pagestart = (Temp_X)*PCount
			PageEnd = (Temp_X+1)*PCount
		end if
		if PageEnd>total_page then:PageEnd = total_page:end if
		if pagestart=0 then:pagestart = 1:end if
		for PageIndex=pagestart to PageEnd
			Response.Write "<li><a href="& GetUrl(total_record) & PageIndex &">"
			if Cint(current_page) = Cint(PageIndex) then
				Response.Write "<font color='red'><U><b>"& PageIndex &"</b></U></font>"
			else
				Response.Write PageIndex
			end if
			Response.Write "</a></li>"&chr(13)
		next
	end if
	%>
	<li><%if Cint(current_page) < Cint(total_page) then:Response.Write "<a href='"& GetUrl(total_record) & (current_page+1) &"'>"&Next_Text&"</a>":else:Response.Write ""&Next_Text&"":end if%></li>
	<li><%if Cint(current_page) < Cint(total_page) then:Response.Write "<a href='"& GetUrl(total_record) & total_page &"'>"&Last_Text&"</a>":else:Response.Write ""&Last_Text&"":end if%></li>
	</div>
<%
	End function
	'分页函数结束
End Class 
'''''''''''''''''''''''''''''''''''''
'以上为系统解析函数类				'
'以下为定义一个全局类				'
'''''''''''''''''''''''''''''''''''''
Set template = New SpCMS_Tempalte
%> 
