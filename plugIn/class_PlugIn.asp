<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�������ƣ�ģ������(�б�,��Ϣ��ϸ,�Ƽ�)					'	
'��    �ߣ�һ֧��(QQ:317340145)							'
'�������ڣ�2008��8��12��								'
'��    ע��ԭ�����������Ȩʹ�á��뱣���ð�Ȩ����		'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class SpCMS_Tempalte
	Dim Reg
	Dim total_record   	'�ܼ�¼��
	Dim current_page 	'��ǰҳ��
	Dim PCount			'ѭ��ҳ��ʾ��Ŀ
	Dim page_pageSize	'ÿҳ��ʾ��¼��
	
	'''''��ʼ����
	Private Sub Class_Initialize()    
		Set Reg = New RegExp    
		Reg.Ignorecase = True   
		Reg.Global = True
		'ȫ�ֱ���
		current_page = 1 
		PCount=6	
		page_pageSize = 10 
	End Sub  
	 
	'���ͷ�       
	Private Sub Class_Terminate()       
	End Sub  	
	'��������
	'Public Property Get ID       
	'End Property
	'Public Property Let ID(intId)       
	'End Property  
	'ȡURL·��
	public Function GetUrl(total) 
		'response.Write total
		Dim ScriptAddress, M_ItemUrl, M_item
		ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME"))&"?"'ȡ�õ�ǰ��ַ 
		If (Request.QueryString <> "") Then 
			M_ItemUrl = "" 
		For Each M_item In Request.QueryString 
			If InStr("page",M_Item)=0 and InStr("total",M_Item)=0 Then 
				M_ItemUrl = M_ItemUrl & M_Item &"="& Server.URLEncode(Request.QueryString(""&M_Item&"")) & "&" 
			End If 
		Next 
		ScriptAddress = ScriptAddress & M_ItemUrl'ȡ�ô�������ַ 
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
	   Content = Replace(Content, Match.Value, Match.SubMatches(2)) '# �滻   
	   end if 
	   If Err Then 
	   Err.Clear
	   Response.Write "<font color=red>ִ��IF��ǩʧ��[" & Match.SubMatches(0) & "]</font>" 
	   Response.End 
	   end if  
	  Next
	  Parser_IF = Content
	End Function   
		
	''�滻��ǩ����
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
	
	'singlepage ��ҳϵͳ��ʾ����ΪIDҲ����Ϊ�ַ�
	'distype	��ʾ���� Instr:��ʾ��� content:��ʾ����  picname:��ʾͼƬ
	public function SinglepageView(Byval singlepage,Byval distype)
		Dim BackValue:BackValue = ""
		Dim sql:sql = "select U.Instr,u.content,U.picname from user_Singlepage U,Sp_SinglePage S where U.singlepageID = S.ID"
		if isnumeric(singlepage) then
		sql = sql &" and S.id="&singlepage&""
		else
		sql = sql &" and S.singlePagename='"&singlepage&"'"
		end if
		'''ִ��sqlȡ������
		set rs = UICon.Query(sql)
		if rs.recordcount<>0 then
			'''ѭ��ȡ����Ϣ
			BackValue = rs(distype)
		rs.close
		set rs=nothing
		end if			
		''''��ӡ
		SinglepageView = BackValue	
	end function
	
	'ModelID 	ģ��ID
	'ItemID		��Ϣ����
	'DisplyStr	��ʾHtml
	public function InfoView(Byval ModelID,Byval ItemID,Byval DisplyStr)
		'''ȡ������
		Dim Modelarray:Modelarray = UICon.QueryRow("select modeltable from Sp_Model where Id="&ModelID&"",0) 
		Dim ModelTableName:ModelTableName = Modelarray(0)		
		'''��ѯsql
		if ItemID="" then ItemID = VerificationUrlParam("ItemID","int","0")
		Dim sql:sql = "select * from "&ModelTableName&" where isdel=0 and Id="&ItemID&""
		'''ִ��sqlȡ������
		set rs = UICon.Query(sql)
		if rs.recordcount<>0 then
			'''ѭ��ȡ����Ϣ
			BackValue = Tags_Value("\[field:(.+?)\]",DisplyStr,rs,0)
		rs.close
		set rs=nothing
		end if			
		''''��ӡ
		InfoView = BackValue	
	end function
	
	'ModelID 	ģ��ID
	'CategoryID ģ����� Ĭ��Ϊ0
	'GetField	��Ҫ���ֶ� Ĭ��Ϊ*ȫ������ֶ���,����
	'OrderBy	�������� Ĭ��Ϊ ID desc �ڵ���ʱ���Ϊ��
	'searchField	������
	'KeyWord	ֻ��ʾ�͸ùؼ�����ص���Ϣ ����Ϊ��
	'TopNum 	�ö�����
	'TitleLen	�������� Ĭ��Ϊ0
	'LoopStr	ѭ��Ԫ��
	public function CommandTopInfo(Byval ModelID,Byval CategoryID,Byval GetField,Byval OrderBy,Byval searchField,Byval KeyWord,Byval TopNum,Byval TitleLen,Byval LoopStr)
		'''�������
		Dim BackValue
		'''ȡ������
		Dim Modelarray:Modelarray = UICon.QueryRow("select modeltable from Sp_Model where Id="&ModelID&"",0) 
		Dim ModelTableName:ModelTableName = Modelarray(0)		
		'''��ϲ�ѯsql
		Dim sql:sql = "select top "&TopNum&" "&GetField&" from "&ModelTableName&" where isdel=0 and ModelID="&ModelID&""
		Dim strWhere
		if CategoryID<>0 then strWhere = strWhere &" and CategoryID="&CategoryID&""
		if searchField<>"" and keyword<>"" then strWhere = strWhere &"  and "&searchField&" like '%"&keyword&"%'"
		'''��������	
		if OrderBy<>"" then strWhere = strWhere &" order by "&OrderBy&""
		'''ִ����Ϻ��sqlȡ������
		'''response.Write sql
		set rs = UICon.Query(sql)
		Dim Rstemp:Rstemp = 1
		if rs.recordcount<>0 then
		do while not rs.eof and Rstemp<=TopNum
			'''ѭ��ȡ����Ϣ
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rs,TitleLen)
		rs.movenext
		Rstemp = Rstemp + 1
		loop
		rs.close
		set rs=nothing
		end if			
		''''��ӡ
		CommandTopInfo = BackValue	
	end function
	
	
	'ModelID 	ģ��ID
	'CategoryID ģ����� Ĭ��Ϊ0
	'pagesize	ÿҳ��ʾ��¼��
	'GetField	��Ҫ���ֶ� Ĭ��Ϊ*ȫ������ֶ���,����
	'OrderBy	����ʽ 0��ByNotIn 1��ByNum
	'searchField	������
	'KeyWord	ֻ��ʾ�͸ùؼ�����ص���Ϣ ����Ϊ��
	'TitleLen	�������� Ĭ��Ϊ0
	'LoopStr	ѭ��Ԫ��
	public function InfoList(Byval ModelID,Byval CategoryID,Byval pagesize,Byval GetField,Byval OrderBy,Byval searchField,Byval KeyWord,Byval TitleLen,Byval LoopStr)
		''''''���ø÷�ҳ���͵ĺô���
		''�����ҳ��ʽ����һ���Խ����ݶ����ڴ�ͨ���α���ʾÿҳ��¼
		''�����ݱ��¼����10������1000��ʱ��
		''����һ֧�ʸ÷�ҳ��ʽ�ĺô�������Ҫ��������¼�ʹӱ��ж���������¼
		''���ݲ������ݼ�¼Խ��Ч��Խ���ԡ����ԣ�asp+sql2000+������500��
		page_pageSize = pagesize
		'''''''
		'''ȡ������
		Dim Modelarray:Modelarray = UICon.QueryRow("select modeltable from Sp_Model where Id="&ModelID&"",0) 
		Dim ModelTableName:ModelTableName = Modelarray(0)
		'''
		Dim strwhere:strwhere = "isdel=0"		'��ѯ����
		'''��ȡ��ѯ����
		if CategoryID=0 then CategoryID = VerificationUrlParam("CategoryID","int","0")
		if CategoryID<>0 then strwhere = strwhere & " and categoryID="&CategoryID&""
		if searchField<>"" and keyword<>"" then strWhere = strWhere &"  and "&searchField&" like '%"&keyword&"%'"
		''''������¼ֻȡһ�ν�ʡ���ݿ�ѹ��
		Dim total:total = VerificationUrlParam("total","int","0")
		if total = 0 then 
			Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from "&ModelTableName&" where "&strwhere&"")
			total_record = Tarray(0,0)
		else
			total_record = total
		end if
		'''''
		current_page = VerificationUrlParam("page","int","1")
		'���ַ�ʽΪ����IDΪ�ؼ�������
		'���з�ҳ��ʽЧ����á�����ʹ��,��������Ч���ܵ�����
		Dim sql
		if isnumeric(OrderBy) and OrderBy=1 then
		sql = UICon.QueryPageByNum(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", true,current_page,page_pageSize)
		else
		'���ַ�ʽΪ����IndexID����IndexIDֵԽ��Խ��ǰ
		sql = UICon.QueryPageByNotIn(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", "indexID desc,ID",false,current_page,page_pageSize)
		end if
		'response.Write sql
		'response.End()
		set rsn = UIcon.Query(sql)
		if rsn.recordcount<>0 then
		dim Rstemp:Rstemp = 1
		do while not rsn.eof and Rstemp<=page_pageSize
			'''ѭ��ȡ����Ϣ
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rsn,TitleLen)
			'''
		rsn.movenext
		Rstemp = Rstemp + 1
		loop
		rsn.close
		set rsn=nothing
		end if			
		''''��ӡ
		InfoList = BackValue	
	end function
	
	'ModelID 	ģ��ID
	'ItemID		ģ����ϢID
	'pagesize	ÿҳ��ʾ��¼��
	'GetField	��Ҫ���ֶ� Ĭ��Ϊ*ȫ������ֶ���,����
	'OrderBy	����ʽ 0���� 1����
	'LoopStr	ѭ��Ԫ��
	public function InfoCommentList(Byval ModelID,Byval ItemID,Byval pagesize,Byval GetField,Byval OrderBy,Byval LoopStr)
		page_pageSize = pagesize
		'''''''
		'''������
		Dim ModelTableName:ModelTableName = "Sp_Comment"
		'''
		Dim strwhere:strwhere = "isdel=0"		'��ѯ����
		'''��ȡ��ѯ����
		if ModelID<>0 then strwhere = strwhere & " and ModelID="&ModelID&""
		if ItemID<>0 then strwhere = strwhere & " and ModelItemID="&ItemID&""
		''''������¼ֻȡһ�ν�ʡ���ݿ�ѹ��
		Dim total:total = VerificationUrlParam("total","int","0")
		if total = 0 then 
			Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from "&ModelTableName&" where "&strwhere&"")
			total_record = Tarray(0,0)
		else
			total_record = total
		end if
		'''''
		current_page = VerificationUrlParam("page","int","1")
		'���ַ�ʽΪ����IDΪ�ؼ�������
		'���з�ҳ��ʽЧ����á�����ʹ��,��������Ч���ܵ�����
		Dim sql
		Dim orderFlag:orderFlag = false
		sql = UICon.QueryPageByNum(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", orderFlag,current_page,page_pageSize)
		'response.Write sql
		'response.End()
		set rsn = UIcon.Query(sql)
		if rsn.recordcount<>0 then
		dim Rstemp:Rstemp = 1
		do while not rsn.eof and Rstemp<=page_pageSize
			'''ѭ��ȡ����Ϣ
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rsn,0)
			'''
		rsn.movenext
		Rstemp = Rstemp + 1
		loop
		rsn.close
		set rsn=nothing
		end if			
		''''��ӡ
		InfoCommentList = BackValue	
	end function
	
	
	'TableName 	������
	'pagesize	ÿҳ��ʾ��¼��
	'GetField	��Ҫ���ֶ� Ĭ��Ϊ*ȫ������ֶ���,����
	'OrderBy	����ʽ 0��ByNotIn 1��ByNum
	'searchField	������
	'KeyWord	ֻ��ʾ�͸ùؼ�����ص���Ϣ ����Ϊ��
	'TitleLen	�������� Ĭ��Ϊ0
	'LoopStr	ѭ��Ԫ��
	public function IrregularInfoList(Byval tableName,Byval pagesize,Byval GetField,Byval OrderBy,Byval searchField,Byval KeyWord,Byval TitleLen,Byval LoopStr)
		''''''���ø÷�ҳ���͵ĺô���
		''�����ҳ��ʽ����һ���Խ����ݶ����ڴ�ͨ���α���ʾÿҳ��¼
		''�����ݱ��¼����10������1000��ʱ��
		''����һ֧�ʸ÷�ҳ��ʽ�ĺô�������Ҫ��������¼�ʹӱ��ж���������¼
		''���ݲ������ݼ�¼Խ��Ч��Խ���ԡ����ԣ�asp+sql2000+������500��
		page_pageSize = pagesize
		'''''''
		Dim ModelTableName:ModelTableName = tableName
		Dim strwhere:strwhere = " 1=1"		'��ѯ����
		'''��ȡ��ѯ����
		if searchField<>"" and keyword<>"" then strWhere = strWhere &"  and "&searchField&" like '%"&keyword&"%'"
		'if CategoryID=0 then CategoryID = VerificationUrlParam("CategoryID","int","0")
		'if CategoryID<>0 then strwhere = strwhere & " and categoryID="&CategoryID&""
		''''������¼ֻȡһ�ν�ʡ���ݿ�ѹ��
		Dim total:total = VerificationUrlParam("total","int","0")
		if total = 0 then 
			Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from "&ModelTableName&" where "&strwhere&"")
			total_record = Tarray(0,0)
		else
			total_record = total
		end if
		'''''
		current_page = VerificationUrlParam("page","int","1")
		'���ַ�ʽΪ����IDΪ�ؼ�������
		'���з�ҳ��ʽЧ����á�����ʹ��,��������Ч���ܵ�����
		Dim sql
		'if isnumeric(OrderBy) and OrderBy=1 then
		sql = UICon.QueryPageByNum(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", true,current_page,page_pageSize)
		'else
		'���ַ�ʽΪ����IndexID����IndexIDֵԽ��Խ��ǰ
		'sql = UICon.QueryPageByNotIn(""&GetField&"",""&ModelTableName&"", ""&strwhere&"", "ID", "indexID desc,ID",false,current_page,page_pageSize)
		'end if
		'response.Write sql
		'response.End()
		set rsn = UIcon.Query(sql)
		if rsn.recordcount<>0 then
		dim Rstemp:Rstemp = 1
		do while not rsn.eof and Rstemp<=page_pageSize
			'''ѭ��ȡ����Ϣ
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rsn,TitleLen)
			'''
		rsn.movenext
		Rstemp = Rstemp + 1
		loop
		rsn.close
		set rsn=nothing
		end if			
		''''��ӡ
		IrregularInfoList = BackValue	
	end function
	
	
	'DisplyType  	1:logo 0:test
	'DisplyNum    	��ʾ����
	'Loopstr		ѭ������
	public function FriendLink(Byval DisplyType,Byval DisplyNum,Byval Loopstr)
		'''�������
		Dim BackValue
		'''��ϲ�ѯsql
		Dim sql:sql = "select top "&DisplyNum&" urltype,sitename,siteurl,sitelogourl from Sp_FriendLink where ISdisplay=1"
		if DisplyType<>"" and isnumeric(DisplyType) then sql = sql &" and urltype="&DisplyType&""
		'''ִ��sqlȡ������
		set rs = UICon.Query(sql)
		Dim Rstemp:Rstemp = 1
		if rs.recordcount<>0 then
		do while not rs.eof and Rstemp<=DisplyNum
			'''ѭ��ȡ����Ϣ
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rs,TitleLen)
		rs.movenext
		Rstemp = Rstemp + 1
		loop
		rs.close
		set rs=nothing
		end if			
		''''��ӡ
		BackValue = Parser_IF(BackValue)
		FriendLink = BackValue		
	end function
	
	'dictionaryID		��ʾ�������ID
	'Depth				��ʾ�������
	'Separate			�ܳ������������ͼƬ
	'Loopstr			ѭ������
	public function ListCategoryTree(Byval dictionaryID,Byval DisplyDepth,Byval Separate,Byval Loopstr)
		''
		Dim sql:sql = "select ID,categoryID,categoryname,Pathstr,ParentID,Depth from Sp_dictionary where categoryID = "&dictionaryID&" and Depth<="&DisplyDepth&" order by Pathstr asc"
		set rs = UICon.Query(sql)
		Dim spaceLen:spaceLen = 0
		''''
		if rs.recordcount<>0 then
		do while not rs.eof
		spaceLen = Cint(rs("Depth"))-1
			'''ѭ��ȡ����Ϣ
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rs,0)
			'''�滻���ǰ�ķ���
			BackValue = replace(BackValue,"[Separate]",string(spaceLen,Separate))
		rs.movenext
		loop
		rs.close
		set rs=nothing
		end if			
		''''��ӡ
		ListCategoryTree = BackValue		
	end function
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''��ҳ����													'
	''total_record  �ܼ�¼��									'
	''current_page  ��ǰҳ										'
	''PCount		   ҳ������									'
	''pagesize	   ÿҳ��¼��									'
	''showpageNum   �Ƿ���ʾѭ��ҳ��							'
	''showpagetotal	�Ƿ���ʾ�ܼ�¼								'
	''IsEnglish		�Ƿ���ʾӢ�ķ�ҳ							'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	public function Infopage(Byval showpageNum,Byval showpagetotal,Byval IsEnglish)
		'''''''''''�������ֱ��'''''''''''''''
		Dim Total_Text
		Dim First_Text 
		Dim Next_Text
		Dim Pre_Text
		Dim Last_Text
		Dim Record_Unit
		Dim Page_Unit
		
		if IsEnglish then:Total_Text = "Total Record:":else:Total_Text = "�ܼ�¼:":end if
		if IsEnglish then:First_Text= "Frist":else:First_Text= "��ҳ":end if
		if IsEnglish then:Next_Text= "next":else:Next_Text= "��һҳ":end if
		if IsEnglish then:Pre_Text= "pre":else:Pre_Text= "��һҳ":end if
		if IsEnglish then:Last_Text= "Last":else:Last_Text= "ĩҳ":end if
		if IsEnglish then:Record_Unit= "record":else:Record_Unit= "��":end if
		if IsEnglish then:Page_Unit= "page":else:Page_Unit= "ҳ":end if
		''''''''''''''''''''''''''
		Dim PageIndex,pagestart,PageEnd
		Dim total_page:total_page = total_record \ page_pageSize   '��ҳ��
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
		Dim Temp_X:Temp_X = (current_page \ PCount)  'ȡ����
		Dim Temp_Y:Temp_Y = (current_page mod PCount)  'ȡ����
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
	'��ҳ��������
End Class 
'''''''''''''''''''''''''''''''''''''
'����Ϊϵͳ����������				'
'����Ϊ����һ��ȫ����				'
'''''''''''''''''''''''''''''''''''''
Set template = New SpCMS_Tempalte
%> 
