<%
''''''''''''''''''''''''''''''''''''''''''''''''''''
''asp模板引擎
''Author:陈海飞
''Datetime：2008-6-30 
''此代码为作者原创请保留版权信息
''如果恰当改变正则表达式，执行速度会更佳！
''''''''''''''''''''''''''''''''''''''''''''''''''''
Class Cls_Tempalte
   
 Dim Reg    
 Dim Code    
 Dim Page    
 Dim Rule    
 Dim Content    
 Dim Template    
 Dim Cachetime    
 Dim DefCachetime
 Dim ModelID	'模型ID
 Dim ItemID		'当前信息ID
''''
 Private Sub Class_Initialize()    
  Set Reg = New RegExp    
  Reg.Ignorecase = True   
  Reg.Global = True   
  Code = "GB2312"   
  Page = 0    
  Rule = ""   
  Content = ""   
  Template = ""   
  Cachetime = 0    
  DefCachetime = Cachetime 
  ModelID=0
  ModelID = request.QueryString("ModelID")
  if ModelID="" or isnumeric(ModelID)=false then ModelID=1
  ItemID=0
  ItemID = request.QueryString("ItemID")
  if ItemID="" or isnumeric(ItemID)=false then ItemID=1
 End Sub   
     
 Private Sub Class_Terminate()    
  Set Reg = Nothing   
 End Sub   
   
 Public Function Parser(Byval Templatefile)    
  Template = Templatefile    
  If ChkCache(Template) Then Parser = GetCache(Template) : Exit Function   
  If Not IsNumeric(Page) Then Page = 0 Else Page = Int(Page)    
  Loadfile    
  Parser_Sys	'系统变量
  Parser_my    	'自定义标签
  Parser_singlepage '单页系统
  Parser_Com    '文章列表不带分页
  'Parser_List '文章列表带分页
  Parser_display
  Parser_IF    '条件语句
  SetCache Template,Content   '缓存模板文件
  Parser = Content    '输出
 End Function   
     
 Public Function Parser_Sys()    
  'On Error Resume Next   
  Dim Matche,SysValue    
  Reg.Pattern = "{Sys:([\s\S]*?)}"   
  Set Matches = Reg.Execute(Content)    
  For Each Match In Matches
   If Len(Replace(Match.SubMatches(0)," ","")) > 0 Then 
   Execute("SysValue = " & Replace(Match.SubMatches(0)," ","")) 
   Else 
   SysValue = ""  
   end if
   Content = Replace(Content, Match.Value, SysValue) '# 替换    
   If Err Then Err.Clear : Response.Write "<font color=red>执行变量标签失败[" & AspArr(i) & "]</font>" : Response.End   
  Next   
 End Function     
     
 Public Function Parser_My()    
  'On Error Resume Next   
  Dim Rs,Ns,i    
  Set Rs = con.execute("Select [Name],[Code] From [sp_Label]")   
  If Rs.recordcount<>0 then
	  Ns = Rs.GetRows()
	  Dim Matche,MyValue
	  Reg.Pattern = "{My:([\s\S]*?)}"   
	  Set Matches = Reg.Execute(Content)   
	  For Each Match In Matches    
	   If Rs.recordcount<>0 Then   
		If Len(Replace(Match.SubMatches(0)," ","")) > 0 Then   	
			 MyValue = Replace(Match.SubMatches(0)," ","")    
			 For i = 0 To Ubound(Ns,2)    
				If Lcase(Ns(0,i)) = Lcase(Replace(Match.SubMatches(0)," ","")) Then MyValue = Ns(1,i) : Exit For   
			 Next   
		End If   
	   Else   
		MyValue = ""   
	   End If   
	   Content = Replace(Content, Match.Value, MyValue) '# 替换    
	   If Err Then Err.Clear : Response.Write "<font color=red>执行自定义标签失败[" & AspArr(i) & "]</font>" : Response.End   
	  Next   
	  Rs.Close    
	  Set Rs = Nothing  
   end if 
 End Function   
 
 
  Public Function Parser_singlepage()    
  'On Error Resume Next   
  Dim Rs,Ns,i    
  Set Rs = con.Query("select S.singlePagename,U.content from user_Singlepage U,Sp_SinglePage S where U.singlepageID = S.ID")   
  If Rs.recordcount<>0 Then 
	  Ns = Rs.GetRows()
	  Dim Matche,MyValue
	  Reg.Pattern = "{Spage:([\s\S]*?)}"   
	  Set Matches = Reg.Execute(Content)   
	  For Each Match In Matches    
		If Len(trim(Match.SubMatches(0))) > 0 Then   	
			 MyValue = trim(Match.SubMatches(0))    
			 For i = 0 To Ubound(Ns,2)    
				If Lcase(Ns(0,i)) = Lcase(trim(Match.SubMatches(0))) Then MyValue = Ns(1,i) : Exit For   
			 Next   
		End If  
	    if MyValue<>"" then 
		Content = Replace(Content, Match.Value, MyValue) '# 替换   
		else
		Content = Replace(Content, Match.Value, "<font color=red>执行单页标签"&Match.Value&"失败</font>") '# 替换   
		end if 
	    If Err Then Err.Clear : Response.Write "<font color=red>执行单页标签失败[" & AspArr(i) & "]</font>" : Response.End   
	  Next   
	  Rs.Close    
	  Set Rs = Nothing 
	end if  
 End Function   

  ''''''显示信息
 Public Function Parser_display()    
  'On Error Resume Next  
  Dim BackValue,loopStr   
  Reg.Pattern = "{sp:view}([\s\S]*?){/sp:view}" 
  Set Matches = Reg.Execute(Content)
  For Each Match In Matches  
  	  loopStr = Match.SubMatches(0)
  	  'response.Write Match.Value
	  Dim ModelTable:ModelTable = "user_news"
	  '''''根据ModelID取modelID
	  set rsb= con.Query("select modeltable from Sp_Model where ID="&ModelID&"")
	  if rsb.recordcount<>0 then
	  ModelTable = rsb("modeltable")
	  end if
	  rsb.close
	  set rsb = nothing
	  'response.Write ModelTable
	  ''''''
	  set rs = con.Query("select * from "&ModelTable&" where id="&itemid&"")
	  if rs.recordcount<>0 then
		 BackValue = BackValue & Tags_Value("\[field:(.+?)\]",loopStr,rs)
	  else
	   	 Response.Write "<font color=red>执行信息标签失败</font>"
	     Response.End   
	  end if
   Content = Replace(Content, Match.Value, BackValue)
   BackValue = ""
  next
 End Function  


Public Function Parser_IF()    
  'On Error Resume Next   
  Dim TestIF    
  Reg.Pattern = "{If:(.+?)}([\s\S]*?){Else}([\s\S]*?){End}"
  Set Matches = Reg.Execute(Content)    
  For Each Match In Matches    
   Execute("If" & Match.SubMatches(0) & " Then TestIf = True Else TestIf = False")  
   If TestIf Then Content = Replace(Content, Match.Value, Match.SubMatches(1)) Else Content = Replace(Content, Match.Value, Match.SubMatches(2)) '# 替换    
   If Err Then Err.Clear : Response.Write "<font color=red>执行IF标签失败[" & Match.SubMatches(0) & "]</font>" : Response.End   
  Next   
 End Function   
    
	 
 Public Function Parser_Com()    
  'On Error Resume Next   
  Dim Rs,i,j    
  Dim Matche,BackValue    
  Dim Tagsstr,Loopstr,tempLoop '取正则需要的标签
  ''''以下定义需要用到的标签''''''''''''''''''''''
  Dim Tag_modelTable		''模型表名称
  Dim Tag_CategoryID		''模型所属的分类ID
  Dim Tag_TopNum			''显示总行数
  Dim Tag_Row				''显示几行
  Dim Tag_Cols				''显示几列
  Dim Tag_IsCommand			''显示推荐
  Dim Tag_IsTop				''显示置顶
  Dim Tag_IsHot				''显示热点
  Dim Tag_KeyWord			''显示关键字相关
  Dim Tag_Sort				''那个字段排序
  Dim Tag_Order				''怎么排序
  Dim Tag_TitleNum			''标题显示多少个字符
  ''''以上定义需要用到的标签''''''''''''''''''''''''
  Dim Exec_sql '执行sql
  Reg.Pattern = "{sp:loop\s([\s\S]*?)}([\s\S]*?){/sp:loop}" 
  Set Matches = Reg.Execute(Content)
  For Each Match In Matches   
   Cachetime    = DefCachetime    
   Tagsstr      = Match.SubMatches(0)    
   Loopstr      = Match.SubMatches(1)
   '''取各个小参数
	  Tag_modelTable = GetLoopAttr(Tagsstr,"modelTable")
	  Tag_CategoryID = GetLoopAttr(Tagsstr,"CategoryID")
	  Tag_TopNum = GetLoopAttr(Tagsstr,"TopNum")
	  Tag_Row = GetLoopAttr(Tagsstr,"Row")
	  Tag_Cols = GetLoopAttr(Tagsstr,"Cols")
	  Tag_IsCommand = GetLoopAttr(Tagsstr,"IsCommand")
	  Tag_IsTop = GetLoopAttr(Tagsstr,"IsTop")
	  Tag_IsHot = GetLoopAttr(Tagsstr,"IsHot")
	  Tag_KeyWord = GetLoopAttr(Tagsstr,"KeyWord")
	  Tag_Sort = GetLoopAttr(Tagsstr,"Sort")
	  Tag_Order = GetLoopAttr(Tagsstr,"Order")
	  Tag_TitleNum = GetLoopAttr(Tagsstr,"TitleNum")
   	  'response.Write Tagsstr
	  'modelTable,CategoryID,TopNum,Row,Cols,IsCommand,IsTop,IsHot,KeyWord,Sort,Order,TitleNum
	  Exec_sql = "select top "&Tag_TopNum&" * from "&Tag_modelTable&" where 1=1"
	  if Tag_CategoryID<>"" then Exec_sql = Exec_sql &" and CategoryID="&Tag_CategoryID&""
   	  '''''执行
	  'response.Write Match.Value
	  set RS = con.Query(Exec_sql)
	  if rs.recordcount<>0 then
	  	do while not rs.eof
			''''替换标签
			BackValue = BackValue & Tags_Value("\[field:(.+?)\]",Loopstr,rs)
		rs.movenext
		loop
		rs.close()
		set rs = nothing
	  end if
   Exec_sql = ""   
   SetCache Tag_Sql,BackValue     
   Content = Replace(Content, Match.Value, BackValue)
   BackValue = ""
  Next   
 End Function
 
Public Function Tags_Value(Byval Pattern,Byval Temp,Byval Dat)    
	  Dim Matche    
	  Dim Tagsstr,Tagsval
	  Reg.Pattern = Pattern    
	  Set Matches = Reg.Execute(Temp)    
	  For Each Match In Matches    
	   Tagsstr = Match.SubMatches(0)   
	   Tagsval= Split(Tagsstr," ")(0)    
	   Tagsval= Dat(Tagsval)    
	   Tagsval= Replace(Replace(Replace(Replace(Tagsval," "," "),"""",chr(34)),">",">"),"<","<")
	   Temp = Replace(Temp, Match.Value, Tagsval)    
	  next
	  Tags_Value = Temp
end function 
 
 
 
 ''''''''''''
 ''取循环页中出现的逐一参数
 Public Function GetLoopAttr(Byval Tagsstr,Byval AttrName)    
  Tagsstr = Tagsstr & ","  
  Reg.Pattern = "sp:([\s\S]*?)=([\s\S]*?),"
  Set Matches = Reg.Execute(Tagsstr)    
  For Each Match In Matches  
  	if Lcase(trim(AttrName))=Lcase(trim(Match.SubMatches(0))) then
   	GetLoopAttr = Match.SubMatches(1)
	exit for
	else
	GetLoopAttr = ""
	end if
  Next   
 End Function   

 Public Function Parser_Tags(Byval Pattern,Byval Temp,Byval Dat)    
  Dim Matche    
  Dim Tagsstr,Tagsval,Tagsvalt    
  Dim Tag_Len,Tag_Lenext,Tag_Format,Tag_Replace,Tag_Clearhtml    
  Dim Re,Re1,Re2    
  Dim i,c,l    
  Reg.Pattern = Pattern    
  Set Matches = Reg.Execute(Temp)    
  For Each Match In Matches    
   Tagsstr        = Match.SubMatches(0)    
   Tag_Len        = GetAttr(Tagsstr,"len")    
   Tag_Lenext     = GetAttr(Tagsstr,"lenext")    
   Tag_Format     = GetAttr(Tagsstr,"format")    
   Tag_Replace    = GetAttr(Tagsstr,"replace")    
   Tag_Clearhtml  = GetAttr(Tagsstr,"clearhtml")    
       
   Tagsval        = Split(Tagsstr," ")(0)    
   Tagsval        = Dat(Tagsval)    
   Tagsval        = Replace(Replace(Replace(Replace(Tagsval," "," "),"""",chr(34)),">",">"),"<","<")
   If Len(Replace(Tag_Replace," ","")) > 0 Then   
    Re = Split(Tag_Replace,"##")    
    If Ubound(Re) >= 0 Then Re1 = Re(0) : Re2 = Re(1) Else Re1 = Re(0): Re2 = Re(0)    
    Tagsval = Replace(Tagsval,Re1,Re2)    
   End If   
   If Len(Replace(Tag_Format," ","")) > 0 Then '# 格式化时间    
    If IsDate(Tagsval) Then   
     Tagsvalt = Tagsval : Tagsvalt = Lcase(Tag_Format)    
     Tagsvalt = Replace(Tagsvalt,"yyyy",Year(Tagsval))                : Tagsvalt = Replace(Tagsvalt,"yy",Right(Year(Tagsval),2))    
     Tagsvalt = Replace(Tagsvalt,"mm",Right("0" & Month(Tagsval),2))  : Tagsvalt = Replace(Tagsvalt,"m",Month(Tagsval))    
     Tagsvalt = Replace(Tagsvalt,"dd",Right("0" & Day(Tagsval),2))    : Tagsvalt = Replace(Tagsvalt,"d",Day(Tagsval))    
     Tagsvalt = Replace(Tagsvalt,"hh",Right("0" & Hour(Tagsval),2))   : Tagsvalt = Replace(Tagsvalt,"h",Hour(Tagsval))    
     Tagsvalt = Replace(Tagsvalt,"nn",Right("0" & Minute(Tagsval),2)) : Tagsvalt = Replace(Tagsvalt,"n",Minute(Tagsval))    
     Tagsvalt = Replace(Tagsvalt,"ss",Right("0" & Second(Tagsval),2)) : Tagsvalt = Replace(Tagsvalt,"s",Second(Tagsval))    
     Tagsval  = Tagsvalt    
    End If   
   End If   
   If Lcase(Replace(Tag_Clearhtml," ","")) = "true" Then   
    Reg.Pattern = "(\<.+?\>)"   
    Tagsval = Reg.Replace(Tagsval, " ")    
    Tagsval = Trim(Tagsval)    
   End If   
   Tag_Len = Replace(Tag_Len," ","")    
   If Len(Tag_Len) > 0 And IsNumeric(Tag_Len) Then   
    For i = 1 To Len(Tagsval)    
     c = Abs(Asc(Mid(str,i,1)))     

     If c > 255  Or c < 1 Then t = t + 2 Else t = t + 1    
     If t >= Tag_Len then Tagsval = Left(Tagsval, i) & Tag_Lenext : Exit For   
    Next   
   End If   
   Temp = Replace(Temp, Match.Value, Tagsval)    
  Next   
  Parser_Tags = Temp    
 End Function 
     
 Public Function RegReplace(Byval Pattern,Byval ReplaceVal)    
  Reg.Pattern = Pattern    
  Set Matches = Reg.Execute(Content)    
  For Each Match In Matches    
   Content = Replace(Content, Match.Value, ReplaceVal)    
  Next   
 End Function   
     
 Public Function GetAttr(Byval Tagsstr,Byval AttrName)    
  'Reg.Pattern = "<!--(.+?):\{(.+?)\}-->([\s\S]*?)<!--\1-->" 
  Tagsstr = Tagsstr & " $"   
  Reg.Pattern = "\$" & AttrName & "=(.+?) \$"   
  Set Matches = Reg.Execute(Tagsstr)    
  For Each Match In Matches    
   GetAttr = Match.SubMatches(0)    
  Next   
 End Function   
     
 Function SetCache(Byval CacheName,Byval CacheValue)    
  CacheName = Filterstr(CacheName)    
  Application.Lock    
  Application("Template_" & CacheName) = CacheValue : Application("Template_" & CacheName & ".Time") = Now()    
  Application.UnLock    
 End Function   
     
 Function ChkCache(Byval CacheName)    
  Dim GetValue,GetTime    
  GetValue = GetCache(CacheName) : GetTime  = GetCache(CacheName & ".Time")    
  If IsNull(GetValue) Or IsEmpty(GetValue) Then ChkCache = False : Exit Function   
  If Not IsDate(GetTime) Then ChkCache = False : Exit Function   
  If DateDiff("s",CDate(GetTime),Now()) >= Cachetime Then ChkCache = False Else ChkCache = True   
 End Function   
     
 Function GetCache(Byval CacheName)    
  CacheName = Filterstr(CacheName) : GetCache = Application("Template_" & CacheName)    
 End Function   
     
 Function Filterstr(Byval Str)    
  Filterstr = LCase(Str) : Filterstr = Replace(Filterstr," ","") : Filterstr = Replace(Filterstr,"'","") : Filterstr = Replace(Filterstr,"""","") : Filterstr = Replace(Filterstr,"=","") : Filterstr = Replace(Filterstr,"*","")    
 End Function   
     
 Public Function Loadfile()    
  Dim Obj    
  'On Error Resume Next   
  Set Obj = Server.Createobject("adodb.stream")    
  With Obj    
  .Type = 2 : .Mode = 3 : .Open : .Charset = Code : .Position = Obj.Size : .Loadfromfile Server.Mappath(Template) : Content = .ReadText : .Close    
  End With   
  Set Obj = Nothing   
  If Err Then Response.Write "<font color=red>无法加载模板[" & Template & "]</font>" : Response.End   
 End Function   
End Class   
%>   
