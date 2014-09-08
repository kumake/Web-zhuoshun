<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''根据Net环境下的数据库层改编
''作者：一支笔
''创建日期：2008-4-29
''修改：
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const DTAccess	 = 1
Const DTMSSQL	 = 2
Const DTMySQL	 = 3
Const DTDB2		 = 4
Const DTOracle	 = 5
Const DTExcel	 = 6
Const adStateClosed = 0

Class HotCMS_Class_Conn
	Public Ver,ClassName	'类版本,名称
	Public Db
	Public Opened
	Public DbType
	Public SqlLocalName,SqlDatabaseName,SqlUsername,SqlPassword
	Private ConnStr
	Public Conn
	Public QueryTimes	'查询次数
	Public LocalCache	'本地记录集缓存，防止重复查询造成浪费
	Public sitename
	
	'类初始化
	Private Sub Class_Initialize()
		'F.DebugStep "开始连接类"
		Ver = "3.6"
		ClassName = "HotCMS_Class_Conn"
		Opened = False
		QueryTimes = 0
		sitename="南京浩麒装饰工程有限公司"
	End Sub

	'类结束时执行的过程
	Private Sub Class_Terminate()
		F.DebugStep "结束连接类"
		If IsObject(Conn) Then Conn.Close
		If IsObject(Conn) Then Set Conn = Nothing
		If IsObject(LocalCache) Then Set LocalCache = Nothing
	End Sub

	'打开数据库(其实，仅仅是根据参数构造连接字符串而已，并没有真正的建立数据库连接真正打开链接，是在第一次请求执行SQL时)
	Public Sub Open(DbName, nDbType)


		Db = DbName
		If An("ConnStr_" & Db) <> "" Then Exit Sub	'缓存链接字符串

		Dim tt	'实践证明，就算在If End If语句里面定义，一样是过程级变量

		'下面是数据库的链接设置(支持六种数据库，AC，MSExcel，MSSQL，MYSQL, DB2, ORACLE)
		DbType = Cint("0" & nDbType)

		If DbType>6 Or DbType<1 Then DbType = DTAccess
		Select Case DbType
			Case DTAccess
			'****************************************************************************		
			'Access数据库的链接
				Db = Server.MapPath(Db)
				'Jet OLEDB:Database Password=123
				ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Db &";Jet OLEDB:Database Password=sp_cms_manage;"
				'response.Write ConnStr
			Case DTExcel
			'****************************************************************************		
			'Excel数据库的链接
				Db = Server.MapPath(Db)
				response.Write Db
				ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;Hdr=Yes"";Data Source = " & Db
			Case DTMSSQL
			'****************************************************************************		
			'SQL Server数据库的链接(更改为自己的数据库名,用户帐号和密码,服务器名)
				tt = split(Db,"$")
				SqlLocalName = tt(0)
				SqlDatabaseName = tt(1)
				SqlUsername = tt(2)
				SqlPassword = tt(3)
				ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
			Case DTMySQL
			'****************************************************************************		
			'MySQL Server数据库的链接(更改为自己的数据库名,用户帐号和密码,服务器名)
				tt = split(Db,"$$$")
				SqlLocalName = tt(1)
				SqlDatabaseName = tt(2)
				SqlUsername = tt(3)
				SqlPassword = tt(4)
				ConnStr = "driver = {MySQL ODBC 3.51 Driver}; UID = " & SqlUsername & "; Pwd = " & SqlPassword & "; Database = " & SqlDatabaseName & "; Server = " & SqlLocalName & ";option = 16386"
			Case DTDB2
			'****************************************************************************		
			'DB2数据库的链接(更改为自己的数据库名,用户帐号和密码,服务器名)
				tt = split(Db,"$$$")
				SqlLocalName = tt(1)	'表模式，注意：要加小数点，如"NULLID."，这个作为表前缀，前三种模式的该变量可以设置为0长度，从而通用
				SqlDatabaseName = tt(2)	'ODBC中数据源名称
				SqlUsername = tt(3)		'DB2管理员帐号
				SqlPassword = tt(4)		'DB2管理员密码
				ConnStr = "DSN="& SqlDatabaseName &";Uid="& SqlUsername &";Pwd="& SqlPassword &";"
			Case DTOracle
			'****************************************************************************		
			'Oracle数据库的链接(更改为自己的数据库名,用户帐号和密码,服务器名)
				tt = split(Db,"$$$")
				SqlLocalName = tt(1)	'表空间，注意：要加小数点，如"SYSTEM."，这个作为表前缀，前三种模式的该变量可以设置为0长度，从而通用
				SqlDatabaseName = tt(2)	'Oracle数据库名称
				SqlUsername = tt(3)		'Oracle管理员帐号
				SqlPassword = tt(4)		'Oracle管理员密码
'				ConnStr = "Provider=MSDAORA.1;Persist Security Info=True;User ID=" & SqlUsername & ";Password=" & SqlPassword & ";Data Source="& SqlDatabaseName &";"
				ConnStr = "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;User ID=" & SqlUsername & ";Password=" & SqlPassword & ";Data Source="& SqlDatabaseName &";"
		End Select
		An("ConnStr_" & Db) = ConnStr
	End Sub
	
	'正式打开数据库
	Private Sub OpenDB()
		If Opened = True Then
			F.Raise "数据库已经打开。"
			Exit Sub
		End If

		Dim TemData
		TemData = An("ConnStr_" & Db)
		If ConnStr = "" And TemData <> "" Then ConnStr = TemData
		If ConnStr = "" Then
			F.Raise "数据库连接字串为空。"
			Exit Sub
		End If
		
		F.DebugStep "正式打开数据库：" & ConnStr

		On Error resume Next
		Err.clear
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		If Err.Number>0 Then F.Raise "大家快来看哪。。。你的系统居然不支持ADO！哈哈，你倒霉咯◎！"

		Conn.Open ConnStr
		
		If Err.Number>0 Then
			Set Conn = Nothing
			An("ConnStr_" & Db) = ""	'这次出错了,期望下一次访问能使它恢复正常
			If F.Debug Then
				F.Raise "数据库" & Db & "连接出错，请检查连接字串。" & ConnStr & " " & Err.Description
			Else
				F.Raise "数据库连接出错，请检查连接字串。"
			End If
			Err.Clear
		End If
		
		'本地记录集缓存，防止重复查询造成浪费。放在这里，可以保证在只有连接数据库的情况下才声明对象
		Set LocalCache = Server.CreateObject("Scripting.Dictionary")
		
		Opened = True
	End Sub

	'执行SQL语句，返回执行结果记录集
	Public Function Execute(sql)
		F.DebugStep "执行SQL语句：" & sql
		
		If Opened = False Then OpenDB
		
		QueryTimes = QueryTimes + 1

		On Error resume Next
		Err.clear
		Set Execute = Conn.Execute(sql)
		
		If Err.number<>0 Then
			If F.Debug Then
				F.Raise "错误的SQL字符串 " & sql & " " & Err.Description
			Else
				F.Raise "错误的SQL字符串 " & sql
			End If
			Err.Clear
		End If
		
		'清空本地缓存
		LocalCache.RemoveAll
	End Function
	
	'仅仅查询，调用GetRows返回数据，马上关键记录集
	Public Function QueryData(Sql)
		Dim key
		key = "Conn_QueryData_" & sql
		If IsObject(LocalCache) Then
			If LocalCache.Exists(key) Then
				QueryData = LocalCache(key)
				F.DebugStep "QueryData查询SQL语句（使用缓存）：" & sql
				Exit Function
			End If
		End If
		
		Dim vRs
		Set vRs = GetRs(Sql, 0)
		If Not vRs.Eof Then 
			QueryData = vRs.GetRows()
		else
			Dim tempNoQueryData(0,0)
			QueryData = tempNoQueryData
		end if
		vRs.Close
		Set vRs = Nothing
		
		'添加到缓存
		LocalCache.Add key, QueryData
	End Function
	
	'查询，取得某一行的数组
	Public Function QueryRow(sql, rowIndex)
		Dim vdata, i, tdata
		vdata = QueryData(sql)
		If Not IsArray(vdata) Then Exit Function
		If Ubound(vdata, 1)<0 Then Exit Function	'没有列
		If Ubound(vdata, 2)<0 Then Exit Function	'没有行
		Redim tdata(Ubound(vdata, 1))	'一行数据，0开始
		For i=0 To Ubound(vdata, 1)
			tdata(i) = vdata(i, rowIndex)
		Next
		QueryRow = tdata
	End Function
	
	'仅仅查询，调用GetRows返回数据，马上关键记录集
	Public Function QueryRows(ByVal Sql, ByVal maximumRows, ByVal fields)
		Dim key
		key = "Conn_QueryRows_" & sql & "_"
		If IsArray(fields) Then
			key = key & maximumRows & "_" & Join(fields, "_")
		ElseIf maximumRows>0 Then
			key = key & maximumRows
		End If
		
		If IsObject(LocalCache) Then
			If LocalCache.Exists(key) Then
				QueryRows = LocalCache(key)
				F.DebugStep "QueryRows查询SQL语句（使用缓存）：" & sql
				Exit Function
			End If
		End If
		
		Dim vRs
		Set vRs = GetRs(Sql, 0)
		If Not vRs.Eof Then
			If IsArray(fields) Then
				QueryRows = vRs.GetRows(maximumRows, 0, fields)
			ElseIf maximumRows>0 Then
				QueryRows = vRs.GetRows(maximumRows)
			Else
				QueryRows = vRs.GetRows()
			End If
		End If
		vRs.Close
		Set vRs = Nothing
		
		'添加到缓存
		LocalCache.Add key, QueryRows
	End Function
	
	'仅仅查询
	Public Function Query(Sql)
		Set Query = GetRs(Sql, 0)
	End Function
	
	'Oracle分页，使用rownum分页
	'查询语句，开始行（1是第一行），最大行数
	Public Function QueryPageForOracle(ByVal sql, ByVal startRowIndex, ByVal maximumRows)
		If startRowIndex<=1 Then	'取第一页
			If maximumRows>=1 Then sql = "Select * From (" & sql & ") a Where rownum<=" & maximumRows
		Else	'取任意页
			If maximumRows<1 Then	'取某一行以后的所有记录
				sql = "Select * From (" & sql & ") a Where rownum>=" & startRowIndex
			Else
				sql = "Select * From (Select a.*, rownum as my_rownum From (" & sql & ") a Where rownum<=" & startRowIndex+maximumRows & ") b Where my_rownum>=" & startRowIndex
			End If
		End If

		QueryPageForOracle = sql
	End Function

	'数字分页查询。不支持Oracle
	'选择字段，表，条件，关键字，是否升序，开始行（1是第一行），最大行数
	'where不带order by
	Public Function QueryPageByNum(ByVal selects, ByVal table, ByVal where, ByVal key, ByVal isAsc, ByVal startRowIndex, ByVal maximumRows)
		Dim sql
		If Not F.IsNullOrEmpty(where) Then
			where = " Where (" & where & ") "	'这个括号很重要，因为可能where里面有or语句，而后面用and连接，就会影响逻辑
		Else
			where = " Where 1=1 "
		End If

		Dim order
		order = " Order By " & key
		If isAsc Then
			order = order & " Asc"
		Else
			order = order & " Desc"
		End If

		If startRowIndex<=1 Then	'取所有记录或者仅取第一页
			If maximumRows<1 Then	'所有记录
				sql = "Select " & selects & " From " & table & where & order
			Else	'第一页，使用Top
				sql = "Select Top " & maximumRows & " " & selects & " From " & table & where & order
			End If
		Else	'取任意页
			If maximumRows<1 Then	'取某一行以后的所有记录
				sql = "Select " & selects & " From " & table & where & " And " & key
				If isAsc Then
					sql = sql & ">(Select Max"
				Else
					sql = sql & "<(Select Min"
				End If
				sql = sql & "(" & key & ") From (Select Top " & (startRowIndex-1)*maximumRows & " " & selects & " From " & table & where & order & ") a)"
			Else	'取一定行数
				sql = "Select Top " & maximumRows & " " & selects & " From " & table & where & " And " & key
				If isAsc Then
					sql = sql & ">(Select Max"
				Else
					sql = sql & "<(Select Min"
				End If
				''''''''''''''''''''				
				if instr(key,".")>0 then
					Dim tempval:tempval = split(key,".")
					key = tempval(1)
				end if
				''''''''''''''''''''
				sql = sql & "(" & key & ") From (Select Top " & (startRowIndex-1)*maximumRows & " " & selects & " From " & table & where & order & ") a) " & order
			End If
		End If

		QueryPageByNum = sql
	End Function

	'2007-05-01 23:40 增加
	'NotIn分页查询。不支持Oracle
	'选择字段，表，条件，关键字，排序，开始行（1是第一行），最大行数
	'where不带order by
	'排序不带order by
	Public Function QueryPageByNotIn(ByVal selects, ByVal table, ByVal where, ByVal key, ByVal order, ByVal isAsc, ByVal startRowIndex, ByVal maximumRows)
		Dim sql
		If Not F.IsNullOrEmpty(where) Then
			where = " Where (" & where & ") "	'这个括号很重要，因为可能where里面有or语句，而后面用and连接，就会影响逻辑
		Else
			where = " Where 1=1 "
		End If
		
		If isAsc Then
			order = order & " Asc"
		Else
			order = order & " Desc"
		End If


		If startRowIndex<=1 Then	'取所有记录或者仅取第一页
			If maximumRows<1 Then	'所有记录
				sql = "Select " & selects & " From " & table & where & " Order By " & order
			Else	'第一页，使用Top
				sql = "Select Top " & maximumRows & " " & selects & " From " & table & where & " Order By " & order
			End If
		Else	'取任意页
			If maximumRows<1 Then	'取某一行以后的所有记录
				sql = "Select " & selects & " From " & table & where & " And " & key & "Not In(Select Top " & startRowIndex-1 & " " & key & " From " & table & where & " Order By " & order & ")" & " Order By " & order
			Else	'取一定行数
				sql = "Select Top " & maximumRows & " " & selects & " From " & table & where & " And " & key & " Not In(Select Top " & (startRowIndex-1)*maximumRows & " " & key & " From " & table & where & " Order By " & order & ") Order By " & order
			End If
		End If
		
		QueryPageByNotIn = sql
	End Function

	'获取记录集
	Public Function GetRs(sql, w)	'w表示是否允许修改记录集数据,0否1是
		F.DebugStep "查询SQL语句：" & sql	'调试模式才使用该状态记录

		If Opened = False Then OpenDB

		QueryTimes = QueryTimes + 1

		If w<>0 And w<>1 Then w = 0

		'清空本地缓存
		If w=1 Then LocalCache.RemoveAll
		
		On Error resume Next
		Err.clear
		Set GetRs = Server.CreateObject("ADODB.RecordSet")
		GetRs.open sql, Conn, 1, w * 2 + 1

		If Err.Number <> 0 Then
			If F.Debug Then
				F.Raise "错误的SQL字符串:" & sql & "<BR>" & Err.Description
			Else
				F.Raise "错误的SQL字符串"
			End If
		End If
		
		If GetRs.State = adStateClosed Then F.Raise "没有打开记录集"
		
		'调试输出是否EOF，是否BOF，以及记录数
		F.T "EOF = " & GetRs.Eof & "  BOF = " & GetRs.Bof & "  RecordCount = " & GetRs.RecordCount
	End Function

	'取得表模式
	Public Default Property Get Data(GetType, TableName)
		If F.Debug Then
			If Not IsArray(TableName) Then
				F.DebugStep "取数据:" & GetType & " -- " & TableName
			Else
				F.DebugStep "取数据:" & GetType & " -- " & Join(TableName,", ")
			End If
		End If
		
		If Opened = False Then OpenDB
		
		Const adSchemaTables = 20
		Const adSchemaColumns = 4
		Const adSchemaPrimaryKeys = 28
		Const adSchemaForeignKeys = 27
		
		On Error Resume Next
		Select Case LCase(GetType)
			Case "tables"
				'TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE

'				Set Data = Conn.OpenSchema(adSchemaTables,Array(empty,empty,empty,"table"))
				Set Data = Conn.OpenSchema(adSchemaTables,Array(empty,empty,empty,TableName))
			Case "table"
				'TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE
'				Set Data = Conn.OpenSchema(adSchemaTables,Array(empty,empty,empty,"table"))
				Set Data = Conn.OpenSchema(adSchemaTables,Array(empty,empty,TableName,empty))
			Case "columns"
				'TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,COLUMN_NAME
				Set Data = Conn.OpenSchema(adSchemaColumns,Array(empty,empty,TableName,empty))
			Case "primarykeys"
				'PK_TABLE_CATALOG,PK_TABLE_SCHEMA,PK_TABLE_NAME
				Set Data = Conn.OpenSchema(adSchemaPrimaryKeys,Array(empty,empty,TableName))
			Case "foreignkeys"
				'PK_TABLE_CATALOG,PK_TABLE_SCHEMA,PK_TABLE_NAME
				Set Data = Conn.OpenSchema(adSchemaForeignKeys,Array(empty,empty,TableName))
			Case "sql"
				Set Data = Conn.Execute(TableName)
			Case else
				Set Data = Conn.OpenSchema(GetType, TableName)
		End select
		F.T "EOF = " & Data.Eof & "  BOF = " & Data.Bof & "  RecordCount = " & Data.RecordCount
	End Property

    '函数:SqlServer(97-2000) to Access(97-2000)
    '参数:Sql,数据库类型(ACCESS,SQLSERVER)
    '说明:
    Public Function SqlServer_To_Access(Sql)
        Dim regEx, Matches, Match
        '创建正则对象
        Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        regEx.MultiLine = True

        '转:GetDate()
        regEx.Pattern = "(?=[^']?)GETDATE\(\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"NOW()")

        '转:UPPER()
        regEx.Pattern = "(?=[^']?)UPPER\([\s]?(.+?)[\s]?\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"UCASE($1)")

        '转:日期表示方式
        '说明:时间格式必须是2004-23-23 11:11:10 标准格式
        regEx.Pattern = "'([\d]{4,4}\-[\d]{1,2}\-[\d]{1,2}(?:[\s][\d]{1,2}:[\d]{1,2}:[\d]{1,2})?)'"
        Sql = regEx.Replace(Sql,"#$1#")
        
        regEx.Pattern = "DATEDIFF\([\s]?(second|minute|hour|day|month|year)[\s]?\,[\s]?(.+?)[\s]?\,[\s]?(.+?)([\s]?\)[\s]?)"
        Set Matches = regEx.ExeCute(Sql)
        Dim temStr
        For Each Match In Matches
            temStr = "DATEDIFF("
            Select Case lcase(Match.SubMatches(0))
                Case "second" :
                    temStr = temStr & "'s'"
                Case "minute" :
                    temStr = temStr & "'n'"
                Case "hour" :
                    temStr = temStr & "'h'"
                Case "day" :
                    temStr = temStr & "'d'"
                Case "month" :
                    temStr = temStr & "'m'"
                Case "year" :
                    temStr = temStr & "'y'"
            End Select
            temStr = temStr & "," & Match.SubMatches(1) & "," &  Match.SubMatches(2) & Match.SubMatches(3)
            Sql = Replace(Sql,Match.Value,temStr,1,1)
        Next

        '转:Insert函数
        regEx.Pattern = "CHARINDEX\([\s]?'(.+?)'[\s]?,[\s]?'(.+?)'[\s]?\)[\s]?"
        Sql = regEx.Replace(Sql,"INSTR('$2','$1')")

        Set regEx = Nothing
        SqlServer_To_Access = Sql
    End Function
	
	'显示记录集。供调试时，用一个表格显示一个记录集的所有数据
	Public Sub ShowRs(ByRef iRs)
		F.DebugStep "显示记录集"
		Dim i
		
		If Not IsObject(iRs) Then
			F.Raise "类型不匹配，非记录集类型!"
			Exit Sub
		End If
		F.O "<table class=""ListView""><thead><tr>" & vbNewLine
		For i = 0 To iRs.fields.Count-1
			F.O "<th>"&iRs(i).name&"</th>"
		Next
		F.O "</tr></thead><tbody>" & vbNewLine
		If iRs.eof And iRs.bof Then
			F.Raise "记录集为空，没有任何数据!"
'			Exit Sub
		Else
			iRs.MoveFirst
			While Not iRs.eof
				F.O "<tr>" & vbNewLine
				For i = 0 To iRs.fields.Count-1
					If iRs(i).Type=205 Then
						F.O "<td>OLE 对象</td>"
					Else
						F.O "<td>" & Server.HTMLEncode(iRs(i) & "") & "</td>"
					End If
				Next
				F.O  vbNewLine & "</tr>" & vbNewLine
				iRs.MoveNext
			Wend
			iRs.MoveFirst
		End If
		F.O "</tbody></table>"
	End Sub

	'把一个记录集转成XML文本
	Public Function RsToXML(ByRef iRs)
		F.DebugStep "记录集转为XML"

		If Not IsObject(iRs) Then
			F.Raise "记录集为空，没有任何数据!"
			Exit Function
		End If

		RsToXML = ""
		If Not IsObject(iRs) Or iRs.RecordCount<1 Then Exit Function
		Dim i
		While Not iRs.eof
			RsToXML = RsToXML & vbNewLine & "<RsData>"
			For i = 0 To iRs.Fields.Count-1
				RsToXML = RsToXML & vbNewLine & "	<" & iRs(i).name & ">" & Server.HTMLEnCode(iRs(i).value & "") & "</" & iRs(i).name & ">"
			Next
			RsToXML = RsToXML & vbNewLine & "</RsData>"
			iRs.MoveNext
		Wend
	End Function

	'建立数据库	
	Public Function CreateDatabase(filename)
		F.DebugStep "建立数据库 " & filename

		Dim Cat
		
		If Mid(filename, 2, 1) <> ":" Then filename = Server.MapPath(filename)
		Set Cat = Server.CreateObject("ADOX.Catalog") 
		Cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filename)
		Set Cat = Nothing
	End Function

	'Oracle函数，建立序列
	Public Function CreateSequence(vTableName)
		F.DebugStep "为 " & vTableName & " 创建序列"
		Con.ExeCute "DROP SEQUENCE Sequence_" & Replace(vTableName, """", "") & ""
		Con.ExeCute "CREATE SEQUENCE Sequence_" & Replace(vTableName, """", "") & " INCREMENT BY 1 START " _
					& "    WITH 100 MAXVALUE 1.0E10 MINVALUE 1 NOCYCLE " _
					& "    CACHE 20 ORDER "
	End Function

	'Oracle函数，建立触发器
	Public Function CreateTrigger(vTableName, vIDName)
		F.DebugStep "为 " & vTableName & " 创建触发器"
		Con.ExeCute "Create Or Replace Trigger Trigger_" & Replace(vTableName, """", "") & " " _
			& "before insert on       " & vTableName & " " _
			& "referencing old as old new as new for each row " _
			& "begin " _
			& "select  Sequence_" & Replace(vTableName, """", "") & ".nextval into :new." & vIDName & " from dual; " _
			& "end;"
	End Function

	'Oracle函数，建立主键
	Public Function CreatePK(vTableName, vIDName)
		F.DebugStep "为 " & vTableName & " 创建主键 " & vIDName
		Con.ExeCute "ALTER TABLE " & vTableName & " " _
				& "ADD (CONSTRAINT ""PK_" & Replace(vTableName, """", "") & """ PRIMARY KEY(""" & vIDName & """)) "
	End Function
	
	'重命名表名
	Public Function RenameTable(vTableName, vNewTableName)
		F.DebugStep "重命名 " & vTableName & " 为 " & vNewTableName
		On Error Resume Next
		Dim ADOX
		Set ADOX = Server.CreateObject("ADOX.Catalog")
		Set ADOX.ActiveConnection = Conn
		If Err.Number>0 Then
			F.T "呵呵，服务器不支持ADOX呀！"
		End If
		ADOX.Tables(vTableName).Name = vNewTableName
		If Err.Number>0 Then
			F.T Err.Description
			Err.Clear
		End If
	End Function
End Class
%>