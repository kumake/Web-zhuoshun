<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''����Net�����µ����ݿ��ı�
''���ߣ�һ֧��
''�������ڣ�2008-4-29
''�޸ģ�
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
	Public Ver,ClassName	'��汾,����
	Public Db
	Public Opened
	Public DbType
	Public SqlLocalName,SqlDatabaseName,SqlUsername,SqlPassword
	Private ConnStr
	Public Conn
	Public QueryTimes	'��ѯ����
	Public LocalCache	'���ؼ�¼�����棬��ֹ�ظ���ѯ����˷�
	Public sitename
	
	'���ʼ��
	Private Sub Class_Initialize()
		'F.DebugStep "��ʼ������"
		Ver = "3.6"
		ClassName = "HotCMS_Class_Conn"
		Opened = False
		QueryTimes = 0
		sitename="�Ͼ�����װ�ι������޹�˾"
	End Sub

	'�����ʱִ�еĹ���
	Private Sub Class_Terminate()
		F.DebugStep "����������"
		If IsObject(Conn) Then Conn.Close
		If IsObject(Conn) Then Set Conn = Nothing
		If IsObject(LocalCache) Then Set LocalCache = Nothing
	End Sub

	'�����ݿ�(��ʵ�������Ǹ��ݲ������������ַ������ѣ���û�������Ľ������ݿ��������������ӣ����ڵ�һ������ִ��SQLʱ)
	Public Sub Open(DbName, nDbType)


		Db = DbName
		If An("ConnStr_" & Db) <> "" Then Exit Sub	'���������ַ���

		Dim tt	'ʵ��֤����������If End If������涨�壬һ���ǹ��̼�����

		'���������ݿ����������(֧���������ݿ⣬AC��MSExcel��MSSQL��MYSQL, DB2, ORACLE)
		DbType = Cint("0" & nDbType)

		If DbType>6 Or DbType<1 Then DbType = DTAccess
		Select Case DbType
			Case DTAccess
			'****************************************************************************		
			'Access���ݿ������
				Db = Server.MapPath(Db)
				'Jet OLEDB:Database Password=123
				ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Db &";Jet OLEDB:Database Password=sp_cms_manage;"
				'response.Write ConnStr
			Case DTExcel
			'****************************************************************************		
			'Excel���ݿ������
				Db = Server.MapPath(Db)
				response.Write Db
				ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;Hdr=Yes"";Data Source = " & Db
			Case DTMSSQL
			'****************************************************************************		
			'SQL Server���ݿ������(����Ϊ�Լ������ݿ���,�û��ʺź�����,��������)
				tt = split(Db,"$")
				SqlLocalName = tt(0)
				SqlDatabaseName = tt(1)
				SqlUsername = tt(2)
				SqlPassword = tt(3)
				ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
			Case DTMySQL
			'****************************************************************************		
			'MySQL Server���ݿ������(����Ϊ�Լ������ݿ���,�û��ʺź�����,��������)
				tt = split(Db,"$$$")
				SqlLocalName = tt(1)
				SqlDatabaseName = tt(2)
				SqlUsername = tt(3)
				SqlPassword = tt(4)
				ConnStr = "driver = {MySQL ODBC 3.51 Driver}; UID = " & SqlUsername & "; Pwd = " & SqlPassword & "; Database = " & SqlDatabaseName & "; Server = " & SqlLocalName & ";option = 16386"
			Case DTDB2
			'****************************************************************************		
			'DB2���ݿ������(����Ϊ�Լ������ݿ���,�û��ʺź�����,��������)
				tt = split(Db,"$$$")
				SqlLocalName = tt(1)	'��ģʽ��ע�⣺Ҫ��С���㣬��"NULLID."�������Ϊ��ǰ׺��ǰ����ģʽ�ĸñ�����������Ϊ0���ȣ��Ӷ�ͨ��
				SqlDatabaseName = tt(2)	'ODBC������Դ����
				SqlUsername = tt(3)		'DB2����Ա�ʺ�
				SqlPassword = tt(4)		'DB2����Ա����
				ConnStr = "DSN="& SqlDatabaseName &";Uid="& SqlUsername &";Pwd="& SqlPassword &";"
			Case DTOracle
			'****************************************************************************		
			'Oracle���ݿ������(����Ϊ�Լ������ݿ���,�û��ʺź�����,��������)
				tt = split(Db,"$$$")
				SqlLocalName = tt(1)	'��ռ䣬ע�⣺Ҫ��С���㣬��"SYSTEM."�������Ϊ��ǰ׺��ǰ����ģʽ�ĸñ�����������Ϊ0���ȣ��Ӷ�ͨ��
				SqlDatabaseName = tt(2)	'Oracle���ݿ�����
				SqlUsername = tt(3)		'Oracle����Ա�ʺ�
				SqlPassword = tt(4)		'Oracle����Ա����
'				ConnStr = "Provider=MSDAORA.1;Persist Security Info=True;User ID=" & SqlUsername & ";Password=" & SqlPassword & ";Data Source="& SqlDatabaseName &";"
				ConnStr = "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;User ID=" & SqlUsername & ";Password=" & SqlPassword & ";Data Source="& SqlDatabaseName &";"
		End Select
		An("ConnStr_" & Db) = ConnStr
	End Sub
	
	'��ʽ�����ݿ�
	Private Sub OpenDB()
		If Opened = True Then
			F.Raise "���ݿ��Ѿ��򿪡�"
			Exit Sub
		End If

		Dim TemData
		TemData = An("ConnStr_" & Db)
		If ConnStr = "" And TemData <> "" Then ConnStr = TemData
		If ConnStr = "" Then
			F.Raise "���ݿ������ִ�Ϊ�ա�"
			Exit Sub
		End If
		
		F.DebugStep "��ʽ�����ݿ⣺" & ConnStr

		On Error resume Next
		Err.clear
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		If Err.Number>0 Then F.Raise "��ҿ������ġ��������ϵͳ��Ȼ��֧��ADO���������㵹ù����"

		Conn.Open ConnStr
		
		If Err.Number>0 Then
			Set Conn = Nothing
			An("ConnStr_" & Db) = ""	'��γ�����,������һ�η�����ʹ���ָ�����
			If F.Debug Then
				F.Raise "���ݿ�" & Db & "���ӳ������������ִ���" & ConnStr & " " & Err.Description
			Else
				F.Raise "���ݿ����ӳ������������ִ���"
			End If
			Err.Clear
		End If
		
		'���ؼ�¼�����棬��ֹ�ظ���ѯ����˷ѡ�����������Ա�֤��ֻ���������ݿ������²���������
		Set LocalCache = Server.CreateObject("Scripting.Dictionary")
		
		Opened = True
	End Sub

	'ִ��SQL��䣬����ִ�н����¼��
	Public Function Execute(sql)
		F.DebugStep "ִ��SQL��䣺" & sql
		
		If Opened = False Then OpenDB
		
		QueryTimes = QueryTimes + 1

		On Error resume Next
		Err.clear
		Set Execute = Conn.Execute(sql)
		
		If Err.number<>0 Then
			If F.Debug Then
				F.Raise "�����SQL�ַ��� " & sql & " " & Err.Description
			Else
				F.Raise "�����SQL�ַ��� " & sql
			End If
			Err.Clear
		End If
		
		'��ձ��ػ���
		LocalCache.RemoveAll
	End Function
	
	'������ѯ������GetRows�������ݣ����Ϲؼ���¼��
	Public Function QueryData(Sql)
		Dim key
		key = "Conn_QueryData_" & sql
		If IsObject(LocalCache) Then
			If LocalCache.Exists(key) Then
				QueryData = LocalCache(key)
				F.DebugStep "QueryData��ѯSQL��䣨ʹ�û��棩��" & sql
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
		
		'��ӵ�����
		LocalCache.Add key, QueryData
	End Function
	
	'��ѯ��ȡ��ĳһ�е�����
	Public Function QueryRow(sql, rowIndex)
		Dim vdata, i, tdata
		vdata = QueryData(sql)
		If Not IsArray(vdata) Then Exit Function
		If Ubound(vdata, 1)<0 Then Exit Function	'û����
		If Ubound(vdata, 2)<0 Then Exit Function	'û����
		Redim tdata(Ubound(vdata, 1))	'һ�����ݣ�0��ʼ
		For i=0 To Ubound(vdata, 1)
			tdata(i) = vdata(i, rowIndex)
		Next
		QueryRow = tdata
	End Function
	
	'������ѯ������GetRows�������ݣ����Ϲؼ���¼��
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
				F.DebugStep "QueryRows��ѯSQL��䣨ʹ�û��棩��" & sql
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
		
		'��ӵ�����
		LocalCache.Add key, QueryRows
	End Function
	
	'������ѯ
	Public Function Query(Sql)
		Set Query = GetRs(Sql, 0)
	End Function
	
	'Oracle��ҳ��ʹ��rownum��ҳ
	'��ѯ��䣬��ʼ�У�1�ǵ�һ�У����������
	Public Function QueryPageForOracle(ByVal sql, ByVal startRowIndex, ByVal maximumRows)
		If startRowIndex<=1 Then	'ȡ��һҳ
			If maximumRows>=1 Then sql = "Select * From (" & sql & ") a Where rownum<=" & maximumRows
		Else	'ȡ����ҳ
			If maximumRows<1 Then	'ȡĳһ���Ժ�����м�¼
				sql = "Select * From (" & sql & ") a Where rownum>=" & startRowIndex
			Else
				sql = "Select * From (Select a.*, rownum as my_rownum From (" & sql & ") a Where rownum<=" & startRowIndex+maximumRows & ") b Where my_rownum>=" & startRowIndex
			End If
		End If

		QueryPageForOracle = sql
	End Function

	'���ַ�ҳ��ѯ����֧��Oracle
	'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
	'where����order by
	Public Function QueryPageByNum(ByVal selects, ByVal table, ByVal where, ByVal key, ByVal isAsc, ByVal startRowIndex, ByVal maximumRows)
		Dim sql
		If Not F.IsNullOrEmpty(where) Then
			where = " Where (" & where & ") "	'������ź���Ҫ����Ϊ����where������or��䣬��������and���ӣ��ͻ�Ӱ���߼�
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

		If startRowIndex<=1 Then	'ȡ���м�¼���߽�ȡ��һҳ
			If maximumRows<1 Then	'���м�¼
				sql = "Select " & selects & " From " & table & where & order
			Else	'��һҳ��ʹ��Top
				sql = "Select Top " & maximumRows & " " & selects & " From " & table & where & order
			End If
		Else	'ȡ����ҳ
			If maximumRows<1 Then	'ȡĳһ���Ժ�����м�¼
				sql = "Select " & selects & " From " & table & where & " And " & key
				If isAsc Then
					sql = sql & ">(Select Max"
				Else
					sql = sql & "<(Select Min"
				End If
				sql = sql & "(" & key & ") From (Select Top " & (startRowIndex-1)*maximumRows & " " & selects & " From " & table & where & order & ") a)"
			Else	'ȡһ������
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

	'2007-05-01 23:40 ����
	'NotIn��ҳ��ѯ����֧��Oracle
	'ѡ���ֶΣ����������ؼ��֣����򣬿�ʼ�У�1�ǵ�һ�У����������
	'where����order by
	'���򲻴�order by
	Public Function QueryPageByNotIn(ByVal selects, ByVal table, ByVal where, ByVal key, ByVal order, ByVal isAsc, ByVal startRowIndex, ByVal maximumRows)
		Dim sql
		If Not F.IsNullOrEmpty(where) Then
			where = " Where (" & where & ") "	'������ź���Ҫ����Ϊ����where������or��䣬��������and���ӣ��ͻ�Ӱ���߼�
		Else
			where = " Where 1=1 "
		End If
		
		If isAsc Then
			order = order & " Asc"
		Else
			order = order & " Desc"
		End If


		If startRowIndex<=1 Then	'ȡ���м�¼���߽�ȡ��һҳ
			If maximumRows<1 Then	'���м�¼
				sql = "Select " & selects & " From " & table & where & " Order By " & order
			Else	'��һҳ��ʹ��Top
				sql = "Select Top " & maximumRows & " " & selects & " From " & table & where & " Order By " & order
			End If
		Else	'ȡ����ҳ
			If maximumRows<1 Then	'ȡĳһ���Ժ�����м�¼
				sql = "Select " & selects & " From " & table & where & " And " & key & "Not In(Select Top " & startRowIndex-1 & " " & key & " From " & table & where & " Order By " & order & ")" & " Order By " & order
			Else	'ȡһ������
				sql = "Select Top " & maximumRows & " " & selects & " From " & table & where & " And " & key & " Not In(Select Top " & (startRowIndex-1)*maximumRows & " " & key & " From " & table & where & " Order By " & order & ") Order By " & order
			End If
		End If
		
		QueryPageByNotIn = sql
	End Function

	'��ȡ��¼��
	Public Function GetRs(sql, w)	'w��ʾ�Ƿ������޸ļ�¼������,0��1��
		F.DebugStep "��ѯSQL��䣺" & sql	'����ģʽ��ʹ�ø�״̬��¼

		If Opened = False Then OpenDB

		QueryTimes = QueryTimes + 1

		If w<>0 And w<>1 Then w = 0

		'��ձ��ػ���
		If w=1 Then LocalCache.RemoveAll
		
		On Error resume Next
		Err.clear
		Set GetRs = Server.CreateObject("ADODB.RecordSet")
		GetRs.open sql, Conn, 1, w * 2 + 1

		If Err.Number <> 0 Then
			If F.Debug Then
				F.Raise "�����SQL�ַ���:" & sql & "<BR>" & Err.Description
			Else
				F.Raise "�����SQL�ַ���"
			End If
		End If
		
		If GetRs.State = adStateClosed Then F.Raise "û�д򿪼�¼��"
		
		'��������Ƿ�EOF���Ƿ�BOF���Լ���¼��
		F.T "EOF = " & GetRs.Eof & "  BOF = " & GetRs.Bof & "  RecordCount = " & GetRs.RecordCount
	End Function

	'ȡ�ñ�ģʽ
	Public Default Property Get Data(GetType, TableName)
		If F.Debug Then
			If Not IsArray(TableName) Then
				F.DebugStep "ȡ����:" & GetType & " -- " & TableName
			Else
				F.DebugStep "ȡ����:" & GetType & " -- " & Join(TableName,", ")
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

    '����:SqlServer(97-2000) to Access(97-2000)
    '����:Sql,���ݿ�����(ACCESS,SQLSERVER)
    '˵��:
    Public Function SqlServer_To_Access(Sql)
        Dim regEx, Matches, Match
        '�����������
        Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        regEx.MultiLine = True

        'ת:GetDate()
        regEx.Pattern = "(?=[^']?)GETDATE\(\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"NOW()")

        'ת:UPPER()
        regEx.Pattern = "(?=[^']?)UPPER\([\s]?(.+?)[\s]?\)(?=[^']?)"
        Sql = regEx.Replace(Sql,"UCASE($1)")

        'ת:���ڱ�ʾ��ʽ
        '˵��:ʱ���ʽ������2004-23-23 11:11:10 ��׼��ʽ
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

        'ת:Insert����
        regEx.Pattern = "CHARINDEX\([\s]?'(.+?)'[\s]?,[\s]?'(.+?)'[\s]?\)[\s]?"
        Sql = regEx.Replace(Sql,"INSTR('$2','$1')")

        Set regEx = Nothing
        SqlServer_To_Access = Sql
    End Function
	
	'��ʾ��¼����������ʱ����һ�������ʾһ����¼������������
	Public Sub ShowRs(ByRef iRs)
		F.DebugStep "��ʾ��¼��"
		Dim i
		
		If Not IsObject(iRs) Then
			F.Raise "���Ͳ�ƥ�䣬�Ǽ�¼������!"
			Exit Sub
		End If
		F.O "<table class=""ListView""><thead><tr>" & vbNewLine
		For i = 0 To iRs.fields.Count-1
			F.O "<th>"&iRs(i).name&"</th>"
		Next
		F.O "</tr></thead><tbody>" & vbNewLine
		If iRs.eof And iRs.bof Then
			F.Raise "��¼��Ϊ�գ�û���κ�����!"
'			Exit Sub
		Else
			iRs.MoveFirst
			While Not iRs.eof
				F.O "<tr>" & vbNewLine
				For i = 0 To iRs.fields.Count-1
					If iRs(i).Type=205 Then
						F.O "<td>OLE ����</td>"
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

	'��һ����¼��ת��XML�ı�
	Public Function RsToXML(ByRef iRs)
		F.DebugStep "��¼��תΪXML"

		If Not IsObject(iRs) Then
			F.Raise "��¼��Ϊ�գ�û���κ�����!"
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

	'�������ݿ�	
	Public Function CreateDatabase(filename)
		F.DebugStep "�������ݿ� " & filename

		Dim Cat
		
		If Mid(filename, 2, 1) <> ":" Then filename = Server.MapPath(filename)
		Set Cat = Server.CreateObject("ADOX.Catalog") 
		Cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filename)
		Set Cat = Nothing
	End Function

	'Oracle��������������
	Public Function CreateSequence(vTableName)
		F.DebugStep "Ϊ " & vTableName & " ��������"
		Con.ExeCute "DROP SEQUENCE Sequence_" & Replace(vTableName, """", "") & ""
		Con.ExeCute "CREATE SEQUENCE Sequence_" & Replace(vTableName, """", "") & " INCREMENT BY 1 START " _
					& "    WITH 100 MAXVALUE 1.0E10 MINVALUE 1 NOCYCLE " _
					& "    CACHE 20 ORDER "
	End Function

	'Oracle����������������
	Public Function CreateTrigger(vTableName, vIDName)
		F.DebugStep "Ϊ " & vTableName & " ����������"
		Con.ExeCute "Create Or Replace Trigger Trigger_" & Replace(vTableName, """", "") & " " _
			& "before insert on       " & vTableName & " " _
			& "referencing old as old new as new for each row " _
			& "begin " _
			& "select  Sequence_" & Replace(vTableName, """", "") & ".nextval into :new." & vIDName & " from dual; " _
			& "end;"
	End Function

	'Oracle��������������
	Public Function CreatePK(vTableName, vIDName)
		F.DebugStep "Ϊ " & vTableName & " �������� " & vIDName
		Con.ExeCute "ALTER TABLE " & vTableName & " " _
				& "ADD (CONSTRAINT ""PK_" & Replace(vTableName, """", "") & """ PRIMARY KEY(""" & vIDName & """)) "
	End Function
	
	'����������
	Public Function RenameTable(vTableName, vNewTableName)
		F.DebugStep "������ " & vTableName & " Ϊ " & vNewTableName
		On Error Resume Next
		Dim ADOX
		Set ADOX = Server.CreateObject("ADOX.Catalog")
		Set ADOX.ActiveConnection = Conn
		If Err.Number>0 Then
			F.T "�Ǻǣ���������֧��ADOXѽ��"
		End If
		ADOX.Tables(vTableName).Name = vNewTableName
		If Err.Number>0 Then
			F.T Err.Description
			Err.Clear
		End If
	End Function
End Class
%>