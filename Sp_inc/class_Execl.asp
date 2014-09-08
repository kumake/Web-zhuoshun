<%
'  使用方法：
'  set newExcel = New Export2Excel
'  newExcel.FilePath = "/mail/excel/"----------------------------------路径
'  newExcel.Sql = "select * from user"-------------------------------查询语句
'  newExcel.Field = "帐号||姓名||所属部门||"----------------------输出列名
'  response.write newExcel.export2Excel()
'类开始
Class Export2Excel
'声明常量、变量
    Private strFilePath,strTitle,strSql,strField,strRows,strCols
    Private strCn,strHtml,strPath,strServerPath,Filename,strExcelName
    Private objDbCn,objRs
    Private objXlsApp,objXlsWorkBook,objXlsWorkSheet
    Private arrField
    '初始化类
    Private Sub Class_Initialize()
     set objDbCn = Con
     strTitle = "查询结果"
     strFilePath="Excel/"
     strRows = 2
     strCols = 1
	 strFileName = ""
    End Sub
    '销毁类
    Private Sub Class_Terminate()
    End Sub
    '属性FilePath
    Public Property Let FilePath(value)
     strFilePath = value
     strServerPath=strFilePath
    End Property
    Public Property Get FilePath()
     FilePath = strFilePat
    End Property
    '属性Title
    Public Property Let Title(value)
     strTitle = value
    End Property
    Public Property Get Title()
     Title = strTitle
    End Property
    '属性Sql
    Public Property Let Sql(value)
     strSql = value
    End Property
    Public Property Get Sql()
     Sql = strSql
    End Property
    '属性Field
    Public Property Let Field(value)
     strField = value
    End Property
    Public Property Get Field()
     Field = strField
    End Property
    '属性Rows
    Public Property Let Rows(value)
     strRows = value
    End Property
    Public Property Get Rows()
     Rows = strRows
    End Property
    '属性Cols
    Public Property Let Cols(value)
     strCols = value
    End Property
    Public Property Get Cols()
     Cols = strCols
    End Property
    'execl文件名称
    Public Property Let ExcelName(value)
     strExcelName = value
    End Property
    Public Property Get ExcelName()
     ExcelName = strExcelName
    End Property
    '
    Public Function export2Excel()
     if strSql = "" or strField = "" then
      response.write "参数设置错误，请与管理员联系！谢谢"
      response.end
     end if
     
     strFilePath = GetFilePath(Server.mappath(strFilePath&"upload.asp"),"\")
     set objFso = createobject("scripting.filesystemobject")
     if objFso.FolderExists(strFilePath) = False then
      objFso.Createfolder(strFilePath)
     end if
	 '''''''''''''''''''''''''''''''''''''''''''''
	 if strExcelName = "" then
		 Filename=cstr(createFileName()) & ".xls"
	 else
		 Filename = strExcelName
	 end if
	 '''''''''''''''''''''''''''''''''''''''''''''
     strFileName = strFilePath & Filename 
     'objDbCn.Open()
     set objRs = objDbCn.execute(strSql)
     if objRs.EOF And objRs.BOF then
      strHtml = "抱歉，暂时没有任何合适的数据导出，如有疑问，请与管理员联系！"
     else
      set objXlsApp = server.CreateObject("Excel.Application")
      objXlsApp.Visible = false
      objXlsApp.WorkBooks.Add
      set objXlsWorkBook = objXlsApp.ActiveWorkBook
      set objXlsWorkSheet = objXlsWorkBook.WorkSheets(1)
      arrField = split(strField,"||")
      
      for f = 0 to Ubound(arrField)
       objXlsWorkSheet.Cells(1,f+1).Value = arrField(f)
       'response.Write arrField(f)&" "
      next
      'response.Write "<br>"
      objRs=objRs.getRows()
      If instr(Sql,"exportEnterprise ")=0 then
          for c=0 to ubound(objRs,2)
              If response.IsClientConnected=false then exit for '数据多导出时间很长，所以需要探测下客户端是否还在连接
              response.Write "正在导出第"&cstr(c+1)&"条<br>"
            response.Flush()
           for f = 0 to ubound(objRs,1)
                   If response.IsClientConnected=false then exit for
             objXlsWorkSheet.Cells(c+2,f+1).Value = trim(Cstr(objRs(f,c)))&VBCR
             'objXlsWorkSheet.Columns(f+1).ColumnWidth=Len(Cstr(objRs(f,c)))*2
           next
          next
          
      Else
            for c=0 to ubound(objRs,2)
              If response.IsClientConnected=false then exit for
              response.Write "正在导出第"&cstr(c+1)&"条<br>"
            response.Flush()
           for f = 0 to ubound(objRs,1)
               If response.IsClientConnected=false then exit for
            If f<>1 then
             objXlsWorkSheet.Cells(c+2,f+1).Value = trim(Cstr(objRs(f,c)))&VBCR
             'objXlsWorkSheet.Columns(f+1).ColumnWidth=Len(Cstr(objRs(f,c)))*2
            Else
             objXlsWorkSheet.Cells(c+2,f+1).Value = trim(replace(replace(Cstr(objRs(f,c)),"0",""),"|"," "))&VBCR
             'objXlsWorkSheet.Columns(f+1).ColumnWidth=Len(Cstr(objXlsWorkSheet.Cells(c+2,f+1).Value))*2            
            End If
           next
          next
      End If
      
      '必不可少，否则会出现错误
      If objFso.fileExists(strFileName)=true then
          objFso.deletefile strFileName
      End if
        response.Write "导出成功！<br>"
        response.Flush()      
  
      objXlsWorkSheet.SaveAs strFileName
      
      strHtml = "<script>location.href='" & GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")&strServerpath&Filename  & "';</script>"
      objXlsApp.Quit'重要
      set objXlsWorkSheet = nothing
      set objXlsWorkBook = nothing
      set objXlsApp = nothing
     end if
     'objDbCn.Close()
     set objRs = nothing
     if err > 0 then
      strHtml = "系统忙,请稍后重试"
     end if
     export2Excel = strHtml
    End Function
    '函数
    Public Function createFileName()
	 fName=now
	 fName=replace(fName,":","")
	 fName=replace(fName,"-","")
	 fName=replace(fName," ","")
     createFileName=fName
    End Function
        
    Public function GetFilePath(FullPath,str)
      If FullPath <> "" Then
        GetFilePath = left(FullPath,InStrRev(FullPath, str))
        Else
        GetFilePath = ""
      End If
    End function     
    'Public Function debug(varStr)
    ' response.write varStr
    ' response.end
    'End Function
    '类结束
End Class
%>