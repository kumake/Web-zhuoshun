<%
'  ʹ�÷�����
'  set newExcel = New Export2Excel
'  newExcel.FilePath = "/mail/excel/"----------------------------------·��
'  newExcel.Sql = "select * from user"-------------------------------��ѯ���
'  newExcel.Field = "�ʺ�||����||��������||"----------------------�������
'  response.write newExcel.export2Excel()
'�࿪ʼ
Class Export2Excel
'��������������
    Private strFilePath,strTitle,strSql,strField,strRows,strCols
    Private strCn,strHtml,strPath,strServerPath,Filename,strExcelName
    Private objDbCn,objRs
    Private objXlsApp,objXlsWorkBook,objXlsWorkSheet
    Private arrField
    '��ʼ����
    Private Sub Class_Initialize()
     set objDbCn = Con
     strTitle = "��ѯ���"
     strFilePath="Excel/"
     strRows = 2
     strCols = 1
	 strFileName = ""
    End Sub
    '������
    Private Sub Class_Terminate()
    End Sub
    '����FilePath
    Public Property Let FilePath(value)
     strFilePath = value
     strServerPath=strFilePath
    End Property
    Public Property Get FilePath()
     FilePath = strFilePat
    End Property
    '����Title
    Public Property Let Title(value)
     strTitle = value
    End Property
    Public Property Get Title()
     Title = strTitle
    End Property
    '����Sql
    Public Property Let Sql(value)
     strSql = value
    End Property
    Public Property Get Sql()
     Sql = strSql
    End Property
    '����Field
    Public Property Let Field(value)
     strField = value
    End Property
    Public Property Get Field()
     Field = strField
    End Property
    '����Rows
    Public Property Let Rows(value)
     strRows = value
    End Property
    Public Property Get Rows()
     Rows = strRows
    End Property
    '����Cols
    Public Property Let Cols(value)
     strCols = value
    End Property
    Public Property Get Cols()
     Cols = strCols
    End Property
    'execl�ļ�����
    Public Property Let ExcelName(value)
     strExcelName = value
    End Property
    Public Property Get ExcelName()
     ExcelName = strExcelName
    End Property
    '
    Public Function export2Excel()
     if strSql = "" or strField = "" then
      response.write "�������ô����������Ա��ϵ��лл"
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
      strHtml = "��Ǹ����ʱû���κκ��ʵ����ݵ������������ʣ��������Ա��ϵ��"
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
              If response.IsClientConnected=false then exit for '���ݶർ��ʱ��ܳ���������Ҫ̽���¿ͻ����Ƿ�������
              response.Write "���ڵ�����"&cstr(c+1)&"��<br>"
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
              response.Write "���ڵ�����"&cstr(c+1)&"��<br>"
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
      
      '�ز����٣��������ִ���
      If objFso.fileExists(strFileName)=true then
          objFso.deletefile strFileName
      End if
        response.Write "�����ɹ���<br>"
        response.Flush()      
  
      objXlsWorkSheet.SaveAs strFileName
      
      strHtml = "<script>location.href='" & GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")&strServerpath&Filename  & "';</script>"
      objXlsApp.Quit'��Ҫ
      set objXlsWorkSheet = nothing
      set objXlsWorkBook = nothing
      set objXlsApp = nothing
     end if
     'objDbCn.Close()
     set objRs = nothing
     if err > 0 then
      strHtml = "ϵͳæ,���Ժ�����"
     end if
     export2Excel = strHtml
    End Function
    '����
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
    '�����
End Class
%>