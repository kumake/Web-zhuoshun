<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<!--#include file="../Sp_inc/class_Upload.inc"-->
<%
	if CK("userid")<>"" and CK("username")<>"" and CK("userrole")<>"" then
	else
		response.Write "<script>alert('Sp_CMSϵͳ��ʾ:\r\n\n\n��û�е�½!���ߵ�½�Ự���ݶ�ʧ!');parent.location.href='/admin/index.asp';</script>"
	end if

	Dim action:action = VerificationUrlParam("action","string","")
	Dim obj:obj = VerificationUrlParam("obj","string","")	
	if action<>"" and action="upload" then
		dim upload,file,formName,formPath,iCount,filename,fileExt,path,txtuploadpassword
		set upload=new upload_5xSoft ''�����ϴ�����
		formPath=upload.form("filepath")
		txtuploadpassword=upload.form("txt_uploadpassword")
		if config_isNeedUploadPassword=1 and cstr(txtuploadpassword)<>cstr(config_UploadPassword) then
			alert "�ϴ����벻��ȷ",""
		else
		'''''
		''''����ļ����Ƿ����
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Not (fso.FolderExists(server.MapPath("../upload/"))) Then
		fso.CreateFolder(server.MapPath("../upload/"))
		End If
		 ''��Ŀ¼���(/)
		if right(formPath,1)<>"/" then formPath=formPath&"/" 
		iCount=0
		
		for each formName in upload.file ''�г������ϴ��˵��ļ�
			set file=upload.file(formName)  ''����һ���ļ�����
			if file.filesize<100 then
				response.write "<span style='color:#999999;font-size:12px;'>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
				response.end
			end if
				
			if file.filesize>10000000 then
				response.write "<span style='color:#999999;font-size:12px;'>�ļ���С����������10M��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
				response.end
			end if
			
			fileExt=lcase(right(file.filename,4))
			
			if instr(config_UPloadFileAllowEXT,fileEXT)<=0 then
				response.write "<span style='color:#999999;font-size:12px;'>�ļ���ʽ����ȷ(ֻ����"&replace(config_UPloadFileAllowEXT,"#","")&")��[<a href=# onclick=history.go(-1)>�����ϴ�</a>]</span>"
				response.end
			end if 
			
			randomize
			ranNum=int(90000*rnd)+10000
			filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt
			 
			if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
			'response.Write Server.mappath("upload/")&FileName
			'response.End()
			''''����ļ����Ƿ����
			Dim fso
			Set fso = CreateObject("Scripting.FileSystemObject")
			If Not (fso.FolderExists(server.MapPath("../upload/"))) Then
			fso.CreateFolder(server.MapPath("../upload/"))
			End If
			'''''
			file.SaveAs Server.mappath("../upload/")&FileName   ''�����ļ�
			if config_isRecordUploadFile = 1 then
				''''''д�����ݿ�
				Con.Execute("insert into Sp_UploadFile (username,filename,filetype,filesize,filepath) values ('"&CK("username")&"','"&filename&"','"&fileExt&"','"&file.filesize&"','/upload"&filename&"')")
				''''''
			end if
			if config_IsRelativePath = 0 then
			response.write "<script>parent.document.forms[0]."&obj&".value='/upload"&FileName&"'</script>"
			else
			response.write "<script>parent.document.forms[0]."&obj&".value='upload"&FileName&"'</script>"
			end if
			end if
			iCount=iCount+1
			set file=nothing
		next
		
		set upload=nothing  ''ɾ���˶���
		Htmend "�ļ��ϴ�����!"
	  end if
	end if	

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub HtmEnd(Msg)
 set upload=nothing
 response.write "<span style='color:#999999;font-size:12px;'>�ļ��ϴ��ɹ� [<a href=# onclick=history.go(-1)>�����ϴ�</a>]</span>"
 response.end
end sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<style>
	.input
	{
		BORDER-RIGHT: #b6c2cc 1px solid; PADDING-RIGHT: 2px; BORDER-TOP: #b6c2cc 1px solid; PADDING-LEFT: 2px; PADDING-BOTTOM: 2px; BORDER-LEFT: #b6c2cc 1px solid; PADDING-TOP: 2px; BORDER-BOTTOM: #b6c2cc 1px solid
	}
	.button 
	{
		BORDER-TOP-WIDTH: 0px; FONT-WEIGHT: bold; BORDER-LEFT-WIDTH: 0px; BACKGROUND: url(images/input-bg.gif) #fff repeat-x 50% top; BORDER-BOTTOM-WIDTH: 0px; WIDTH: 85px; COLOR: #fff; LINE-HEIGHT: 24px; HEIGHT: 24px; BORDER-RIGHT-WIDTH: 0px
	}
	body 
	{
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	}
	</style>
</head>
<body>
<div style="width:100%; height:30px; margin:auto;">
<form action="?action=upload&obj=<%=obj%>" enctype="multipart/form-data" method="post">
<input name="filepath" type="file" size="30">&nbsp;
<input name="txt_uploadpassword" type="text" size="6" maxlength="6" class="input">&nbsp;
<input name="Submit" type="submit" class="button" value="�ϴ�">
</form>
</div>
</body>
</html>
