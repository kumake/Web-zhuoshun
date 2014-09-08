<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<!--#include file="../Sp_inc/class_Upload.inc"-->
<%
	if CK("userid")<>"" and CK("username")<>"" and CK("userrole")<>"" then
	else
		response.Write "<script>alert('Sp_CMS系统提示:\r\n\n\n你没有登陆!或者登陆会话数据丢失!');parent.location.href='/admin/index.asp';</script>"
	end if
	Dim action:action = VerificationUrlParam("action","string","")
	Dim fpath:fpath = VerificationUrlParam("fpath","string","")	
	if fpath="" then:response.Write "请将需要上传的文件上传到制定的文件夹!":response.End()
	if action<>"" and action="upload" then
		dim upload,file,formName,formPath,iCount,filename,fileExt,path
		set upload=new upload_5xSoft ''建立上传对象
		formPath=upload.form("filepath")
		 ''在目录后加(/)
		if right(formPath,1)<>"/" then formPath=formPath&"/" 
		iCount=0
		
		for each formName in upload.file ''列出所有上传了的文件
			set file=upload.file(formName)  ''生成一个文件对象
			if file.filesize<100 then
				response.write "<span style='color:#999999;font-size:12px;'>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
				response.end
			end if
				
			if file.filesize>10000000 then
				response.write "<span style='color:#999999;font-size:12px;'>文件大小超过了限制10M　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
				response.end
			end if
			
			fileExt=lcase(right(file.filename,4))
			filename = Mid(file.filename,InStrRev(file.filename, "/")+1)
			
			if instr(config_UPloadFileAllowEXT,fileEXT)<=0 then
				response.write "<span style='color:#999999;font-size:12px;'>文件格式不正确(只允许："&replace(config_UPloadFileAllowEXT,"#","")&")　[<a href=# onclick=history.go(-1)>重新上传</a>]</span>"
				response.end
			end if 
			if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
				'response.Write Server.mappath("../template/"&fpath)&"/"&FileName   ''保存文件
				'response.End()
				''''检查文件夹是否存在
				Dim fso
				Set fso = CreateObject("Scripting.FileSystemObject")
				If Not (fso.FolderExists(server.MapPath("../template/"))) Then
				fso.CreateFolder(server.MapPath("../template/"))
				End If
				file.SaveAs Server.mappath("../template/"&fpath)&"/"&FileName   ''保存文件
			end if
			iCount=iCount+1
			set file=nothing
		next
		
		set upload=nothing  ''删除此对象
		Htmend "文件上传结束!"
		sub HtmEnd(Msg)
		 set upload=nothing
		 response.write "<span style='color:#999999;font-size:12px;'>文件上传成功 [<a href=# onclick=history.go(-1)>继续上传</a>]</span>"
		 response.end
		end sub
	end if	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<TITLE>后台首页</TITLE>
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
<form action="?action=upload&fpath=<%=fpath%>" enctype="multipart/form-data" method="post">
<input name="filepath" type="file" size="30">
<input name="Submit" type="submit" class="button" value="上传">
</form>
</div>
</body>
</html>
