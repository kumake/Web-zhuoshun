<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''��̬��������
''2008-3-14
''Author:һ֧��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
public function FieldUIRs(tablename)
	dim fieldSql:fieldSql = "select * from sys_Field where tablename='"&tablename&"'"
	call openrs(FieldUIRs,fieldSql,3,3)
end function

public function FieldUIType(fieldname,fieldUI,Datasource,defaultValue,FieldIsAllowNull,FieldAttribute)
	'response.Write fieldname&","&fieldUI&","&Datasource&","&defaultValue
	dim datas:datas = split(Datasource,vbcrlf)
	dim tempsel:tempsel=""
	dim tempFieldUIType:tempFieldUIType = ""
	select case fieldUI
		case "textarea"
			FieldUIType = ""
			FieldUIType = "<textarea name='txt_"&fieldname&"' cols='60' rows='5' class='textarea'>"&defaultValue&"</textarea>"
		case "text"
			FieldUIType = "<input name='txt_"&fieldname&"' value='"&defaultValue&"' class='input'>"
		case "html"			
			FieldUIType = "<textarea name='txt_"&fieldname&"' style='display:none;'>"&defaultValue&"</textarea>"&vbCrlf
			FieldUIType = FieldUIType & "<script>"&vbCrlf
			FieldUIType = FieldUIType & "// ʹ��ʾ��"&vbCrlf
			FieldUIType = FieldUIType & "var rnd = Math.round(Math.random()*10000);"
			FieldUIType = FieldUIType & "__ImageRoot = 'images/'; // �༭��JS�ļ�����ڱ�ҳ��·��"&vbCrlf
			FieldUIType = FieldUIType & "__FileDialogPage = 'img.asp?ramd='+rnd;// �ϴ��ļ�ҳ������ڵ�ǰҳ��·��"&vbCrlf
			FieldUIType = FieldUIType & "__ContentID = 'txt_"&fieldname&"'; // ������ȡ��д���HTML�ؼ�ID"&vbCrlf
			FieldUIType = FieldUIType & "new HtmlEditor('txt_"&fieldname&"','500','300','txt_"&fieldname&"');"&vbCrlf
			FieldUIType = FieldUIType & "</script>"&vbCrlf
			'FieldUIType = FieldUIType & "<iframe id='uploadfile' name='uploadfile_"&fieldname&"' width='100%' height='30px;' frameborder='0' scrolling='no' style='padding-top:5px;' src='File_Upload.asp?obj=txt_"&fieldname&"'></iframe>"
		case "eWebeditor"
			FieldUIType = "<textarea name='txt_"&fieldname&"' style='display:none;' cols='56' rows='10' >"&defaultValue&"</textarea>"
			FieldUIType = FieldUIType & "<iframe id='eWebEditor' src='Html/eWebEditor.asp?id=txt_"&fieldname&"&style=s_light' height='300' width='100%' scrolling='no'></iframe>"
		case "datetime"
			FieldUIType = "<input name='txt_"&fieldname&"' value='"&defaultValue&"' onfocus='javascript:HS_setDate(this);'>"
		case "checkbox"
			FieldUIType = "<input type='checkbox' name='txt_"&fieldname&"'"
			if trim(defaultValue)="on" then FieldUIType = FieldUIType +" checked" 
			FieldUIType = FieldUIType +">"
		case "checkboxList"
			for i=0 to UBound(datas)
			  if defaultValue<>"" then
				if instr(defaultValue,datas(i))>0 then
					tempsel=" checked"
				else
					tempsel=""
				end if
				FieldUIType = FieldUIType & "<input type='checkbox' name='txt_"&fieldname&"' value='"&datas(i)&"' "&tempsel&">"&datas(i)
			  else
				FieldUIType = FieldUIType & "<input type='checkbox' name='txt_"&fieldname&"' value='"&datas(i)&"'>"&datas(i)
			  end if
			next				
		case "select"
			FieldUIType = "<select name=""txt_"&fieldname&""">"&vbCrlf
			'FieldUIType = FieldUIType & "<option>"&UBound(datas)&"</option>"
			for i=0 to UBound(datas)
			  'response.Write "<option>1</option>"
			  if defaultValue<>"" then
				if trim(defaultValue)=datas(i) then tempsel=" selected"
				tempFieldUIType = tempFieldUIType & "<option value="""&trim(datas(i))&""" "&tempsel&">"&trim(datas(i))&"</option>"&vbCrlf
			  else
				tempFieldUIType = tempFieldUIType & "<option value="""&trim(datas(i))&""">"&trim(datas(i))&"</option>"&vbCrlf
			  end if
			  tempsel = ""
			next				
			FieldUIType = FieldUIType & tempFieldUIType & " </select>"&vbCrlf	
		case "file"
			FieldUIType = "<input name='txt_"&fieldname&"' class='input' size='60' type='text' value='"&defaultValue&"'>&nbsp;&nbsp;<a href='#' onClick=""javascript:SelectTemplate('pulg.uploadfile.list.asp?obj=txt_"&fieldname&"');"" class='huitext'>���������</a><br>"
			FieldUIType = FieldUIType & "<iframe name='file_"&fieldname&"' frameborder='0'   width='100%' height='30px;' scrolling='no' src='File_Upload.asp?obj=txt_"&fieldname&"'></iframe>"
		case "textaddSelect"
			Dim defaultArray
			if defaultValue<>"" then
				defaultArray = split(defaultValue,"#")
				FieldUIType = "<input name='txt_"&fieldname&"' value='"&defaultArray(0)&"' class='input'>"
			else
				FieldUIType = "<input name='txt_"&fieldname&"' value='"&defaultValue&"' class='input'>"
			end if
			''''''
			FieldUIType = FieldUIType &"&nbsp;<select name='txt_"&fieldname&"_add'>"
			for i=0 to UBound(datas)
			  if defaultValue<>"" then
				if Cstr(defaultArray(1))=Cstr(datas(i)) then tempsel=" selected"
				FieldUIType = FieldUIType & "<option value='"&datas(i)&"' "&tempsel&">"&datas(i)&"</option>"
			  else
				FieldUIType = FieldUIType & "<option value='"&datas(i)&"'>"&datas(i)&"</option>"
			  end if
			  tempsel = ""
			next				
			FieldUIType = FieldUIType & "</select>"
							
		case "SelectToText"
			FieldUIType = "<input name='txt_"&fieldname&"' value='"&defaultValue&"' class='input'>"
			FieldUIType = FieldUIType &"&nbsp;<select name='txt_"&fieldname&"_add' onchange='javascript:"&fieldname&"change(this);'>"
			for i=0 to UBound(datas)
				FieldUIType = FieldUIType & "<option value='"&datas(i)&"'>"&datas(i)&"</option>"
			next				
			FieldUIType = FieldUIType & "</select>"
			FieldUIType = FieldUIType &"<script>function "&fieldname&"change(obj){document.all.txt_"&fieldname&".value = document.all.txt_"&fieldname&".value +' ' + obj.value;}</script>"
		case "Maptext"
			FieldUIType = "<input name='txt_"&fieldname&"' size='35' readonly value='"&defaultValue&"' class='input'>"
			FieldUIType = FieldUIType & "&nbsp;&nbsp;<a href='#' onClick=""javascript:SelectTemplate('pulg.map.asp?text=txt_"&fieldname&"');"" class='huitext'>��ע</a>"
		case else 
			FieldUIType = "<input name='txt_"&fieldname&"' value='"&defaultValue&"' class='input'>"
	end select
	if FieldIsAllowNull = 1 then FieldUIType = FieldUIType & "<span class='huitext'>&nbsp;����</span>"
end function

'''''��ʾҳ����Ϣ
''fieldname  	�ֶ�
''fieldUI		ҳ����ʾ���ֶ�����
''defaultValue	Ĭ��ֵ
public function FieldWebUI(fieldname,fieldUI,defaultValue)
	select case fieldUI
		case "textarea","text","html","eWebeditor","checkbox","select","SelectToText","textaddSelect"
			FieldWebUI = defaultValue
		case "datetime"
			FieldWebUI = year(defaultValue)&"-"&month(defaultValue)&"-"&day(defaultValue)
		case "checkboxList"			
			  if defaultValue<>"" then
			  	for i=0 to UBound(defaultValue)
				FieldWebUI = FieldUIType & defaultValue(i) &vbCrlf
			    next	
			  end if						
		case "file"
			if defaultValue<>"" then
				Dim dotpos:dotpos = InStrRev(defaultValue,".")
				Dim FileExp:FileExp = Mid (defaultValue,dotpos+1)
				if FileExp="gif" or FileExp="bmp" or FileExp="png" or FileExp="jpg" then
					FieldWebUI = "<img src='"&defaultValue&"' align='absmiddle'>"
				else
					FieldWebUI = "<a href='"&defaultValue&"'><img src='images/"&FileExp&".gif' border='0' align='absmiddle'><span class='huitext'>�ļ�����</span></a>"
				end if
			else
				FieldWebUI= ""
			end if
		case "Maptext"
			FieldWebUI = "��ͼ�ؼ�������"
		case else 
			FieldWebUI = defaultValue
	end select
end function

%>