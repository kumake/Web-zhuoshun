<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","0")
	'���ģ������ID
	Dim ModelCategoryID
	'������Ķ�ά����
	Dim ModelCategoryArray
	if ModelID=0 then
		response.Write "ģ�ͽ��ͳ��ִ���!"
	else
		'ȡ��Ӧģ�͵����ݴ�����ݱ�
		Modelarray = Con.QueryRow("select modelCategoryID from Sp_Model where Id="&ModelID&"",0) 
		ModelCategoryID = Modelarray(0)
		'''ȡ���ݴ�Ž�ModelCategoryArray
		ModelCategoryArray = Con.QueryData("select D.ID,D.categoryID,D.categoryname,D.Pathstr,D.ParentID from Sp_dictionaryCategory C,Sp_dictionary D where C.id=D.categoryID and D.categoryID="&ModelCategoryID&" order by D.Depth asc,D.IndexID desc,D.ID asc")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script type="text/javascript" src="scripts/dtree.js" language="javascript"></script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">���������ʾ�б�</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
			<div class="dtree">
				<script type="text/javascript">
					<!--
					//Node(id, pid, name, url, title, target, icon, iconOpen, open) {
					d = new dTree('d');
					d.add(0,-1,'ϵͳ����Ŀ','Model_list.asp?ModelID=<%=ModelID%>','','mainModel');
					<%
					for i=0 to ubound(ModelCategoryArray,2)
					%>
					d.add(<%=ModelCategoryArray(0,i)%>,<%=ModelCategoryArray(4,i)%>,'<%=ModelCategoryArray(2,i)%>','Model_list.asp?ModelID=<%=ModelID%>&categoryID=<%=ModelCategoryArray(3,i)%>','','mainModel','','',true,true);
					<%
					next
					%>
					document.write(d);
					//d.openAll();
					//-->
				</script>
			</div>
	  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
<%
end  if
%>