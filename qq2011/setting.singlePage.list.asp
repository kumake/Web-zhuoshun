<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","")
	if action<>"" and action="del" then
		''删除信息
		Con.Execute("delete * from Sp_SinglePage where ID="&ItemID&"")
		Con.Execute("delete * from user_Singlepage where singlepageID="&ItemID&"")
		ALert "删除成功","setting.singlePage.list.asp"
	end if	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">单页类别管理</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="85%" align="left">类别名称</TH>
			    <TH width="10%" align="left"></TH>
			  </TR>
			  <%
			  	'''''''''''''''''''''''''''''''''''''''''''''''''''
			  	function ListSingleArray(SingleArray,parentID)
				for arraystep=0 to ubound(SingleArray,2) step 1
					if SingleArray(2,arraystep)=parentID and SingleArray(3,arraystep)<>"关于系统" then
			  %>
				  <TR class="content-td1">
					<TD>&nbsp;</TD>
					<TD><%if parentID<>0 then response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"%><%=SingleArray(1,arraystep)%>&nbsp;<u><%=SingleArray(3,arraystep)%></u></TD>
					<TD><a href="setting.singlePage.edit.asp?ItemID=<%=SingleArray(0,arraystep)%>">修改</a>&nbsp;<a href="javascript:if(confirm('Sp_CMS系统提示:\r\n\n确认删除?删除后不可恢复!'))location.href='?action=del&ItemID=<%=SingleArray(0,arraystep)%>';">删除</a></TD>
				  </TR>
			  <%
						ListSingleArray SingleArray,SingleArray(0,arraystep)
					end if
				next
				end function
			    '''''''''''''''''''''''''''''''''''''''''''''''
				Dim SingleArray:SingleArray = Con.QueryData("select ID,singlePagename,parentID,singlePagename_En from Sp_SinglePage order by id asc")
				'response.Write ubound(SingleArray,1)
				if IsArray(SingleArray) and Ubound(SingleArray,2)<>0 then 
					ListSingleArray SingleArray,0
			    end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
