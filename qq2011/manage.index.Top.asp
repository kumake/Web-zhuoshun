<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨����ҳ��</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY style="BACKGROUND-COLOR: white">
<DIV id=top>
	<UL id=top-tags>
	  <LI><A onFocus="this.blur()"  onclick="selectTag(this);switchApplication('setting')"  href="javascript:void(0)">ȫ������</A> </LI>
	  <LI><A onFocus="this.blur()" onClick="selectTag(this);switchApplication('admin')" href="javascript:void(0)">�������</A> </LI>
	  <LI><A onFocus="this.blur()" onClick="selectTag(this);switchApplication('model')" href="javascript:void(0)">���ݹ���</A> </LI>
	  <LI class="top-selectTag"><A onFocus="this.blur()" onClick="selectTag(this);switchApplication('help')" href="javascript:void(0)">ϵͳ����</A> </LI>
	</UL>
	<DIV id=logo><!--<IMG src="images/top-logo.gif"> --></DIV>
  <DIV id=top-menu>
		<STRONG><%=CK("username")%></STRONG>
		<select name="userrole" class="select">
		<%
			set Rs = Con.Query("select id,RoleName from Sp_ManageRole where id in ("&CK("userrole")&")")
			if rs.recordcount<>0 then
			do while not rs.eof
			response.Write "<option value="&rs("id")&">"&rs("RoleName")&"</option>"
			rs.movenext
			loop
			end if
		%>
		</select> 
		[<A id="SignOut" target="_top" href="javascript:if(confirm('ȷ���˳�?')) location.href='manage.index.quit.asp';">�˳�</A>,<A href="manage.index.main.asp" target="main">�޸�����</A>,<A href="/" target="_blank">������ҳ</A>] 
	</DIV>
</DIV>

<DIV id=top-nav></DIV>
<SCRIPT type=text/javascript>
function selectTag(selfObj){
	// ������ǩ
	var tag = document.getElementById("top-tags").getElementsByTagName("li");
	var taglength = tag.length;
	for(i=0; i<taglength; i++){
		tag[i].className = "";
	}
	selfObj.parentNode.className = "top-selectTag";
}
function switchApplication(url)
{
    parent.left.location.href = "manage.index.Menu.asp?menutype="+url;
    //parent.main.location.href = url + "_Main.html";
}
</SCRIPT>
</BODY>
</HTML>
