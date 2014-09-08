<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
Dim text:text = VerificationUrlParam("text","string","")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" src="http://api.51ditu.com/js/maps.js"></script>
	<script language="javascript" src="http://api.51ditu.com/js/ezmarker.js"></script>
</HEAD>
<BODY>
<form>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">地图标注</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<!--地图隐藏控件//-->
			<input name="map_x" type="hidden" value="">
			<input name="map_y" type="hidden" value="">
			<input name="map_z" type="hidden" value="">
			<!--地图隐藏控件//-->
			<TABLE cellSpacing=1 width="100%">
			<TR class=content-td1>
				<TH scope=col align=left>选择城市 </TH>
			</TR>
			<%
			'set rs = Con.Query("select * from Sp_City where ParentID=0")
			'if rs.recordcount<>0 then
			'do while not rs.eof
			%>
			<TR class=content-td1>
				<TD><input onclick="javascript:setCity();" type="radio" name="city" value="beijing">北京<br></TD>
			</TR>
			<TR class=content-td1>
				<TD><input onclick="javascript:setCity();" type="radio" name="city" value="shanghai">上海<br></TD>
			</TR>
			<TR class=content-td1>
				<TD><input onclick="javascript:setCity();" type="radio" name="city" value="nanjing">南京<br></TD>
			</TR>
			<%
			'rs.movenext
			'loop
			'end if
			%>
			<TR class=content-td1>
				<TD><input name="btnaddmap" disabled onClick="javascript:setmap('<%=text%>');" value="确定已经标注" class="button" type="button"></TD>
			</TR>
			</TABLE>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
<script language="javascript">
	function setmap(ModelText)
	{
		//alert(text+"-"+value);
		var map_x = document.getElementById("map_x").value;
		var map_y = document.getElementById("map_y").value;
		var map_z = document.getElementById("map_z").value;
		parent.document.getElementById(ModelText).value = map_x+"-"+map_y+"-"+map_z;
		parent.ClosePop();
	}
</script>
<script language="JavaScript">
<!--
	var ezmarker = new LTEZMarker("ezmarker");
	function setCity()
	{
		var collection;
		collection = document.all["city"];
		for (i=0;i<collection.length;i++) 
		{
			if (collection[i].checked)
			{
				var c = collection[i].value;
				ezmarker.setDefaultView(c,7);
				break;
			}
		}
	}
	//setMap是ezmarker内部定义的接口，这里可以根据实际需要实现该接口
	function setMap(point,zoom)
	{
		document.getElementById("map_x").value=point.getLongitude();
		document.getElementById("map_y").value=point.getLatitude();
		document.getElementById("map_z").value=zoom;
		///
		document.getElementById("btnaddmap").disabled = false;
	}
	LTEvent.addListener(ezmarker,"mark",setMap);//"mark"是标注事件
//-->
</script>
</form>