<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	dim groupArray:groupArray = Con.QueryData("select ID,upgradeNeedScore from Sp_UserGroup order by ID asc")
	Dim useruid,groupID,score,upgradescore
	set rs = Con.Query("select id,groupID,score from sp_user")
	if rs.recordcount<>0 then
	do while not rs.eof
	  useruid = rs("id")
	  groupID = rs("groupID")
	  score = rs("score")
	  ''''
	  for i=0 to ubound(groupArray,2)
	  	if Cint(groupArray(0,i))=Cint(groupID) then
		upgradescore = groupArray(1,i)
		exit for
		end if
	  next
	  ''''
	  if Cint(score)>=cint(upgradescore) then
	  	'''升级
		  for i=0 to ubound(groupArray,2)
			if Cint(groupArray(0,i))>Cint(groupID) then
			groupID = groupArray(0,i)
			exit for
			else
			groupID = -1
			end if
		  next
		  if groupID<>-1 then Con.execute("update Sp_user set groupID="&groupID&" where ID="&useruid&"")
	  end if
	  ''''
	rs.movenext
	loop
	end if
	rs.close
	set rs = nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">会员升级</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
	  	升级成功......
	  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
