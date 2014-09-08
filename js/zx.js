<!--
var mypath='ZX/',mypaths;
for(i=0;i<document.scripts.length;i++)
  {
    if(document.scripts[i].src.indexOf('js/zx.js?')!='-1')
    {
      mypath=document.scripts[i].src ;
      mypaths=mypath.split('?') ;
      mypath=mypaths[1] ;
      break ;
    }
  }

var qq='244524371';
var tb='南京苯帮';


function WriteQqStr()
{
	document.write('<DIV id=backi style="right:10px; OVERFLOW: visible; POSITION: absolute; TOP: 140px">');
	document.write('<table border="0" cellpadding="0" cellspacing="0" width="78">');
	document.write('<tr><td><a href="javascript:close_float_left();void(0);" title="关闭本浮动条"><IMG src="images/qq_01.gif" border=0></a></td></tr>');
	document.write('<tr><td><A href="http://wpa.qq.com/msgrd?V=1&Uin=291653065&Site=南京浩麒装饰工程有限公司&Menu=no" target=_blank><IMG src="images/qq_02.gif" border=0></A></td></tr>');
		
	document.write('<tr><td><A href="contact.asp" target=_self><IMG src="images/qq_05.gif" border=0></A></td></tr>');
	document.write('</table>');
	document.write('</DIV>');
}

function close_float_left(){backi.style.visibility='hidden';}

lastScrollY=0; 
function heartBeat(){ 
diffY=document.body.scrollTop; 
percent=.1*(diffY-lastScrollY); 
if(percent>0)percent=Math.ceil(percent); 
else percent=Math.floor(percent); 
document.all.backi.style.pixelTop+=percent; 
lastScrollY=lastScrollY+percent; 
} 

if (!document.layers) {WriteQqStr();window.setInterval("heartBeat()",1); }
//-->