// Web HTML Edit Code By nece001 
// Email: nece001@163.com
/* ---------------------------- Utils Area ---------------------------- */

// Get Absolute X
function __GetX(e){
var l=e.offsetLeft;
while(e=e.offsetParent){l+=e.offsetLeft;}
return l;
}

// Get Absolute Y
function __GetY(e){
var t=e.offsetTop;
while(e=e.offsetParent){t+=e.offsetTop;}
return t;
}

// Attach Event
function __AddEventHandler(Target, EventType, Handler){
  if(Target.attachEvent){
    Target.attachEvent("on" + EventType, Handler);
  }
  else if(Target.addEventListener){
    Target.addEventListener(EventType, Handler, false);
  }
  else{
    Target["on" + EventType] = Handler;
  }
}

// On Click/Left Button Down
function __MousLeftButton(e){
  __CloseMenu();
}

// On Context Menu/Right Button Down
function __MousRightButton(e){
  __CloseMenu();
  __ShowContextMenu(e);
  if(document.all){return false;}else{e.preventDefault();}
}

// Mouse Left Button Up
function __MouseUp(e){
  __CloseMenu();
  if(document.all){__CurrentGEditor = e.srcElement.document.parentWindow;}
  else{__CurrentGEditor = e.target.ownerDocument.defaultView;}
}

// On Key Up
function __KeyUp(e){
  __CloseMenu();
}

// On Key Down
function __KeyDown(e){
  __CheckInput(e);
  __CloseMenu();
}

// ToolButton Mouse Over
function  __TBmor(obj){
  obj.style.border = "1px outset";
  if(document.all){
    obj.style.width = "20px";
    obj.style.height = "20px";
  }else{
    obj.style.width = "18px";
    obj.style.height = "18px";
  }
}

// Tool Button Mouse Out
function __TBmot(obj){
  obj.style.border = "none";
  obj.style.width = "20px";
  obj.style.height = "20px";
}

// Hidden Or Show Select Element
function __HiddenSelect()
{
  var selects = document.getElementsByTagName("select");
  if(selects)
  {
    for (var i in selects)
    {
      if(selects[i].tagName && selects[i].tagName.toLowerCase() == "select"){
        selects[i].style.visibility = __FullScreenMode==1 ? "hidden" : "visible";
      }
    }
  }
}

// Simple HTML <BR>=><br />,<p></P>=><br />,<div></div>=><span></span>
function __SimpleHTML(html)
{
  html = html.replace(/<(img [^>]+)>/gi, '<$1 \/>');
  html = html.replace(/<(input [^>]+)>/gi, '<$1 \/>');
  html = html.replace(/<\/?(BR|HR)[^>]*>/gi,'<$1 \/>');
  html = html.replace(/<P[^>]*>/gi,"");
  html = html.replace(/<\/P[^>]*>/gi,"<br />");
  var re = new RegExp("(<DIV)([^>]*>.*?)(<\/DIV>)","gi");
  html = html.replace(re,"<span$2</span>" );
  return html;
}

// GetPageSize From sina.com
function __GetPageSize() {
    var xScroll, yScroll;

    if (window.innerHeight && window.scrollMaxY) {
        xScroll = document.body.scrollWidth;
        yScroll = window.innerHeight + window.scrollMaxY;
    }
    else if (document.body.scrollHeight > document.body.offsetHeight){ // all but Explorer Mac
        xScroll = document.body.scrollWidth;
        yScroll = document.body.scrollHeight;
    }
    else { // Explorer Mac...would also work in Explorer 6 Strict, Mozilla and Safari
        xScroll = document.body.offsetWidth;
        yScroll = document.body.offsetHeight;
    }

    var windowWidth, windowHeight;
    if (self.innerHeight) {    // all except Explorer
        windowWidth = self.innerWidth;
        windowHeight = self.innerHeight;
    }
    else if (document.documentElement && document.documentElement.clientHeight) { // Explorer 6 Strict Mode
        windowWidth = document.documentElement.clientWidth;
        windowHeight = document.documentElement.clientHeight;
    }
    else if (document.body) { // other Explorers
        windowWidth = document.body.clientWidth;
        windowHeight = document.body.clientHeight;
    }

    // for small pages with total height less then height of the viewport
    if(yScroll < windowHeight){
        pageHeight = windowHeight;
    }
    else {
        pageHeight = yScroll;
    }

    // for small pages with total width less then width of the viewport
    if(xScroll < windowWidth){
        pageWidth = windowWidth;
    }
    else {
        pageWidth = xScroll;
    }

    arrayPageSize = new Array(pageWidth,pageHeight,windowWidth,windowHeight)
    return arrayPageSize;
}

// HTML Tags To Lower Case From SAX_Editor
function cleanNodes(s) {
	var htmlPattern = new RegExp('<[ ]*([\\w]+).*?>', 'gi');
	s = s.replace(htmlPattern, function(ref) {
		var cleanStartTag = "";
		var ref = ref.replace('^<[ ]*', '<');
		var ndx = ref.search(/\s/);
		var tagname = ref.substring(0 ,ndx);

		var attributes = ref.substring(ndx, ref.length);

		if (ndx == -1) return ref.toLowerCase();
		cleanStartTag += tagname.toLowerCase();

		var pairs = attributes.match(/[\w]+\s*=\s*("[^"]*"|[^">\s]*)/gi);
		if (pairs) {
			for (var t = 0; t < pairs.length; t++) {
				var pair = pairs[t];
				var ndx = pair.search(/=/);
				var attrname = pair.substring(0, ndx).toLowerCase();
				if (attrname.match(/on\w+/i)) {
					continue;
				}
				var attrval = pair.substring(ndx, pair.length);
				var wellFormed = new RegExp('=[ ]*"[^"]*"', "g");
				if (!wellFormed.test(attrval)) {
					var attrvalPattern = new RegExp('=(.*?)', 'g');
					attrval = attrval.replace(attrvalPattern, '=\"$1');
					attrval += '"';
				}
				var attr = attrname + attrval;
				cleanStartTag += " " + attr;
			}
		}
		cleanStartTag += ">";
		return cleanStartTag;
	});

	s = s.replace(/<\/\s*[\w]*\s*>/g, function (ref) {return ref.toLowerCase();});
	return s;
}

// Kill Word From eWebEditor Not Use 2007-7-14
function __KillWord( html ) {
  // Remove all SPAN tags
  html = html.replace(/<\/?SPAN[^>]*>/gi, "" );
  // Remove Class attributes
  html = html.replace(/<(\w[^>]*) class=([^ |>]*)([^>]*)/gi, "<$1$3") ;
  // Remove Style attributes
  html = html.replace(/<(\w[^>]*) style="([^"]*)"([^>]*)/gi, "<$1$3") ;
  // Remove Lang attributes
  html = html.replace(/<(\w[^>]*) lang=([^ |>]*)([^>]*)/gi, "<$1$3") ;
  // Remove XML elements and declarations
  html = html.replace(/<\\?\?xml[^>]*>/gi, "") ;
  // Remove Tags with XML namespace declarations: <o:p></o:p>
  html = html.replace(/<\/?\w+:[^>]*>/gi, "") ;
  // Replace the &nbsp;
  html = html.replace(/&nbsp;/, " " );
  // Transform <P> to <DIV>
  var re = new RegExp("(<P)([^>]*>.*?)(<\/P>)","gi") ;    // Different because of a IE 5.0 error
  return html = html.replace( re, "<div$2</div>" ) ;
}

function __WinOpen(url,w,h){ // Open New Window
  if(__EditorDialog==null || __EditorDialog.closed){
    __EditorDialog = window.open(url,"","width="+w+",height="+h+",top,left,toolbar=no,menubar=no,scrollbars=no, resizable=no,location=no, status=no");
  }else{
    __EditorDialog.location.href=url;
    __EditorDialog.focus();
  }
}

/* ---------------------------- Data Area ---------------------------- */
var __TBData = {
 'icoWgt' : {'bgp':"-69px 0",'fun':"__Format('Bold')",'tip':"Bold",'class':"toolButton"},
 'icoIta' : {'bgp':"-88px 0",'fun':"__Format('Italic')",'tip':"Italic",'class':"toolButton"},
 'icoUln' : {'bgp':"-108px 0",'fun':"__Format('Underline')",'tip':"Underline",'class':"toolButton"},
 'icoNft' : {'bgp':"-235px 0",'fun':"__Format('RemoveFormat')",'tip':"Remove Format",'class':"toolButton"},
 'icoFnt' : {'bgp':"-26px 0",'fun':"__ShowMenu(this,'__FontNameMenu__')",'tip':"Font Name",'class':"toolButton"},
 'icoFsz' : {'bgp':"-49px 0",'fun':"__ShowMenu(this,'__FontSizeMenu__')",'tip':"Font Size",'class':"toolButton"},
 'icoCol' : {'bgp':"-2px 0",'fun':"__ShowMenu(this,'__FontColorMenu__')",'tip':"Font Color",'class':"toolButton"},
 'line_1' : {'bgp':"-392px 0",'fun':"",'tip':"",'class':"toolLine"},
 'icoAgl' : {'bgp':"-173px 0",'fun':"__Format('justifyleft')",'tip':"Align Left",'class':"toolButton"},
 'icoAgc' : {'bgp':"-194px 0",'fun':"__Format('justifycenter')",'tip':"Align Center",'class':"toolButton"},
 'icoAgr' : {'bgp':"-213px 0",'fun':"__Format('justifyright')",'tip':"Align Right",'class':"toolButton"},
 'icoLul' : {'bgp':"-128px 0",'fun':"__Format('insertunorderedlist')",'tip':"Unorder List",'class':"toolButton"},
 'icoLol' : {'bgp':"-150px 0",'fun':"__Format('insertorderedlist')",'tip':"Order List",'class':"toolButton"},
 'icoLnk' : {'bgp':"-258px 0",'fun':"__CreateLink()",'tip':"Create Link",'class':"toolButton"},
 'icoUlk' : {'bgp':"-283px 0",'fun':"__Format('Unlink');",'tip':"Clean Link",'class':"toolButton"},
 'icoBrd' : {'bgp':"-327px 0",'fun':"__ElementsBorder()",'tip':"Helper Border",'class':"toolButton"},
 'icoImg' : {'bgp':"-306px 0",'fun':"__OpenImageDialog()",'tip':"Insert Image",'class':"toolButton"},
 'line_2' : {'bgp':"-392px 0",'fun':"",'tip':"",'class':"toolLine"},
 'icoHtm' : {'bgp':"-346px 0",'fun':"__CodeArea()",'tip':"Html Code",'class':"toolButton"},
 'icoFul' : {'bgp':"-369px 0",'fun':"__FullScreen()",'tip':"Full Screen",'class':"toolButton"}
};

var __TBFont = {
 'heiti' : {'name':"&#40657;&#20307;",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'songti' : {'name':"&#23435;&#20307;",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'lishu' : {'name':"&#38582;&#20070;",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'kaiti' : {'name':"&#26999;&#20307;",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'youyuan' : {'name':"&#24188;&#22278;",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'Arial' : {'name':"Arial",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'Arial Black' : {'name':"Arial Black",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'Georgia' : {'name':"Georgia",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'Times New Roman' : {'name':"Times New Roman",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'Verdana' : {'name':"Verdana",'fun':"__Format(\"FontName\",this.innerHTML)"},
 'impact' : {'name':"impact",'fun':"__Format(\"FontName\",this.innerHTML)"}
};

var __TBColor = {
 'wine' : {'name':"&#26263;&#32418;&#33394;",'color':"#800000"},
 'purple' : {'name':"&#32043;&#33394;",'color':"#800080"},
 'red' : {'name':"&#32418;&#33394;",'color':"#F00000"},
 'pink' : {'name':"&#40092;&#31881;&#33394;",'color':"#F000F0"},
 'navy' : {'name':"&#28145;&#34013;&#33394;",'color':"#000080"},
 'blue' : {'name':"&#34013;&#33394;",'color':"#0000F0"},
 'lake' : {'name':"&#28246;&#34013;&#33394;",'color':"#00F0F0"},
 'cyan' : {'name':"&#34013;&#32511;&#33394;",'color':"#008080"},
 'green' : {'name':"&#32511;&#33394;",'color':"#008000"},
 'olive' : {'name':"&#27204;&#27012;&#33394;",'color':"#808000"},
 'aqua' : {'name':"&#27973;&#32511;&#33394;",'color':"#00F000"},
 'saff' : {'name':"&#27225;&#40644;&#33394;",'color':"#F0C000"},
 'black' : {'name':"&#40657;&#33394;",'color':"#000"},
 'gray' : {'name':"&#28784;&#33394;",'color':"#808080"},
 'arge' : {'name':"&#38134;&#33394;",'color':"#C0C0C0"},
 'white' : {'name':"&#30333;&#33394;",'color':"#FFF"}
};

var __CMItem = {
  'code':{'name':"&#32534;&#36753;&#36873;&#20013;&#20195;&#30721;",'fun':"__OpenFragmentCodeWindow()"}
};

var __TBMsg = {
 'codeMod':"Now Is Code Mode!\r\nMust Be Back Of View Mode!"
};

var _W,_H,_TH,_IH;
var __Mode = 1;
var __div;
var __iframe;
var __textarea;
var __EditorWindow;
var __FullScreenMode = 0;
var __EditorValue = '';
var __CurrentGEditor;
var __CurrentGMenuId;
var __CurrentRange;
var __EditorDialog;

// OutSide Vars
var __ImageRoot = '';
var __FileDialogPage = "";
var __SummaryID = "";
var __SummaryLen = 50;
var __ContentID = "";


/* ---------------------------- Pante Area ---------------------------- */
// Output The Editor
// @param string $InitValue  The Init_Value of Editor
// @param int/% arguments[1] The Width of Editor
// @param int arguments[2] The Height of Editor
function HtmlEditor(InitValue)
{
  CreateStyle();
  __FontNameMenu();
  __FontSizeMenu();
  __FontColorMenu();
  __ContextMenu();

  _W = 500;
  if(arguments[1])_W = arguments[1];
  _W = (_W.toString().indexOf("%") == -1 ? parseInt(_W)+"px" : _W);

  _H = 300;
  if(arguments[2])_H = parseInt(arguments[2]);
  _TH = 30;
  _IH = (_H - _TH) + "px";
  _H += "px";
  _TH += "px";

  if(document.getElementById(__ContentID))
  {
	  InitValue = document.getElementById(__ContentID).value;
	  //alert(InitValue);
  }
  document.write("<div id='"+arguments[3]+"__EB__'>ÏÔÊ¾±à¼­Æ÷</div>");
  var ebDiv = document.getElementById(arguments[3]+"__EB__");
  __div = ebDiv;
  ebDiv.style.width = _W;
  ebDiv.style.height = _H;
  ebDiv.style.border = "1px solid #CCC";
  ebDiv.style.background = "#CCC";
  ebDiv.innerHTML = CreateToolBar();

  __textarea = document.createElement("textarea");
  //__textarea.id = arguments[3]+"_textarea";
  __textarea.style.display = "none";
  __textarea.style.width = _W;
  __textarea.style.height = _IH;
  __textarea.style.fontSize = "12px";
  __textarea.style.color = "#00F";
  ebDiv.appendChild(__textarea);
  
  __iframe = document.createElement("iframe");
  //__iframe.id = arguments[3]+"_iframe";
  __iframe.style.width = _W;
  __iframe.style.height = _IH;
  __iframe.style.background = "#FFF";
  __iframe.frameBorder = 0;
  __iframe.marginHeight = 0;
  __iframe.marginWidth = 0;
  __iframe.src = "about:blank";
  ebDiv.appendChild(__iframe);

  __EditorWindow = __iframe.contentWindow;
  with(__EditorWindow.document)
  {
    designMode = "On";
    open();
    write("<html><head><title>Editor Area</title><style type=\"text/css\"></style></head><body>&nbsp;</body></html>");
    close();
	body.style.width = "100%";
	body.style.width = "100%";
    body.style.fontSize = "14px";
    body.innerHTML = InitValue;
  }
  
  __AddEventHandler(__EditorWindow.document,"contextmenu", __MousRightButton);
  __AddEventHandler(__EditorWindow.document,"mouseup", __MouseUp);
  __AddEventHandler(__EditorWindow.document,"click", __MousLeftButton);
  __AddEventHandler(__EditorWindow.document,"keydown", __KeyDown);
  __AddEventHandler(__EditorWindow.document,"keyup", __KeyUp);
}

// Create Style
function CreateStyle()
{
  var style = "<style type='text/css'>\
  .toolBar {background:url("+ __ImageRoot +"bg.gif);overflow:hidden;background-repeat:repeat-x;}\
  .toolButton {float:left;margin:4px 2px;width:20px;height:20px;cursor:pointer;background:url("+ __ImageRoot +"mtoolallbg.gif); background-repeat:repeat-x;}\
  .toolLine {float:left;background:url("+ __ImageRoot +"mtoolallbg.gif);width:4px;margin:4px 0; background-repeat:repeat-x;}\
  .toolMenu {position:absolute;display:none;width:100px;font-size:12px;border:1px solid #CCC;background:#FFF;overflow:hidden;z-index:1001;}\
    .toolMenu a {width:100%;display:block;margin:2px;padding:0 2px;text-decoration:none;color:#000;height:18px;line-height:18px;}\
    .toolMenu a:hover {background:#876;color:#FFF;}\
      .toolMenu .egColor {float:left;display:block;margin:4px 0;width:6px;height:6px;border:1px solid #CCC;overflow:hidden;}\
  </style>";
  document.write(style);
}

// Create Tool Bar
function CreateToolBar()
{
    var item;
    var bdDiv = document.createElement("div");
    var tbDiv = document.createElement("div");
    tbDiv.style.height = _TH;
    tbDiv.className = "toolBar";

    for(var i in __TBData)
    {
        item = document.createElement("a");
        item.href="javascript:void(0)";
        item.style.backgroundPosition = __TBData[i]["bgp"];
        item.className = __TBData[i]["class"];
        item.title = __TBData[i]["tip"];
        item.setAttribute("onclick", __TBData[i]["fun"]);
        if(__TBData[i]["class"] != "toolLine")
        {
          item.setAttribute("onmouseover", "__TBmor(this)");
          item.setAttribute("onmouseout", "__TBmot(this)");
        }
        tbDiv.appendChild(item);
    }

    bdDiv.appendChild(tbDiv);
    return bdDiv.innerHTML;
}


/* ---------------------------- Menu Area ---------------------------- */
// Font Name Menu
function __FontNameMenu()
{
  var items = "";
  for(var i in __TBFont)
  {
    items += "<a href='javascript:void(0)' style='font-family:"+ __TBFont[i]["name"] +"' title='"+ __TBFont[i]["name"] +"' onclick='"+ __TBFont[i]["fun"] +"'>"+ __TBFont[i]["name"] +"</a>";
  }
  document.write("<div id='__FontNameMenu__' class='toolMenu'>"+ items +"</div>");
}

// Font Size Menu
function __FontSizeMenu()
{
  var items = "";
  var fz = 0;
  for(var i=1;i<=7;i++)
  { fz = (i*10-i*3) > 10 ? (i*10-i*3) : 12;
    items += "<a href='javascript:void(0)' style='font-size:"+fz+"px;line-height:"+fz+"px;height:"+fz+"px;' title='"+ i +"&#21495;' onclick=\"__Format('FontSize',"+ i +")\">"+ i +"&#21495;</a>";
  }
  document.write("<div id='__FontSizeMenu__' class='toolMenu'>"+ items +"</div>");
}

// Font Color Menu
function __FontColorMenu()
{
  var items = "";
  for(var i in __TBColor)
  {
    items += "<a href='javascript:void(0)' title='"+ __TBColor[i]["name"] +"' onclick=\"__Format('ForeColor','"+ __TBColor[i]["color"] +"')\">";
    items += "<span style='background:"+ __TBColor[i]["color"] +"' class='egColor'>&nbsp;</span>&nbsp;";
    items += __TBColor[i]["name"] + "</a>";
  }
  document.write("<div id='__FontColorMenu__' class='toolMenu'>"+ items +"</div>");
}


// Context Menu
function __ContextMenu()
{
  var items = "";
  for(var i in __CMItem)
  {
    items += "<a href='javascript:void(0)' title='"+ __CMItem[i]["name"] +"' onclick='"+ __CMItem[i]["fun"] +"'>"+ __CMItem[i]["name"] +"</a>";
  }
  document.write("<div id='__ContextMenu__' class='toolMenu'>"+ items +"</div>");
}


/* ---------------------------- Func Area ---------------------------- */
// Return And Get Editor innerHTML, And Return Summary, And Fill Content Element
function __EditGetValue()
{
  __EditorValue = __textarea.value = __SimpleHTML(cleanNodes(__EditorWindow.document.body.innerHTML));
  if(document.getElementById(__SummaryID))
  {
    if(isNaN(__SummaryLen)){__SummaryLen = 50;}
    var obj = document.getElementById(__SummaryID);
	var text = __EditorValue.replace(/<\/?[^>]*>|[\r\n]/gi, "");
	obj.value = text.substr(0, __SummaryLen);
  }
  //alert(__EditorValue);
  if(document.getElementById(__ContentID))
  {
	  document.getElementById(__ContentID).value = __EditorValue;
  }
  return __EditorValue;
}

// Check Input
function __CheckInput(e)
{
  if(e.keyCode==13){
    if(document.all){__PasteCode("<br />");e.returnValue = false;}else{/*e.preventDefault();*/}
  }
}

// Get Current Range
function __GetRange()
{
  with(__EditorWindow)
  {
    focus();
    if(document.all){return document.selection.createRange();}
    else{
      var selection = getSelection();
      return selection.getRangeAt(selection.rangeCount - 1).cloneRange();
    }
  }
}

// Get Selected Contents HTML
function __GetContentHtml()
{
  var oRange = __GetRange();
  if(document.all){return oRange.htmlText;}
  else{
	var contents = oRange.cloneContents();
	var cnt = document.createElement("div");
	cnt.appendChild(contents);
    return cnt.innerHTML;
  }
}

// paset Codes
function __PasteCode(sHTML)
{
  if (!__Mode){alert(__TBMsg['codeMod']);return;}
  var oRange = __GetRange();

  if(document.all){oRange.pasteHTML(sHTML);}
  else{
    var oFrag = oRange.createContextualFragment(sHTML);
	oRange.deleteContents();
	oRange.insertNode(oFrag);
  }
  __CloseMenu();
}

// Execute Command
function __Format(what,opt){
  if (!__Mode){alert(__TBMsg['codeMod']);return;}
  with(__EditorWindow)
  {
      focus();
      if(!opt){document.execCommand(what,false,false);}
      else{
        if(document.all){document.execCommand(what,"",opt);}else{document.execCommand(what,false,opt);}
      }
  }
  __CloseMenu();
}

// Show Menu __FontNameMenu__,__FontSizeMenu__,__FontColorMenu__
function __ShowMenu(obj,menuID)
{
  __CloseMenu();
  var w = document.defaultView ? document.defaultView.getComputedStyle(obj,null).width : obj.currentStyle.width;
  var h = document.defaultView ? document.defaultView.getComputedStyle(obj,null).height : obj.currentStyle.height;
  var x = __GetX(obj) + "px";
  var y = (__GetY(obj) + parseInt(h) - 2) + "px";

  menu = document.getElementById(menuID);
  menu.style.left = x;
  menu.style.top = y;
  menu.style.display = "block";

  __CurrentGMenuId = menuID;
}

// Show Context Menu
function __ShowContextMenu(e)
{
  __CloseMenu();
  var menuID = "__ContextMenu__";
  var x = (__GetX(__iframe) + e.clientX) + "px";
  var y = (__GetY(__iframe) + e.clientY) + "px";

  menu = document.getElementById(menuID);
  menu.style.left = x;
  menu.style.top = y;
  menu.style.display = "block";
  __CurrentGMenuId = menuID;
}

// Hidden Menu
function __CloseMenu()
{
  __EditGetValue();
  if(__CurrentGMenuId)
  {
    document.getElementById(__CurrentGMenuId).style.display = "none";
    __CurrentGMenuId = "";
  }
}

// Show And Hidden Code Area
function __CodeArea()
{
  var selects = document.getElementsByTagName("select");
  if(__textarea.style.display == "none")
  {
    __textarea.style.display = "";
    __iframe.style.display ="none";
    __Mode = 0;
  }
  else
  {
    __EditorWindow.document.body.innerHTML = __textarea.value;
    __textarea.style.display = "none";
    __iframe.style.display = "";
    __Mode = 1;
  }
  __HiddenSelect();
  __CloseMenu();
}

// Open Dialog For Edit Code
function __OpenFragmentCodeWindow()
{
  __CloseMenu();
  var html = __GetContentHtml();
  __WinOpen("about:blank",600,400);
  __EditorDialog.document.open();
  __EditorDialog.document.write("<textarea id='code' style='width:100%;height:90%;'>"+html+"</textarea>");
  __EditorDialog.document.write("<input type='button' value=' OK ' onclick='opener.__PasteCode(document.getElementById(\"code\").value);opener.__EditGetValue();self.close();' />");
  __EditorDialog.document.close();
}

// Show And Hidden Edit->document.body Elements Borders
function __ElementsBorder()
{
    var oCssRules;
    var cssTxt = "span{border:1px dashed #000;} table {border:1px solid #F00;}td {border:1px dashed #CCC;}";

    with(__EditorWindow){oCssRules = document.styleSheets[0].ownerNode || document.styleSheets[0];}
    if(document.all){
      if(oCssRules.cssText == ""){oCssRules.cssText = cssTxt;}
      else{oCssRules.cssText = "";}
    }
    else{
      if(oCssRules.innerHTML == ""){oCssRules.innerHTML = cssTxt;}
      else{oCssRules.innerHTML = "";}
    }
}

// Create Link
function __CreateLink(){
  if (document.all){document.execCommand("CreateLink",true,true);}
  else{
    var sURL = window.prompt("Enter link location (e.g. http://www.cqqzx.com/):", "http://");
    if ((sURL!=null) && (sURL!="http://") && (sURL!="")){__Format("CreateLink", sURL);}
    else{return false;}
  }
  __CloseMenu();
}

// Full Screen
function __FullScreen()
{
  var oHtml = document.getElementsByTagName("html");
  var sl = parseInt(document.body.scrollLeft);
  var st = parseInt(document.body.scrollTop);
  if(__FullScreenMode == 1)
  {
    document.body.style.width = document.body.style.height = "";
	document.body.style.overflow = "";
	oHtml[0].style.overflow = "";
    __div.style.position = "";
	__div.style.zIndex = "";
    __iframe.style.width = __textarea.style.width = __div.style.width = _W;
    __iframe.style.height = __textarea.style.height = _IH;
    __div.style.height = _H;
    __FullScreenMode = 0;
  }
  else
  {
	oHtml[0].style.overflow = "hidden";
	document.body.style.overflow = document.all ? "hidden" : "visible";
	document.body.style.width = document.body.style.height = 0;

    var page = __GetPageSize();
    __div.style.position = "absolute";
    __div.style.zIndex = 1000;
    __div.style.left = sl + "px";
    __div.style.top = st + "px";
    __iframe.style.width = __textarea.style.width = __div.style.width = page[2] + "px";
    __iframe.style.height = __textarea.style.height = __div.style.height = (page[3] - parseInt(_TH)) + "px";
    __FullScreenMode = 1;
  }
  document.body.scrollLeft = sl;
  document.body.scrollTop = st;
  __HiddenSelect();
}

// Open Image Dialog
function __OpenImageDialog()
{
    if(__FileDialogPage!="")
	{
		//__WinOpen(__FileDialogPage,100,400);
		var tempvalue = window.showModalDialog(__FileDialogPage,"","dialogTop:250px,dialogLeft:600px;dialogWidth:500px,DialogHeight:500px,status:no'");
		__Attachement(tempvalue);
	}
    else
	{
      __Attachement(window.prompt("Enter image location (e.g. http://www.xxx.com/):", "http://"));
    }
}

// Insert Attachements
function __Attachement(value)
{
  if(value)
  {
     var code = value.split(".");
     var ext = code[code.length-1].toLowerCase();

      switch(ext)
      {
          case "gif":
              code = "<img src=\""+ value +"\" />";
              break;
          case "jpg":
              code = "<img src=\""+ value +"\" />";
              break;
          case "png":
              code = "<img src=\""+ value +"\" />";
              break;
          case "swf":
              code = "<object classid=\"clsid:D27CDB6E-AE6D-11cf-96B8-444553540000\">";
              code += "<param name=\"movie\" value=\""+ value +"\" />";
              code += "<param name=\"quality\" value=\"high\" />";
              code += "</object>";
              break;
          case "rm":
              code = "<object classid=\"clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA\" width=\"100\" height=\"50\">";
              code += "<param name=\"src\" value=\""+ value +"\" />";
              code += "<param name=\"autostart\" value=\"-1\" />";
              code += "</object>";
              break;
          case "mp3":
              code = "<embed src=\""+ value +"\" autostart=\"true\"></embed>";
              break;
          case "wma":
              code = "<embed src=\""+ value +"\" autostart=\"true\"></embed>";
              break;
          case "wmv":
              code = "<embed src=\""+ value +"\" autostart=\"true\"></embed>";
              break;
          default:
              code = "<a href=\""+ value +"\">"+ value +"</a>";
      }
      __PasteCode(code);
  }
}