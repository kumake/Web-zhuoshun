function selectTagByLiIndexMethod(tagsID,tagContentName,showContent,index)
{
    //操作标签
	var tag = document.getElementById(tagsID).getElementsByTagName("li");
	var taglength = tag.length;
	var tagNameTemp ='';
    for(i=0; i<taglength; i++)
    {
	    tag[i].className = "";
	    tagNameTemp = tagContentName+i;
	    if(document.getElementById(tagNameTemp))
	    {
	        document.getElementById(tagNameTemp).style.display = "none";
	    }
    }
    tag[index].className = "selectTag";
	document.getElementById(showContent).style.display = "block";
}

function selectTagMethod(tagsID,tagContentName,showContent,Obj)
{
    //操作标签
	var tag = document.getElementById(tagsID).getElementsByTagName("li");
	var taglength = tag.length;
	var tagNameTemp ='';
    for(i=0; i<taglength; i++)
    {
	    tag[i].className = "";
	    tagNameTemp = tagContentName+i;
	    if(document.getElementById(tagNameTemp))
	    {
	        document.getElementById(tagNameTemp).style.display = "none";
	    }
    }
    Obj.parentNode.className = "selectTag";
	//操作内容
	document.getElementById(showContent).style.display = "block";    
}

//选项卡的切换
function selectTag(showContent,Obj)
{
	selectTagMethod('tags','tagContent',showContent,Obj);

}

//选项卡的切换
function selectTagByLiIndex(showContent,index)
{
	selectTagByLiIndexMethod('tags','tagContent',showContent,index);
    
}
//切换选项卡并传递参数
function selectTagTransferValue(arg,value)
{
    SelectItem(arg,value);
}
//切换选项卡并传递参数
function selectTagTransferArgs(showContent,index,args,values)
{
    //添加编辑时并传递ID
	selectTagByLiIndex(showContent,index)
	
    for(j = 0; j < args.length; j++)
    {
        selectTagTransferValue(args[j],values[j]);
    }
}


//关闭选项卡
function CloseTag(Obj,showContent)
{
    var index =0;
    if(arguments.length==3)
    {
        index=arguments[2];
    }
    //操作标签
	var tag = document.getElementById("tags").getElementsByTagName("li");
	var taglength = tag.length;
	for(i=0; i<taglength; i++)
    {
	    tag[i].className = "";
	    document.getElementById("tagContent"+i).style.display = "none";
    }
    tag[index].className = "selectTag";
    //关闭选项
    Obj.parentNode.style.display = "none";
	document.getElementById(showContent).style.display = "block";
}
