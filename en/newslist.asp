<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
Dim newstitle(5)
Set rsn = conn.Execute("select ennewstitle1,ennewstitle2,ennewstitle3,ennewstitle4,ennewstitle5 from shopsetup") 
newstitle(1)=rsn("ennewstitle1")
newstitle(2)=rsn("ennewstitle2")
newstitle(3)=rsn("ennewstitle3")
newstitle(4)=rsn("ennewstitle4")
newstitle(5)=rsn("ennewstitle5")
rsn.close
set rsn=Nothing

'call aspsql()
Page=request("page")
If not isNumeric(page) then
	response.redirect "newslist.asp"
	response.end
end if

ClassId = Trim(request("ClassId"))
title = "News Center"
sqlnews = "select * from News where online=true and enNewsTitle<>''"

if ClassId<>"" then 
	title = newstitle(CInt(ClassId))
	sqlnews = sqlnews & " and NewsClass='"&ClassId&"'"
End If

sqlnews = sqlnews & " order by uup desc,Pubdate desc"
%>

<!DOCTYPE html>
<head id="Head">
	<title><%=title%>-<%=ensitename%>-<%=SiteUrl%></title>
	<meta name="keywords" content="<%=title%>,<%=ensitename%>" />
	<meta name="description" content="<%=title%>,<%=ensitename%>"/>

	<meta charset="<%=WebPageCharset%>">
	<meta http-equiv="Content-Type" content="text/html; charset=<%=WebPageCharset%>" />
	<meta http-equiv="PAGE-ENTER" content="RevealTrans(Duration=0,Transition=1)" />

	<meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
	<meta name="AUTHOR" content="www.ol.sg" />
	<meta name="RESOURCE-TYPE" content="DOCUMENT" />
	<meta name="DISTRIBUTION" content="GLOBAL" />
	<meta name="ROBOTS" content="INDEX, FOLLOW" />
	<meta name="REVISIT-AFTER" content="1 DAYS" />
	<meta name="RATING" content="GENERAL" />

	<link  rel="stylesheet" type="text/css" href="../css/qhdcontent.css?v=40" />
	<link  rel="stylesheet" type="text/css" href="../css/content.css?ver=1.0" />
	<link  rel="stylesheet" type="text/css" href="../css/menu.css?ver=1.0" />
	<link  rel="stylesheet" type="text/css" href="../css/jquery.fancybox-1.3.4.css?ver=1.0" />
	<link  rel="stylesheet" type="text/css" href="../css/pgwslideshow.css?ver=1.0" />
	<link  rel="stylesheet" type="text/css" href="../css/animate.min.css?ver=1.0" />
	<link  rel="stylesheet" type="text/css" href="../css/style.css?ver=1.2" />
	<link  rel="stylesheet" type="text/css" href="../css/style-red.css" />

	<script src="../js/jquery-1.7.2.min.js?ver=1.0"></script>
	<script src="../js/superfish.js?ver=1.0"></script>
	<script src="../js/jquery.carouFredSel.js?ver=1.0"></script>
	<script src="../js/jquery.touchSwipe.min.js?ver=1.0"></script>
	<script src="../js/jquery.tools.min.js?ver=1.0"></script>
	<script src="../js/jquery.fancybox-1.3.4.pack.js?ver=1.0"></script>
	<script src="../js/pgwslideshow.min.js?ver=1.0"></script>
	<script src="../js/jquery.fixed.js?ver=1.0"></script>
	<script src="../js/cloud-zoom.1.0.2.min.js?ver=1.0"></script>
	<script src="../js/device.min.js?ver=1.0"></script>
	<script src="../js/html5media-1.2.js?ver=1.0"></script>
	<script src="../js/animate.min.js?ver=1.0"></script>
	<script src="../js/custom.js?ver=1.2"></script>
	<script type="text/javascript" src="../js/common.js"></script>
</head>

<body>

<div id="wrapper" class="home-page">
<!-- ==================== S top ==================== -->
<!--#include file="goldweb_top.asp"-->
<!-- ==================== E top ==================== -->


<!-- ==================== S main ==================== -->
<section class="main">
<div class="page-width clearfix">

<!-- S content -->
<section class="content float-right">

<!-- S page-title -->
<section class="page-title page-title-inner clearfix">
<!-- S breadcrumbs -->
<div class="breadcrumbs float-left" >
<span>CURRENT: </span><a class="breadcrumbs-home" href="javascript:;">NEWSLETTER</a><i></i>
<strong>
<%
  If ClassId<>"" Then response.write "&nbsp;&nbsp;&raquo;&nbsp;&nbsp;" & UCase(newstitle(CInt(ClassId)))
%>
</strong>
</div>
<!-- E breadcrumbs -->
</section>
<!-- E page-title -->

<%
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.open sqlnews,conn,1,1
if rsnews.bof and rsnews.eof then
response.write "Coming soon..."
else

pages=12
rsnews.pagesize=per_page_num
allPages = rsnews.pageCount	'总页数
If not isNumeric(page) then page=1
if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rsnews.AbsolutePage=page
%>
<div  class="content-wrapper">
<!-- Start_Module_63255 -->
<div class="module-default">
<div class="module-inner">
<div  class="module-content">
<!-- Start_Module_63255 -->

<!-- S entry-list -->
<div class="entry-list entry-list-article">
<%
Do While Not rsnews.eof	and pages>0
%> 

<!-- S entry-item -->
<div class="entry-item">
<div class="typo">

<!-- S typo-text -->
<div class="typo-text">
<div class="entry-title"><h2><a href="news.asp?NewsId=<%= rsnews("NewsId")%>"><%= lleft(rsnews("enNewsTitle"),50)%></a></h2></div>
<div class="entry-summary" style="height:128px; "><%= rsnews("enNewsContain")%></div>
</div>
<!-- E typo-text -->

</div>
</div>
<!-- E entry-item -->

<%
pages = pages - 1
rsnews.movenext
Loop
%>
</div>
<!-- E entry-list -->

<!-- S pagination -->
<div class="pagination pagination-default">
<%
if page = 1 then
	response.write "<span class=""page-first disabled"">First</span><span class=""page-prev disabled"">Previous</span>"
else
	response.write "<a class=""page-first"" href=""newslist.asp?ClassId="&ClassId&"&page=1"">First</a><a class=""page-prev"" href=""newslist.asp?ClassId="&ClassId&"&page="&page-1&""">Previous</a>"
end If

response.write "<span class=""current"">" & page & "/" & allpages & "</span>"

if page = allpages then
response.write "<span class=""page-next disabled"">Next</span><span class=""page-last disabled"">Last</span>"
else
response.write " <a class=""page-next"" href=""newslist.asp?ClassId="&ClassId&"&page="&page+1&""">Next</a><a class=""page-last"" href=""newslist.asp?ClassId="&ClassId&"&page="&allpages&""">Last</a>"
end if
%>
</div>
<!-- E pagination -->

<!-- End_Module_63255 -->
</div></div></div></div>
<%
rsnews.close
set rsnews=nothing
end if
%>
</section>
<!-- E content -->


<!-- S sidebar -->
<section class="sidebar float-left">
<!--#include file="goldweb_newsclasstree.asp"-->
</section>
<!-- E sidebar -->

</div>
</section>


<!-- ==================== E main ==================== -->


<!-- ==================== S footer ==================== -->
<!--#include file="goldweb_bottom.asp"-->
<!-- ==================== E footer ==================== -->

<!-- ==================== S bottom ==================== -->
<section class="bottom"><div  class="QHDEmptyArea page-width clearfix"></div></section>
<!-- ==================== E bottom ==================== -->

</div>
<!-- end of wrapper -->

<!--#include file="goldweb_fixed.asp"-->
</body>
</html>
<%
conn.close
set conn=nothing
%>
