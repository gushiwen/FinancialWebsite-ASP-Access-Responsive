<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
NewsId=Request("NewsId")
If not isNumeric(NewsId) Then
	conn.close
	set conn=nothing
	response.redirect "newslist.asp"
	response.end
end If

set rs=Server.CreateObject("ADODB.RecordSet")
sql= "select * from News where online=true and enNewsTitle<>'' and NewsId="&request("NewsId")
rs.open sql,conn,1,1
if rs.bof and rs.eof Then
	rs.close
	set rs=Nothing
	conn.close
	set conn=nothing
	response.redirect "newslist.asp"
	response.end
Else
	Title= rs("enNewsTitle")
	NewsContain= rs("enNewsContain")
	NewsClass= rs("enNewsClass")
	NewsSource= rs("enSource")
	PubDate= Year(rs("PubDate")) & "-" & Month(rs("PubDate")) & "-" & Day(rs("PubDate"))
	Cktimes= rs("cktimes")
	conn.execute "UPDATE news SET ckTimes ="&Cktimes+1&" WHERE NewsId="&NewsId
	rs.close
	set rs=Nothing
end If

Dim newstitle(5)
Set rsn = conn.Execute("select ennewstitle1,ennewstitle2,ennewstitle3,ennewstitle4,ennewstitle5 from shopsetup") 
newstitle(1)=rsn("ennewstitle1")
newstitle(2)=rsn("ennewstitle2")
newstitle(3)=rsn("ennewstitle3")
newstitle(4)=rsn("ennewstitle4")
newstitle(5)=rsn("ennewstitle5")
rsn.close
set rsn=nothing
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

<div  class="content-wrapper">
<div class="module-default">
<div class="module-inner">
<div  class="module-content">

<!-- S article-detail -->
<div class="article-detail">
<div class="article-title"><h1><%= Title%></h1></div>
<div class="entry-meta"><span><strong>Categoty: </strong><a href="javascript:;"><%=newstitle(NewsClass)%></a></span><span><strong>Published: </strong><strong><%= PubDate%></strong></span><span><strong></strong></span></div>

<div class="article-content-wrapper">
<div class="article-content">
<%= NewsContain%>
</div>
</div>

</div>
<!-- E article-detail -->

</div></div></div></div>

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
