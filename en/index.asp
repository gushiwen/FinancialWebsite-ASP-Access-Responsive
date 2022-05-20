<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->
<!DOCTYPE html>

<head id="Head">
	<title><%=ensitename%>-<%=siteurl%></title>
	<meta name="keywords" content="<%=ensitekeywords%>" />
	<meta name="description" content="<%=ensitedescription%>"/>

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
	<link  rel="stylesheet" type="text/css" href="../css/style-red.css" /><script src="../js/jquery-1.7.2.min.js?ver=1.0"></script>

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

<!-- 外部样式 -->
<body>
<div id="wrapper" class="home-page">

<!-- ==================== S top ==================== -->
<!--#include file="goldweb_top.asp"-->
<!-- ==================== E top ==================== -->


<!-- ==================== S header ==================== -->
<div  class="header">
<!-- Start_Module_63216 -->
<div class="module-default">
<div class="module-inner">
<div  class="module-content">

<!-- Start_Module_63216 -->
<!-- S slideshow -->
<div class="slideshow carousel clearfix" style="height:650px; overflow:hidden;">
<div id="carousel-slideshow">
<div class="carousel-item"><div class="carousel-img"><a href="javascript:;" target=""><img src="<%=pic1%>" height="650" alt="1" /></a></div></div>
<div class="carousel-item"><div class="carousel-img"><a href="javascript:;" target=""><img src="<%=pic2%>" height="650" alt="2" /></a></div></div>
<div class="carousel-item"><div class="carousel-img"><a href="javascript:;" target=""><img src="<%=pic3%>" height="650" alt="3" /></a></div></div>
<div class="carousel-item"><div class="carousel-img"><a href="javascript:;" target=""><img src="<%=pic4%>" height="650" alt="4" /></a></div></div>
</div>
<div class="carousel-btn carousel-btn-fixed" id="carousel-page-slideshow"></div>
</div>

<script type="text/javascript">
 $(window).bind("load resize",function(){
  $("#carousel-slideshow").carouFredSel({
   width       : '100%',
   items  : { visible : 1 },
   auto     : { pauseOnHover: true, timeoutDuration:5000 },
   swipe     : { onTouch:true, onMouse:true },
   pagination  : "#carousel-page-slideshow"
   //prev   : { button:"#carousel-prev-slideshow"},
   //next   : { button:"#carousel-next-slideshow"},
   //scroll   : { fx : "crossfade" }
  });
 });
</script>

<!-- E slideshow -->
<!-- End_Module_63216 -->
</div>
</div>
</div>
</div>
<!-- ==================== E Header ==================== -->


<!-- ==================== S main ==================== -->
<section class="main">
<div  class="full-screen clearfix">

<!-- S module-full-screen products -->
<div class="module-full-screen" style="">
<div class="module-inner not-animated" data-animate="fadeInUp" data-delay="200">
<div class="page-width">

<!-- S search -->
<!--#include file="goldweb_search.asp"-->
<!-- E search -->

<div  class="module-full-screen-content">

<!-- S tabs -->
<div class="tabs-default tabs-center">
<%
  Set rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from goldweb_class where MidSeq=1 order by larseq"
  rs.open sql,conn,1,3
  classseqid = 1
%>

<ul class="tabs-nav clearfix">
<%
  do while not rs.eof
	If classseqid = 1 Then 
		response.write "<li><a href=""javascript:;"" class=""current""><span>" & rs("enLarCode") & "</span></a></li>"
	Else
		response.write "<li><a href=""javascript:;""><span>" & rs("enLarCode") & "</span></a></li>"
	End If 
	rs.movenext
	classseqid = classseqid + 1
  loop
%>
</ul>

<div class="tabs-panes">
<%
  rs.movefirst
  classseqid = 1
  do while not rs.eof
 %>
<div class="tab-box clearfix" style="<% If classseqid=1 Then response.write "display:block;" %>">
<div class="tab-box-content">
<!-- Start_Module_63235 -->
<div class="module-default">
<div class="module-inner">
<div  class="module-content">
<!-- Start_Module_63235 -->
<!-- S portfolio-list -->
<div class="portfolio-list product-list">
<ul class="column clearfix">
 <%
	Set rsm=Server.CreateObject("ADODB.Recordset")
	sqlm="select top 3 * from goldweb_product where enlarcode='"&rs("enLarCode")&"' and online=true order by AddDate desc"
	rsm.open sqlm,conn,1,3
	productseqid = 1
	do while not rsm.eof
%>
<li class="col-3-1 not-animated <% If productseqid Mod 3 = 0 Then response.write "last" %>" data-animate="fadeInUp" data-delay="100"><div class="product-item"><a href="product.asp?ProdId=<%=rsm("ProdId")%>"><div class="portfolio-img"><img src="<%=rsm("ImgPrev")%>" alt="<%=rsm("enProdName")%>" /></div><div class="portfolio-text"><div class="portfolio-title"><h2><%=rsm("enProdName")%></h2></div><p class="icon-detail"><span>VIEW DETAILS</span></p></div><div class="opacity-overlay"></div></a></div></li>
<%
	  rsm.movenext
	  productseqid = productseqid + 1
	loop
	rsm.close
	set rsm=Nothing
%>
</ul>
<ul class="column clearfix"></ul>
</div>
<!-- E portfolio-list -->
<!-- End_Module_63235 -->
</div></div></div></div>
<p class="tab-more tab-more-center"><a href="productlist.asp?LarCode=<%=Server.URLEncode(rs("enLarCode"))%>&ClassId=<%=rs("ClassId")%>">VIEW MORE >></a></p>
</div>
<%
	rs.movenext
	classseqid = classseqid + 1
  loop
%>
</div>

<%
  rs.close
  set rs=Nothing
%>
</div>
<!-- E tabs -->


</div></div></div></div>
<!-- E module-full-screen products -->


<!-- S module-full-screen introduction -->
<div class="module-full-screen module-full-screen-hl module-full-screen-bg-img" style="background-image:url(../images/aboutbg.jpg);">
<div class="module-inner not-animated" data-animate="fadeInDown" data-delay="200">
<div class="page-width">
<div class="module-full-screen-title">
<div class="module-title-content"><i class="mark-left"></i><h2>ABOUT  US</h2><i class="mark-right"></i></div>
</div>
<div  class="module-full-screen-content">
<!-- Start_Module_63218 -->
<div class="qhd-content"><p><%= conn.Execute("select enbody99 from page")("enbody99")%></p></div>
<!-- End_Module_63218 -->
</div>
<div class="module-full-screen-more"><a href="page.asp?id=11">VIEW MORE >></a></div>
</div></div></div>
<!-- E module-full-screen introduction -->


<!-- S module-full-screen company -->

<!-- E module-full-screen company -->


<!-- S module-full-screen news -->
<div class="module-full-screen" >
<div class="module-inner not-animated" data-animate="fadeInUp" data-delay="200">
<div class="page-width">
<div  class="module-full-screen-content">
<!-- Start_Module_63220 -->
<div class="qhd-module">
<div class="column">


<%
  dim newstitle(5)
  Set rs = conn.Execute("select * from shopsetup") 
  newstitle(1)=rs("ennewstitle1")
  newstitle(2)=rs("ennewstitle2")
  newstitle(3)=rs("ennewstitle3")
  newstitle(4)=rs("ennewstitle4")
  newstitle(5)=rs("ennewstitle5")
  rs.close
  set rs=Nothing

  for i=1 to 2

'开始资讯分类显示
    if newstitle(i)<>"" Then
%>
<div class="col-2-1 <% If i Mod 2 = 0 Then response.write "last" %>">
<div  class="qhd_column_contain">
<!-- Start_Module_63233 -->
<div class="module-default">
<div class="module-inner">

<div class="module-title module-title-default clearfix">
<div class="module-more-default float-right"><a href="newslist.asp?ClassId=<%=i%>">More</a></div>
<div class="module-title-content clearfix"><h3 style="background-image:url(../images/title.png);" class="module-icon-default"><%= newstitle(i)%></h3></div>
</div>

<div  class="module-content">
<!-- Start_Module_63233 -->
<!-- S article-list-row -->
<div class="article-list-row headlines-list headlines-set entry-list-mobile">
<ul>
<%
      set rs=conn.execute ("select top 5 * from News where online=true and NewsClass='"&i&"' and enNewsTitle<>'' order by uup desc,Pubdate desc")
	  Do while Not rs.eof
%>
	  <li><a class="article-title" href="news.asp?NewsId=<%=Cstr(rs("NewsId"))%>" title="<%= rs("enNewsTitle")%>"><%= lleft(rs("enNewsTitle"),40)%></a></li>
<%
	    rs.movenext
	  loop
      rs.close
      set rs=Nothing
%>
</ul>
</div>
<!-- E article-list-row -->
<!-- End_Module_63233 -->
</div></div></div></div></div>
<%
    end If
' 结束资讯分类显示

  Next
%>

</div></div>
<!-- End_Module_63220 -->
</div></div></div></div>
<!-- E module-full-screen news -->

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
