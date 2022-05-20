<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
call checklogin()

Set rs = conn.Execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'") 
if not (rs.eof and rs.bof) then Purchased=rs("Purchased")
if isnull(Purchased)=true then Purchased=""
rs.close
set rs=Nothing
%>

<!DOCTYPE html>
<head id="Head">
	<title>Purchased Products-<%=ensitename%>-<%=siteurl%></title>
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
<span>CURRENT: </span><a class="breadcrumbs-home" href="javascript:;">ACCOUNT</a><i></i>
<strong>
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;PURCHASED PRODUCTS
</strong>
</div>
<!-- E breadcrumbs -->
</section>
<!-- E page-title -->

<!-- S search -->
<!--#include file="goldweb_search.asp"-->
<!-- E search -->

<%
sqlprod="select * from goldweb_product where ProdId in ("&Purchased&") order by instr ('"&Replace(Purchased,"'","")&"',ProdId) desc"
Set rsprod=Server.CreateObject("ADODB.RecordSet") 
rsprod.open sqlprod,conn,1,1

if rsprod.bof and rsprod.eof then
	response.write "You haven't purchased any product!"
else
	pages=per_page_num '每页显示商品数量
	rsprod.pagesize=pages
	allPages = rsprod.pageCount	'总页数
	page=request("page")
	If not isNumeric(page) then page=1
	if isEmpty(page) or clng(page) < 1 then
		page = 1
	elseif clng(page) >= allPages then
		page = allPages 
	end if
	rsprod.AbsolutePage=page

%> 
<div  class="content-wrapper">
<!-- Start_Module_63268 -->
<div class="module-default">
<div class="module-inner">
<div  class="module-content">
<!-- Start_Module_63268 -->

<!-- S portfolio-list -->
<div class="portfolio-list product-list">

<ul class="column clearfix">
<%
productseqid = 1
Do While Not rsprod.eof	and pages>0
%> 

<li class="col-3-1 not-animated <% If productseqid Mod 3 = 0 Then response.write "last" %>" data-animate="fadeInUp" data-delay="100">
<div class="product-item">
<a href="product.asp?ProdId=<%=rsprod("ProdId")%>">
<div class="portfolio-img"><img src="<%=rsprod("ImgPrev")%>" alt="<%=rsprod("enProdName")%>" /></div>
<div class="portfolio-text"><div class="portfolio-title"><h2><%=rsprod("enProdName")%></h2></div><p class="icon-detail"><span>VIEW DETAILS</span></p></div>
<div class="opacity-overlay"></div>
</a>
</div>

<!-- S price addtocart add to favorites -->
<p style="text-align:left; margin-top:10px; ">
<% 
  If CSng(rsprod("enPriceList")) > 0 And rsprod("AddtoCart") = 1 Then 
	response.write "<span style=""font-size:15px; font-weight:bold; color:#df0032;"">" & rsprod("enPriceUnit") & comma(rsprod("enPriceList")) & "</span>"
%>
<script type="text/javascript">
  if(getCookie('cart').indexOf('<%=rsprod("ProdId")%>')>-1)
  {
    document.write('<input style="margin-left:10px; " type="image" src="../images/incart.png"  onclick="javascript:location.href=\'shoppingcart.asp\';">');
  }
  else
  {
    if('<%=rsprod("AddtoCart")%>'=='1')
	{
	  if(<%=reg%>==1)
	  {
	    if(getCookie('userid')=='')
		{
		  document.write('<input style="margin-left:10px; " type="image" src="../images/addtocart.png"  onclick="javascript:location.href=\'login.asp\';">');
		}
		else
		{
		  document.write('<input style="margin-left:10px; " type="image" src="../images/addtocart.png"  onclick="javascript:location.href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\';">');
		}
	  }
	  else
	  {
	    document.write('<input style="margin-left:10px; " type="image" src="../images/addtocart.png"  onclick="javascript:location.href=\'productorder.asp?prodid=<%=rsprod("ProdId")%>\';">');
	  }
    }
  }
</script>
<%
  End If 
%>
<input style="margin-left:10px; "  type="image" src="../images/addtofavorites.png" onclick="javascript:location.href='my_fav.asp?ProdId=<%=rsprod("ProdId")%>';">
</p>
<!-- E price addtocart add to favorites -->

</li>

<%
pages = pages - 1
rsprod.movenext
productseqid = productseqid + 1
Loop
%>
</ul>
</div>
<!-- E portfolio-list -->

<!-- S pagination -->
<div class="pagination pagination-default">
<%
if page = 1 then
	response.write "<span class=""page-first disabled"">First</span><span class=""page-prev disabled"">Previous</span>"
else
	response.write "<a class=""page-first"" href=""my_purchased.asp?page=1"">First</a><a class=""page-prev"" href=""my_purchased.asp?page="&page-1&""">Previous</a>"
end If

response.write "<span class=""current"">" & page & "/" & allpages & "</span>"

if page = allpages then
response.write "<span class=""page-next disabled"">Next</span><span class=""page-last disabled"">Last</span>"
else
response.write " <a class=""page-next"" href=""my_purchased.asp?page="&page+1&""">Next</a><a class=""page-last"" href=""my_purchased.asp?page="&allpages&""">Last</a>"
end if
%>
</div>
<!-- E pagination -->


<!-- End_Module_63268 -->
</div></div></div></div>
<%
rsprod.close
set rsprod=nothing
end if
%>

</section>
<!-- E content -->


<!-- S sidebar -->
<section class="sidebar float-left">
<!--#include file="goldweb_usertree.asp"-->
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
