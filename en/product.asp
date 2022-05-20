<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->

<%
'call aspsql()
id=trim(request("prodid"))
if id="" then response.redirect "index.asp"
Set rsprod=conn.execute ("select * from goldweb_product where Online=true and ProdId='"&id&"'")
if (rsprod.bof and rsprod.eof) then
  response.redirect "index.asp"
Else
  Model=rsprod("enModel")
  ModelText=rsprod("enModelText")
  ProdName=rsprod("enProdName")	
  LarCode=rsprod("enLarCode")	
  MidCode=rsprod("enMidCode")	
  ProdId=rsprod("ProdId")
  ImgPrev=rsprod("ImgPrev")
  PriceOrigin=comma(rsprod("enPriceOrigin"))
  PriceList=comma(rsprod("enPriceList"))
  ClickTimes=rsprod("ClickTimes")
  conn.execute "UPDATE goldweb_product SET ClickTimes ="&ClickTimes+1&" WHERE ProdId ='"&id&"'"
end if
%>

<!DOCTYPE html>
<head id="Head">
	<title><%=ProdName%>-<%=ensitename%>-<%=siteurl%></title>
	<meta name="keywords" content="<%=Prodname%>,<%=MidCode%>,<%=LarCode%>,<%=ensitename%>" />
	<meta name="description" content="<%=Prodname%>,<%=ensitename%>"/>

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

	<link rel="stylesheet" type="text/css" href="../css/qhdcontent.css?v=40" />
	<link rel="stylesheet" type="text/css" href="../css/content.css?ver=1.0" />
	<link rel="stylesheet" type="text/css" href="../css/menu.css?ver=1.0" />
	<link rel="stylesheet" type="text/css" href="../css/jquery.fancybox-1.3.4.css?ver=1.0" />
	<link rel="stylesheet" type="text/css" href="../css/pgwslideshow.css?ver=1.0" />
	<link rel="stylesheet" type="text/css" href="../css/animate.min.css?ver=1.0" />
	<link rel="stylesheet" type="text/css" href="../css/style.css?ver=1.2" />
	<link rel="stylesheet" type="text/css" href="../css/style-red.css" />

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

<!-- 外部样式 -->
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
<span>CURRENT: </span><a class="breadcrumbs-home" href="javascript:;">INFORMATION</a><i></i>
<strong>
<% 
  If Trim(LarCode)<>"" Then response.write "&nbsp;&nbsp;&raquo;&nbsp;&nbsp;" & UCase(LarCode)
  If Trim(MidCode)<>"" Then response.write "&nbsp;&nbsp;&raquo;&nbsp;&nbsp;" & UCase(MidCode)
  response.write "&nbsp;&nbsp;&raquo;&nbsp;&nbsp;" & UCase(lleft(Prodname,42))
%>
</strong>
</div>
<!-- E breadcrumbs -->
</section>
<!-- E page-title -->

<!-- S search -->
<!--#include file="goldweb_search.asp"-->
<!-- E search -->

<div  class="content-wrapper">
<!-- Start_Module_63268 -->
<div class="module-default">
<div class="module-inner">
<div  class="module-content">
<!-- Start_Module_63268 -->

<!-- S product-detail-complete -->
<div class="product-detail product-detail-complete">

<!-- S product-intr -->
<div class="product-intr clearfix">

<!-- S product-preview -->
<div class="product-preview">
<!-- S 单张图片 -->
<!-- E 单张图片 -->

<!-- S　图库 -->
<div class="gallery-img-wrap gallery-zoom-img-wrap" data-rel="pgwSlideshow-imgZoom-group-view">
<ul class="pgwSlideshow-gallery-zoom pgwSlideshow clearfix">
  <li><a href="<%= rsprod("ImgPrev")%>"><img src="<%= rsprod("ImgPrev")%>" width="85" height="85" data-large-src="<%= rsprod("ImgPrev")%>" alt="" /></a></li>
<%
Set more_pic=Server.CreateObject("ADODB.Recordset")
sqlmp="select ID,ProdId,FilePath from more_pic where ProdId='"&id&"' order by Num asc"
more_pic.Open sqlmp,conn,1,1
do while not more_pic.eof
%>
  <li><a href="<%= more_pic("FilePath")%>"><img src="<%= more_pic("FilePath")%>" width="85" height="85" data-large-src="<%= more_pic("FilePath")%>" alt="" /></a></li>
<%
  more_pic.movenext
Loop
more_pic.close
set more_pic=Nothing
%>
</ul>
</div>

<script type="text/javascript">
    $(document).ready(function(){
     $('.gallery-zoom-img-wrap .ps-current > ul > li > a').attr("rel", $('.gallery-zoom-img-wrap').attr('data-rel')).append('<span class="icon-zoom">VIEW LARGE</span>');
     //弹窗
     $('.gallery-zoom-img-wrap .ps-current > ul > li > a').fancybox({
      'transitionIn'  : 'elastic',
      'transitionOut'  : 'elastic',
      'hideOnOverlayClick': false,
      'centerOnScroll' : true,
      'padding'    : 0,
      'overlayColor'   : '#000'
     });
    });
</script>
<!-- E　图库 -->

</div>
<!-- E product-preview -->

<!-- S product-info -->
<div class="product-info">

<div class="product-info-item">
<div class="product-name"><h1><%= ProdName%></h1></div>
</div>

<% If Model <> "" Then %>
<div class="product-info-item">
<span><%= ModelText%>: <%= Model%></span>
</div>
<% End If %>

<% If rsprod("enProdDisc") <> "" Then %>
<div class="product-info-item">
<div class="product-summary">
<%=rsprod("enProdDisc")%>
</div>
</div>
<% End If %>

<!-- S price addtocart add to favorites -->
<div class="product-info-item">
<% 
  If CSng(rsprod("enPriceList")) > 0 And rsprod("AddtoCart") = 1 Then 
	response.write "<span style=""font-size:15px; font-weight:bold; color:#df0032;"">" & rsprod("enPriceUnit") & comma(rsprod("enPriceList")) & "</span>"
	If CSng(rsprod("enPriceOrigin")) >0 Then response.write "<span style=""margin-left:10px; font-size:13px; text-decoration:line-through; color:#999999; "">"&rsprod("enPriceUnit")&comma(rsprod("enPriceOrigin"))&"</span>"
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
</div>
<!-- E price addtocart add to favorites -->

<% If ProdId<>"00025" Then %>
<div class="product-info-item">
<span><a href="product.asp?ProdId=00025" style="color:blue; font-size:10px; ">Join trading experts' Wechat group, US$30/month,<br />Get buy & sell notifications, view stock analysis free.</a></span>
</div>
<% End If %>

</div>
<!-- E product-info -->

</div>
<!-- E product-intr -->

<!-- S product-detail-wrapper -->
<div class="product-detail-wrapper">
<div class="product-detail-title"><h3>DETAILS</h3></div>
<div class="product-detail">
<%
If rsprod("enprod1")<>"" And trim(request("KeyChain"))=rsprod("enprod1") Then '加密链接
	response.write rsprod("enMemoSpec")
Else		
	If rsprod("AddtoCart")="0" Then '免费信息
		If rsprod("prod1")="yes" Then '会员尊享
			If request.cookies("goldweb")("userid")<>"" Then
				response.write rsprod("enMemoSpec")
			Else
				response.write "Information below is for free members, you can view it after login."
			End If 
		Else
			response.write rsprod("enMemoSpec")
		End If 
	ElseIf rsprod("AddtoCart")="1" Then '收费信息
			If request.cookies("goldweb")("userid")<>"" Then 
				Purchased=conn.Execute("select Purchased from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")("Purchased")
			End If 
			If InStr(Purchased, trim(request("prodid")))>0 Then '是否购买
				response.write rsprod("enMemoSpec")
			Else
				response.write "Information below is chargeable, you can view it after purchasing."
			End If 
	End If 
End If 
%>
</div>
</div>
<!-- E product-detail-wrapper -->

</div>
<!-- E product-detail-complete -->

<!-- End_Module_63268 -->
</div></div></div></div>

</section>
<!-- E content -->


<!-- S sidebar -->
<section class="sidebar float-left">
<!--#include file="goldweb_proclasstree.asp"-->
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
rsprod.close
set rsprod=Nothing

conn.close
set conn=nothing
%>
