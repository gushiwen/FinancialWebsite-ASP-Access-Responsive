<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
call checklogin()

Set rs = conn.Execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'") 
if not (rs.eof and rs.bof) then fav=rs("fav")
if isnull(fav)=true then fav=""
rs.close
set rs=Nothing

' 添加、编辑 favorite products 
if request("edit")="ok" then fav=""
buyid = Split(request("prodid"), ", ")
For I=UBound(buyid) To 0 Step -1 '降序显示，升序添加
	if fav="" then
		fav = "'"&buyid(I)&"'"
	ElseIf InStr(fav, buyid(i)) <= 0 Then
		fav = fav & ", '" & buyid(i) &"'"
	End If
Next

if len(fav)>255 then
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('You choose too many favorite products.');"
	response.write "location.href='javascript:history.go(-1)';"		
	response.write "</script>"
	response.end
end If

Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
rs.open sql,conn,3,3
rs("fav")=fav
rs.update
rs.close
set rs=Nothing

' 添加 favorite products 后返回Information页面
if request("prodid")<>"" And request("edit")<>"ok" then 
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
End If 
%>
<!DOCTYPE html>
<head id="Head">
	<title>Favorite Products-<%=ensitename%>-<%=siteurl%></title>
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
<span>CURRENT: </span><a class="breadcrumbs-home" href="javascript:;">ACCOUNT</a><i></i>
<strong>
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;FAVORITE PRODUCTS
</strong>
</div>
<!-- E breadcrumbs -->
</section>
<!-- E page-title -->

<!-- S search -->
<!--#include file="goldweb_search.asp"-->
<!-- E search -->

<%
Set rsprod = conn.Execute("select * from goldweb_product where ProdId in ("&fav&") order by instr ('"&Replace(fav,"'","")&"',ProdId) desc") 

if rsprod.bof and rsprod.eof then
	response.write "You haven't added any favorite product!"
else

%> 
<form action="my_fav.asp" method="POST" name="check">
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
Do While Not rsprod.eof
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
<input style="margin-left:10px; " type="CheckBox" name="ProdId" value="<%=rsprod("ProdId")%>" Checked>
</p>
<!-- E price addtocart add to favorites -->

</li>

<%
rsprod.movenext
productseqid = productseqid + 1
Loop
%>
</ul>
</div>
<!-- E portfolio-list -->

<!-- S update button -->
<div style="text-align:center; margin-bottom:30px;">
<style>
input[type="submit"]{-webkit-appearance:none;appearance:none;outline:none;-webkit-tap-highlight-color:rgba(0,0,0,0);}
</style>
<input type="submit" value="  Update Select Box  " style="font-size:16px; cursor:pointer; width:90%; height:30px; line-height:30px; color:#db3e54; background-color:#fee; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px;"><input type="hidden" name="edit" value="ok">
</div>
<!-- E update button -->


<!-- End_Module_63268 -->
</div></div></div></div>



</form>
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
set conn=Nothing
%>
