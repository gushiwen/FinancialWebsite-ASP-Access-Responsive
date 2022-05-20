<%
if kaiguan=0 then
response.write guanbi
response.end 
end if

Set ggrs = conn.Execute("select * from adv") 
pic1= ggrs("pic1")
pic1_lnk= ggrs("pic1_lnk")
tit1= ggrs("tit1")
pic2= ggrs("pic2")
pic2_lnk= ggrs("pic2_lnk")
tit2= ggrs("tit2")
pic3= ggrs("pic3")
pic3_lnk= ggrs("pic3_lnk")
tit3= ggrs("tit3")
pic4= ggrs("pic4")
pic4_lnk=ggrs("pic4_lnk")
tit4= ggrs("tit4")
flash=ggrs("flash")
flashwidth=ggrs("flashwidth")
flashheight=ggrs("flashheight")
flashurl=ggrs("flashurl")
flash2=ggrs("flash2")
flash2width=ggrs("flash2width")
flash2height=ggrs("flash2height")
flash2url=ggrs("flash2url")
bannerxiaoguo=ggrs("bannerxiaoguo")
bannerxiaoguo_top=ggrs("bannerxiaoguo_top")
logo=ggrs("logo")
hfpic= ggrs("hfpic")
hfurl= ggrs("hfurl")
hftit= ggrs("hftit")
hf2pic= ggrs("hf2pic")
hf2url= ggrs("hf2url")
hf2tit= ggrs("hf2tit")
piaofu=ggrs("piaofu")
piaofuurl=ggrs("piaofuurl")
piaofupic=ggrs("piaofupic")
piaofutit=ggrs("piaofutit")
tanchu=ggrs("tanchu")
tanchu_time=ggrs("tanchu_time")
tanurl=ggrs("tanurl")
tanheight=ggrs("tanheight")
tanwidth=ggrs("tanwidth")
tantop=ggrs("tantop")
tanleft=ggrs("tanleft")
cebian=ggrs("cebian")
lefturl=ggrs("lefturl")
righturl=ggrs("righturl")
leftpic=ggrs("leftpic")
rightpic=ggrs("rightpic")

iadv1=ggrs("iadv1")
iadv2=ggrs("iadv2")
iadv3=ggrs("iadv3")
iadv4=ggrs("iadv4")
iadv5=ggrs("iadv5")
iadv6=ggrs("iadv6")
iadv7=ggrs("iadv7")
iadv8=ggrs("iadv8")
iadvurl1=ggrs("iadvurl1")
iadvurl2=ggrs("iadvurl2")
iadvurl3=ggrs("iadvurl3")
iadvurl4=ggrs("iadvurl4")
iadvurl5=ggrs("iadvurl5")
iadvurl6=ggrs("iadvurl6")
iadvurl7=ggrs("iadvurl7")
iadvurl8=ggrs("iadvurl8")
ggrs.close
set ggrs=nothing
%>

<header class="top header-v2 desktops-section default-top">
<!-- top-bar --><!-- top-bar -->

<!-- S top-main -->
<div class="top-main">
<div class="page-width clearfix">

<div class="logo float-left" ><a href="index.asp"><img src="<%=logo%>" alt="<%=ensitename%>" /></a></div>

<!-- S language -->
<div class="language">
  <a href="shoppingcart.asp" class="first-level">Shopping Cart</a> <span style="color:#fff; ">|</span> <a href="my_fav.asp" class="first-level">Favorite Info</a> <span style="color:#fff; ">|</span> 
<script type="text/javascript">
// JS checking login and JS write page
if(getCookie("userid")=="")
{
  document.write('<a href="login.asp" class="first-level">Log In</a> <span style="color:#fff; ">|</span> <a href="reg_member.asp" class="first-level">Register</a> <span style="color:#fff; ">|</span> ');
}
else
{
  document.write('<a href="user_center.asp" class="first-level">My Account</a> <span style="color:#fff; ">|</span> <a href="logout.asp" class="first-level">Log Out</a> <span style="color:#fff; ">|</span> ');
}
</script>
  <a href="../ch/index.asp" class="first-level">中文</a>
</div>
<!-- E language -->

</div>
</div>
<!-- E top-main -->

<div class="clear"></div>

<div class="nav-wrapper">
<div class="page-width clearfix">
<!-- S nav -->
<nav class="nav"><div class="main-nav clearfix" >
<ul class="sf-menu">	

	<li <%If InStr(request.servervariables("PATH_INFO"),"index.asp")>0 Then response.write " class=""current"""%> ><a class="first-level" href="index.asp"><strong>HOME</strong></a></li>
	
	<li <%If InStr(request.servervariables("PATH_INFO"),"newslist.asp")>0 Or InStr(request.servervariables("PATH_INFO"),"news.asp")>0 Then response.write " class=""current"""%> ><a class="first-level" href="newslist.asp" ><strong>NEWS</strong></a></li>

<%
	Dim LarCodeStr
	If InStr(request.servervariables("PATH_INFO"),"productlist.asp")>0 Then
		LarCodeStr = request.servervariables("QUERY_STRING")
	ElseIf InStr(request.servervariables("PATH_INFO"),"product.asp")>0 Then
		LarCodeStr = Server.URLEncode(conn.execute("Select enLarCode from goldweb_product where ProdId='"&Trim(Split(request.servervariables("QUERY_STRING"),"=")(1))&"'")("enLarCode"))
	End If
						  
	Set rsL=Server.CreateObject("ADODB.Recordset")
	sqlL="select ClassId,enLarCode from goldweb_class where MidSeq=1 order by larseq"
	rsL.open sqlL,conn,1,3
	do while not rsL.eof
%>
		<li <%If InStr(LarCodeStr,Server.URLEncode(rsL("enLarCode")))>0 Then response.write " class=""current"""%> ><a class="first-level" href="productlist.asp?LarCode=<%=Server.URLEncode(rsL("enLarCode"))%>&ClassId=<%=rsL("ClassId")%>" ><strong><%=UCase(rsL("enLarCode"))%></strong></a></li>
<%
		rsL.movenext
	loop
	rsL.close
	set rsL=Nothing		 
%>

</ul>
</div></nav>
<!-- E nav-->
</div>
</div>

</header>
	
<!-- S touch-top-wrapper -->
<div class="touch-top mobile-section clearfix">

<div class="touch-top-wrapper clearfix">
<div class="touch-logo" ><a  href="index.asp"><img src="<%=logo%>" alt="<%=ensitename%>" /></a></div>
<!-- S touch-navigation -->
<div class="touch-navigation">
<div class="touch-toggle">
<ul>
	<li class="touch-toggle-item-first" ><a href="javascript:;" class="drawer-search" data-drawer="drawer-section-search"><i class="touch-icon-search"></i><span>Search</span></a></li>
	<li class="touch-toggle-item-first" ><a href="shoppingcart.asp" class="drawer-cart" data-drawer="drawer-section-cart"><i class="touch-icon-cart"></i><span>Cart</span></a></li>
	<li class="touch-toggle-item-first" ><a href="javascript:;" class="drawer-user" data-drawer="drawer-section-user"><i class="touch-icon-user"></i><span>User</span></a></li>
	<li class="touch-toggle-item-last"><a href="javascript:;" class="drawer-menu" data-drawer="drawer-section-menu"><i class="touch-icon-menu"></i><span>Menu</span></a></li>
</ul>
</div>
</div>
<!-- E touch-navigation -->
</div>

<!-- S touch-top -->
<div class="touch-toggle-content touch-top-home">

<div class="drawer-section drawer-section-user">
<div class="touch-menu" >
<ul>
<script type="text/javascript">
// JS checking login and JS write page
if(getCookie("userid")=="")
{
  document.write('<li ><a href="login.asp"><span>Log In</span></a></li><li ><a href="reg_member.asp"><span>Register</span></a></li>');
}
else
{
  document.write('<li ><a href="user_center.asp"><span>Account Overview</span></a></li><li ><a href="my_order.asp"><span>Order Management</span></a></li><li ><a href="my_purchased.asp"><span>Pruchased Products</span></a></li><li ><a href="my_fav.asp"><span>Favorite Products</span></a></li><li ><a href="my_info.asp"><span>Personal Information</span></a></li><li ><a href="my_psw.asp"><span>Password Setting</span></a></li><li ><a href="logout.asp"><span>Log Out</span></a></li>');
}
</script>
</ul>
</div>
</div>

<div class="drawer-section drawer-section-search">
	<form name="mobilesearch" action="productsearch.asp" method="post">
		<input name="action" type="hidden" value="topsearch">
		<input name="keywords" value="<%=keywords%>" type="text" placeholder="Search Information" autocomplete="off" style="padding-left:4px; width:80%; height:42px; line-height:42px; BORDER:#CCCCCF 1px solid; COLOR: #666666;">
		<a href="javascript:;" onclick="document.mobilesearch.submit();" style="display:inline-block; width:10%; padding:0 10px; line-height:42px; background-color:#e5e5e5; font-size:16px; font-size:1.6rem; color:#808080; text-decoration:none; text-align:center; cursor:pointer; transition:all 0.5s ease 0s; border-radius:2px; " onMouseOver="this.style.backgroundColor='#d9d9d9';" onMouseOut="this.style.backgroundColor='#e5e5e5';" >GO</a>
	</form>
</div>

<div class="drawer-section drawer-section-menu">
<div class="touch-menu" >
<ul>
	<li><a href="../ch/index.asp"><span>中文</span></a></li>
	<li><a href="index.asp"><span>HOME</span></a></li>
	<li><a href="newslist.asp"><span>NEWS</span></a></li>

<%
Set rsL=Server.CreateObject("ADODB.Recordset")
sqlL="select ClassId,enLarCode from goldweb_class where MidSeq=1 order by larseq"
rsL.open sqlL,conn,1,3
do while not rsL.eof
%>
	<li>
	<a href="productlist.asp?LarCode=<%=Server.URLEncode(rsL("enLarCode"))%>&ClassId=<%=rsL("ClassId")%>"><span><%=UCase(rsL("enLarCode"))%></span></a>
	</li>
<%
	rsL.movenext
loop
rsL.close
set rsL=Nothing		 
%>

</ul>
</div>
</div>

<script type="text/javascript">

    $(document).ready(function(){

     

     $(".touch-toggle a").click(function(event){

      var className = $(this).attr("data-drawer");

      

      if( $("."+className).css('display') == 'none' ){      

       $("."+className).slideDown().siblings(".drawer-section").slideUp();

      }else{

       $(".drawer-section").slideUp(); 

      }

      event.stopPropagation();

     });

     

     /*$(document).click(function(){

      $(".drawer-section").slideUp();     

     })*/

     

     $('.touch-menu a').click(function(){     

      if( $(this).next().is('ul') ){

       if( $(this).next('ul').css('display') == 'none' ){

        $(this).next('ul').slideDown();

        $(this).find('i').attr("class","touch-arrow-up");     

       }else{

        $(this).next('ul').slideUp();

        $(this).next('ul').find('ul').slideUp();

        $(this).find('i').attr("class","touch-arrow-down");

       }   

      }

     });

    });

</script>

</div>
<!-- E touch-top -->

</div>
<!-- E touch-top-wrapper -->