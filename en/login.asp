<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
If request.cookies("goldweb")("userid")<>"" Then 
	conn.close
	set conn=Nothing
	response.write "<meta http-equiv='refresh' content='0;URL=user_center.asp'>"
End If 

randomize
yzm=Int(8999*rnd()+1000)
yzm_a=Int(yzm/1000)
yzm_b=Int((yzm-yzm_a*1000)/100)
yzm_c=Int((yzm-yzm_a*1000-yzm_b *100)/10)
yzm_d=Int(yzm-yzm_a*1000-yzm_b *100-yzm_c*10)
%>

<!DOCTYPE html>
<head id="Head">
	<title>Member Login-<%=ensitename%>-<%=siteurl%></title>
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

	<script type="text/javascript">
	// 个人用户登陆
	function  CheckPLoginForm() 
	{
		if(document.ploginform.userid.value=="")
		{
			alert("Please input your account name!");
			return false;
		}
		if(document.ploginform.password.value=="") 
		{
			alert("Please input password!");
			return false; 
		}	
		if(document.ploginform.verifycode.value != "<%=yzm%>")
		{
			alert("Please input correct verify code!");
			return false;
		}

		document.ploginform.submit();
	}
	</script>
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
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;MEMBER LOGIN
</strong>
</div>
<!-- E breadcrumbs -->
</section>
<!-- E page-title -->

<div  class="content-wrapper">
<div class="module-default">
<div class="module-inner">
<div  class="module-content">

<!-- S form -->
<style>
input[type="text"]{-webkit-appearance:none;appearance:none;outline:none;-webkit-tap-highlight-color:rgba(0,0,0,0);}
input[type="button"]{-webkit-appearance:none;appearance:none;outline:none;-webkit-tap-highlight-color:rgba(0,0,0,0);}
input[type="reset"]{-webkit-appearance:none;appearance:none;outline:none;-webkit-tap-highlight-color:rgba(0,0,0,0);}
input[type="submit"]{-webkit-appearance:none;appearance:none;outline:none;-webkit-tap-highlight-color:rgba(0,0,0,0);}

.account_details{max-width:480px; margin:auto; }
.account_details .details_title{background:#d2364c; color:#fff; display:block; width:100%; height:36px; line-height:36px; padding:0px; text-align:center; font-size:14px; font-weight:600; }
.account_details .details_row{width:100%; margin-top:10px; text-align:center; }
.account_details .details_row .input_common{width:90%; height:30px; line-height:30px; text-indent:2px; }
.account_details .details_row .input_verify{width:59%; height:30px; line-height:30px; text-indent:2px; }
.account_details .details_row .verifycode{width:30%; display:inline-block; height:30px; line-height:30px; }
.account_details .details_row .btn_submit{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; color:#db3e54; background-color:#fee; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
.account_details .details_row .btn_common{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; color:#888; background-color:#ffe; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
</style>

		<div class="account_details">
		<form name="ploginform" action="indexlogin.asp" method="post" onSubmit="return CheckPLoginForm();">
			<div class="mobile-section"><div class="details_title">Member Login</div></div>
			<div class="details_row">
				<input class="input_common" type="text"  name="userid" maxlength="16" placeholder="Account ID" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="password" maxlength="16" placeholder="Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_verify" type="text" name="verifycode" maxlength="12" placeholder="Verify Code" autocomplete="off">
				<div class="verifycode"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_a%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_b%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_c%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_d%>.gif"></div>
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" value="Log in">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Forgot Password" onclick="document.location.href='getpsw.asp';">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Member Register" onclick="document.location.href='reg_member.asp';">
			</div>
			<input type="hidden" name="login" value="ok">
			<input type="hidden" name="cook" value="0">
			<input type="hidden" name="referer" value="<%=Request.ServerVariables("HTTP_REFERER")%>">
        </form>
		</div>
<!-- E form -->

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
conn.close
set conn=nothing
%>
