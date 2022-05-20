<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="chopchar.asp"-->
<%
'call aspsql()
call checklogin()
if request("edit")="ok" then call edit()
%>

<!DOCTYPE html>
<head id="Head">
	<title>Password Setting-<%=ensitename%>-<%=siteurl%></title>
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
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;PASSWORD SETTING
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

<%
set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
%>

		<div class="account_details">
		<form name="myinfo" action="my_psw.asp" method="post">
			<div class="mobile-section"><div class="details_title">Password Setting</div></div>
			<input type="hidden" name="userid" value="<%=rs("UserID")%>">
			<div class="details_row">
				<input class="input_common" type="password" name="oldpassword" maxlength="16" placeholder="Old Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw1" maxlength="16" placeholder="New Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw2" maxlength="16" placeholder="Repeat Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" name="Submit" value="Submit Form" >
			</div>
			<div class="details_row">
				<input class="btn_submit" type="button" value="Retrieve password setting" onclick="document.location.href='my_psw_set.asp';" >
			</div>
			<div class="details_row">
				<input class="btn_common" type="reset" name="Reset" value="Reset Form" >
			</div>
			<input type="hidden" name="edit" value="ok">
        </form>
		</div>

<%
rs.close
set rs=nothing
%>
<!-- E form -->

</div></div></div></div>

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
sub edit()
call goldweb_check_path()
oldpassword=trim(request("oldpassword"))
Pw1=trim(request("pw1"))
Pw2=trim(request("pw2"))

if oldpassword="" or Pw1="" or pw2=""  then 
response.write "<script language='javascript'>"
response.write "alert('Please fill in the password.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if	

if Pw1<>pw2 then 
response.write "<script language='javascript'>"
response.write "alert('You have input two different passwords.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if llen(pw1)<6 then 
response.write "<script language='javascript'>"
response.write "alert('The password should be at least 6 characters.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
if rs("userpassword")<>md5(oldpassword) then 
response.write "<script language='javascript'>"
response.write "alert('The old password is incorrect.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if	

if ucase(request.cookies("goldweb")("userid"))<>ucase(request.form("userid")) then
response.write "<script language='javascript'>"
response.write "alert('You are not authorized to do this.');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

sql = "select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
rs("userpassword")=md5(pw1)
rs.update
rs.close
set rs=nothing

response.write "<script language='javascript'>"
response.write "alert('Your password has been changed.');"
response.write "location.href='my_psw.asp';"
response.write "</script>"
response.end
end sub

conn.close
set conn=nothing
%>
