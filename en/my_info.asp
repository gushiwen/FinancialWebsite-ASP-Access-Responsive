<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
call checklogin()
if request("edit")="ok" then call edit()
%>

<!DOCTYPE html>
<head id="Head">
	<title>Personal Details-<%=ensitename%>-<%=siteurl%></title>
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
	//判断表单输入正误
	function Checkmodify()
	{
		if (document.myinfo.UserName.value.length < 2 || document.myinfo.UserName.value.length >30) {
			alert("Contact Name should be 2-30 characters.");
			document.myinfo.UserName.focus();
			return false;
		}
		if (document.myinfo.HomePhone.value.length <6 || document.myinfo.HomePhone.value.length >20) {
			alert("Telephone Number should be 6-20 numbers.");
			document.myinfo.HomePhone.focus();
			return false;
		}
		if (document.myinfo.UserMail.value.length <10 || document.myinfo.UserMail.value.length >50) {
			alert("Please input a valid email address.");
			document.myinfo.UserMail.focus();
			return false;
		}
		if (document.myinfo.UserMail.value.length > 0 && !document.myinfo.UserMail.value.match( /^.+@.+$/ ) ) {
			alert("Please input a valid email address.");
			document.myinfo.UserMail.focus();
			return false;
		}
		if (document.myinfo.Address.value.length <3 || document.myinfo.Address.value.length >150) {
			alert("Contact address should be 3-100 characters.");
			document.myinfo.Address.focus();
			return false;
		}
		if (document.myinfo.Country.value.length <3 || document.myinfo.Country.value.length >50) {
			alert("Country should be 3-50 characters.");
			document.myinfo.Address.focus();
			return false;
		}
		if (document.myinfo.ZipCode.value.length <4 || document.myinfo.ZipCode.value.length >12) {
			alert("Post code should be 4-10 characters.");
			document.myinfo.ZipCode.focus();
			return false;
		}
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
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;PERSONAL DETAILS
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

.account_details{ }
.account_details .details_title{background:#d2364c; color:#fff; display:block; width:100%; height:36px; line-height:36px; padding:0px; text-align:center; font-size:14px; font-weight:600; }
.account_details .details_row{width:100%; margin-top:10px; text-align:center; }
.account_details .details_row .input_common{width:90%; height:30px; line-height:30px; text-indent:2px; }
.account_details .details_row .textarea_common{width:90%; height:60px; text-indent:2px; }
.account_details .details_row .btn_submit{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; color:#db3e54; background-color:#fee; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
.account_details .details_row .btn_common{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; color:#888; background-color:#ffe; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
</style>

<%
set rs=conn.execute("select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")
%>

		<!--客户详细资料-->
		<div class="account_details">
		<form name="myinfo" action="my_info.asp" method="post" onSubmit="return Checkmodify();">
			<div class="mobile-section"><div class="details_title">Personal Details</div></div>

			<div class="details_row">
				<input class="input_common" type="text" name="UserName" value="<%=rs("UserName")%>" maxlength="30" placeholder="Full Name Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="HomePhone" value="<%=rs("HomePhone")%>" maxlength="20" placeholder="Mobile Number Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="UserMail" value="<%=rs("UserMail")%>" maxlength="50" placeholder="Email Address Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="Address" value="<%=rs("Address")%>" maxlength="150" placeholder="Detailed Address Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="Country" value="<%=rs("Country")%>" maxlength="50" placeholder="Country Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="ZipCode" value="<%=rs("ZipCode")%>" maxlength="12" placeholder="Post Code Required" autocomplete="off">
			</div>

			<div class="details_row">
				<input class="btn_submit" type="submit" value="Submit Form" name="Submit">
			</div>
			<div class="details_row">
				<input class="btn_common"  type="reset" value="Reset Form" name="Reset">
			</div>
			<input type="hidden" name="edit" value="ok">
		 </form>
		</div>
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
'修改资料
sub edit()
call goldweb_check_path()

'Required
UserName=Trim(request.form("UserName"))
HomePhone=Trim(request.form("HomePhone"))
UserMail=Trim(request.form("UserMail"))
Address=Trim(request.form("Address"))
Country=Trim(request.form("Country"))
ZipCode=Trim(request.form("ZipCode"))

sql = "select * from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
rs("UserName")=UserName
rs("HomePhone")=HomePhone
rs("UserMail")=UserMail
rs("Address")=Address
rs("Country")=Country
rs("ZipCode")=ZipCode
rs.update
rs.close
set rs = Nothing
conn.close
set conn=nothing
response.write "<script language='javascript'>"
response.write "alert('Personal information has been changed.');"
response.write "location.href='my_info.asp';"
response.write "</script>"
response.end
end sub

conn.close
set conn=nothing
%>
