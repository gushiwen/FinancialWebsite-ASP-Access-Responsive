<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->
<%
randomize
yzm=Int(8999*rnd()+1000)
yzm_a=Int(yzm/1000)
yzm_b=Int((yzm-yzm_a*1000)/100)
yzm_c=Int((yzm-yzm_a*1000-yzm_b *100)/10)
yzm_d=Int(yzm-yzm_a*1000-yzm_b *100-yzm_c*10)
%>

<!DOCTYPE html>
<head id="Head">
	<title>Member Register-<%=ensitename%>-<%=siteurl%></title>
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

		<script language="JavaScript">
			function JM_sh(ob){
				if (ob.style.display=="none"){ob.style.display=""}else{ob.style.display="none"};
			}

			function fucPWDchk(str) 
			{ 
				var strSource ="0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"; 
				var ch; 
				var i; 
				var temp; 

				for (i=0;i<=(str.length-1);i++) 
				{ 
					ch = str.charAt(i); 
					temp = strSource.indexOf(ch); 
					if (temp==-1) 
					{ 
						return 0; 
					} 
				} 
				if (strSource.indexOf(ch)==-1) 
				{ 
					return 0; 
				} 
				else 
				{ 
					return 1; 
				} 
			} 
			function jtrim(str) 
			{ 
				while (str.charAt(0)==" ") 
				{str=str.substr(1);} 
				while (str.charAt(str.length-1)==" ") 
				{str=str.substr(0,str.length-1);} 
				return(str); 
			} 

			//判断表单输入正误
			function Checkreg()
			{
				if (document.ADDUser.UserId.value.length < 4 || document.ADDUser.UserId.value.length >16) {
					alert("Account ID should be 4-16 characters.");
					document.ADDUser.UserId.focus();
					return false;
			}
			if (document.ADDUser.pw1.value.length <6 || document.ADDUser.pw1.value.length >16) {
				alert("Password should be 6-16 characters.");
				document.ADDUser.pw1.focus();
				return false;
			}
			if (!fucPWDchk(document.ADDUser.pw1.value)){
				alert("Password should be numbers, capital or small letters.");
				document.ADDUser.pw1.focus();
				return false;
			}
			if (document.ADDUser.pw1.value != document.ADDUser.pw2.value) {
				alert("You have entered  two different passwords.");
				document.ADDUser.pw2.focus();
				return false;
			}
			if (document.ADDUser.UserMail.value.length <10 || document.ADDUser.UserMail.value.length >50) {
				alert("Please input a valid email address.");
				document.ADDUser.UserMail.focus();
				return false;
			}
			if (document.ADDUser.UserMail.value.length > 0 && !document.ADDUser.UserMail.value.match( /^.+@.+$/ ) ) {
				alert("Please input a valid email address.");
				document.ADDUser.UserMail.focus();
				return false;
			}
			if (document.ADDUser.VerifyCode.value != "<%=yzm%>") {
				alert("Please input correct verify code.");
				document.ADDUser.VerifyCode.focus();
				return false;
			}
			document.ADDUser.Submit.disabled=true;
			return true;
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
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;MEMBER REGISTER
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

.account_details .details_row .ul_usertype{list-style: none; display:inline-block; width:90%; padding:0px; }
.account_details .details_row .li_usertype{display:inline-block; float:left; height:30px; line-height:30px; margin-right:10px; }
.account_details .details_row .input_usertype{vertical-align:middle; margin-top:-2px; }

.account_details .details_row .input_verify{width:59%; height:30px; line-height:30px; text-indent:2px; }
.account_details .details_row .verifycode{width:30%; display:inline-block; height:30px; line-height:30px; }
.account_details .details_row .btn_submit{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; color:#db3e54; background-color:#fee; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
.account_details .details_row .btn_common{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; color:#888; background-color:#ffe; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
</style>

		<div class="account_details">
		<form name="ADDUser" method="POST" action="reg_save.asp" onSubmit="return Checkreg();">
			<div class="mobile-section"><div class="details_title">Member Register</div></div>
			<div class="details_row">
				<input class="input_common" type="text"  name="UserId" value="" maxlength="16" placeholder="Account ID (Numbers, Letters Only)" autocomplete="off" onkeyup="value=value.replace(/[^a-zA-Z0-9]/g,'')">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw1" maxlength="16" placeholder="Password" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="password" name="pw2" maxlength="16" placeholder="Repeat Password" autocomplete="off">
			</div>

            <%
			  '账户类型所在行
			  NormalAccounts=""
			  SpecialAccounts=""
			  BasicKou=9 ' 比它大的放NormalAccounts，比它小的放SpecialAccounts
			  If enusertype1<>"" and  instr(enusertype1,"VIP")=0 Then
			    If kou1>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""1"" checked> "&enusertype1
				  If kou1<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou1*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""1"" checked> "&enusertype1&" (Special)</li>"
				End If 
			  End If 
			  If enusertype2<>"" and  instr(enusertype2,"VIP")=0 Then
			    If kou2>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""2""> "&enusertype2
				  If kou2<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou2*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""2""> "&enusertype2&" (Special)</li>"
				End If 
			  End If 
			  If enusertype3<>"" and  instr(enusertype3,"VIP")=0 Then
			    If kou3>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""3""> "&enusertype3
				  If kou3<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou3*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""3""> "&enusertype3&" (Special)</li>"
				End If 
			  End If 
			  If enusertype4<>"" and  instr(enusertype4,"VIP")=0 Then
			    If kou4>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""4""> "&enusertype4
				  If kou4<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou4*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""4""> "&enusertype4&" (Special)</li>"
				End If 
			  End If 
			  If enusertype5<>"" and  instr(enusertype5,"VIP")=0 Then
			    If kou5>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""5""> "&enusertype5
				  If kou5<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou5*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""5""> "&enusertype5&" (Special)</li>"
				End If 
			  End If 
			  If enusertype6<>"" and  instr(enusertype6,"VIP")=0 Then
			    If kou6>=BasicKou Then
			      NormalAccounts=NormalAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""6""> "&enusertype6
				  If kou6<>10 Then NormalAccounts=NormalAccounts&" (-"&(100-kou6*10)&"%)"
				  NormalAccounts=NormalAccounts&"</li>"
				Else
				  SpecialAccounts=SpecialAccounts&"<li class=""li_usertype""><input class=""input_usertype"" type=""radio"" name=""UserType"" value=""6""> "&enusertype6&" (Special)</li>"
				End If 
			  End If 
			%>
			<div class="details_row">
				<ul class="ul_usertype">
				<%
				   If NormalAccounts<>"" Then response.write NormalAccounts
				   If SpecialAccounts<>"" Then response.write SpecialAccounts
				%>
				</ul>
			</div>

			<div class="details_row">
				<input class="input_common" type="text" name="UserMail" maxlength="50" placeholder="Email Address" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_verify" type="text" name="VerifyCode" maxlength="12" placeholder="Verify Code" autocomplete="off">
				<div class="verifycode"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_a%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_b%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_c%>.gif"><img align="top" height="15" border="0" src="../images/yzm/<%=yzm_skin%>/<%=yzm_d%>.gif"></div>
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" value="Submit Form" name="Submit">
			</div>
			<div class="details_row">
				<input class="btn_common"  type="reset" value="Reset Form" name="Reset">
			</div>
			<input type="hidden" name="adduser" value="true">
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
