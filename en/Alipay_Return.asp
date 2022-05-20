<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp" -->
<!--#include file="../include/Alipay_md5.asp"-->

<%
Dim partner,key
set rs_pay_online = conn.execute("select * from pay_type where PayTypeClass='online' and Display=true and instr(enPayTypeDefine,'Alipay')>0")
If Not rs_pay_online.eof	Then 
	partner = rs_pay_online("key1")
	key = rs_pay_online("key2")
End If 
rs_pay_online.close
set rs_pay_online=Nothing

'*******************获取支付宝POST过来通知消息**********************
For Each varItem in request.QueryString 
	mystr=varItem&"="&Request.QueryString(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
	mystr=Left(mystr,Len(mystr)-1)
End If 
mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'对参数排序
For i = Count TO 0 Step -1
	minmax = mystr( 0 )
	minmaxSlot = 0
	For j = 1 To i
		mark = (mystr( j ) > minmax)
		If mark Then 
			minmax = mystr( j )
			minmaxSlot = j
		End If 
	Next
	If minmaxSlot <> i Then 
		temp = mystr( minmaxSlot )
		mystr( minmaxSlot ) = mystr( i )
		mystr( i ) = temp
	End If
Next

'构造md5摘要字符串
 For j = 0 To Count Step 1
	value = SPLIT(mystr( j ), "=")
	If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  Then
		If j=Count Then
			md5str= md5str&mystr( j )
		Else 
			md5str= md5str&mystr( j )&"&"
		End If 
	End If 
 Next
md5str=md5str&key
mysign=md5(md5str)

'如果服务器支持文本写入日志，那么可以打开下边注释部分，方便测试。
'set fso = createobject("scripting.filesystemobject")
'set file = fso.opentextfile(server.mappath("test.txt"),8,true) '1读取 2写入 8追加
'file.write "partner: "& partner & vbCrLf
'file.write "key: "& key & vbCrLf
'file.write md5str & " -->MD5: "& mysign & vbCrLf
'file.write "request.QueryString(""sign""): " & request.QueryString("sign") & vbCrLf & vbCrLf
'file.close
'set file = nothing
'set fso = nothing

'*************************交易状态返回处理*************************
If mysign=request.QueryString("sign") And (request.QueryString("trade_status")="TRADE_FINISHED" Or request.QueryString("trade_status")="TRADE_SUCCESS") Then 
	return_title = "Alipay Online Payment Successful"
	return_content = "Your Alipay online payment is successful.<br>We will process your ordersoon."
Else
	return_title = "Alipay Online Payment Failed"
	return_content = "Your Alipay online payment is failed.<br>Please arrange payment again."
End If 
%>

<!DOCTYPE html>
<head id="Head">
	<title><%=return_title%>-<%=ensitename%>-<%=siteurl%></title>
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

		<script language=javascript>
			setTimeout("window.opener.location.href='my_purchased.asp'",3000);
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
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;<%=UCase(return_title)%>
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
			<div class="mobile-section"><div class="details_title"><%=return_title%></div></div>
			<div class="details_row">
				<%=return_content%>
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Close this Window" onclick="javascript:window.opener.location.reload();window.close();">
			</div>
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
