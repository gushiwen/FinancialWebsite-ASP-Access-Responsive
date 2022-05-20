<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
OrderNum=request("OrderNum")
MobilePhone=request("MobilePhone")

if request("send")="ok" then 
	call send()
else

If request.cookies("goldweb")("userid")<>"" Then 
	sqlinfo = "select * from goldweb_OrderList where OrderNum='"&OrderNum&"' and UserId='"&request.cookies("goldweb")("userid")&"'"
ElseIf request.cookies("goldweb")("userid")="" And MobilePhone<>"" Then 
	sqlinfo = "select * from goldweb_OrderList where OrderNum='"&OrderNum&"' and RecHomePhone='"&MobilePhone&"'"
Else	
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('You are not authorized to view this order.');"
	response.write "location.href='index.asp';"
	response.write "</script>"
	response.end
End If 

set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sqlinfo,conn,1,1
if rs.eof and rs.bof  then
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('We can not find this order.');"
	response.write "location.href='index.asp';"
	response.write "</script>"
	response.end
Else
	OrderSum=formatnumber(rs("OrderSum"),2)
	UserName=rs("RecName")
	HomePhone=rs("RecHomePhone")
	UserMail=rs("RecMail")
	Address=rs("RecAddress")
	Country=rs("RecCountry")
	ZipCode=rs("RecZipcode")
End If 

rs.close
set rs=Nothing
%>

<!DOCTYPE html>
<head id="Head">
	<title>Remittance Report-<%=ensitename%>-<%=siteurl%></title>
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
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;REMITTANCE REPORT
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
.account_details .details_row .textarea_common{width:90%; height:60px; text-indent:2px; overflow:hidden; }
.account_details .details_row .btn_submit{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; font-weight:bold; color:#db3e54; background-color:#fee; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
.account_details .details_row .btn_common{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; color:#888; background-color:#ffe; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
</style>

		<div class="account_details">
		<form name="add" method="post" action="my_remittance.asp">
			<div class="mobile-section"><div class="details_title">Remittance Report</div></div>

			<div class="details_row">
				<input class="input_common" type="text" name="hk1" maxlength="14" value="<%=OrderNum%>" placeholder="Order No. Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="hk2" maxlength="10" value="<%=englobalpriceunit%> <%=OrderSum%>" placeholder="Payment Amount Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="hk3" maxlength="25" placeholder="From Bank Name Required" autocomplete="off">
			</div>

			<div class="details_row">
				<select class="input_common" name="hk4">
					<option selected value="">Payment by Required</option>
<%
	 ' 列出银行转账方式
	set rs_pay_bank=conn.execute("select * from pay_type where PayTypeClass='bank' and Display=true order by PayType")
	if rs_pay_bank.eof and rs_pay_bank.bof  then
	Else
	  do while not rs_pay_bank.eof
	    response.write "<option value="""&rs_pay_bank("enPayTypeDefine")&"""> "&rs_pay_bank("enPayTypeDefine")&"</option>"
	    rs_pay_bank.movenext
	  Loop
	End If 
	rs_pay_bank.close
	set rs_pay_bank=nothing
	%>
				</select>
			</div>

			<div class="details_row">
				<textarea class="textarea_common" name="hk11" placeholder="Notes ≤200 Characters"></textarea>
			</div>
			<div class="details_row">
				<input class="btn_submit" type="submit" value="Submit Form" name="Submit">
			</div>
			<div class="details_row">
				<input class="btn_common"  type="reset" value="Reset Form" name="Reset">
			</div>
			<div class="details_row">
				<input class="btn_common" type="button" value="Return to Last Page" onclick="document.location.href='my_order_view.asp?OrderNum=<%=OrderNum%>&MobilePhone=<%=server.URLEncode(MobilePhone)%>';">
			</div>
			<input name="send" type="hidden" value="ok">
			<input name="OrderNum" type="hidden" value="<%=OrderNum%>">
			<input name="MobilePhone" type="hidden" value="<%=MobilePhone%>">
			<input name="hk5" type="hidden" value="<%=UserName%>">
			<input name="hk6" type="hidden" value="<%=HomePhone%>">
			<input name="hk7" type="hidden" value="<%=UserMail%>">
			<input name="hk8" type="hidden" value="<%=Address%>">
			<input name="hk9" type="hidden" value="<%=Country%>">
			<input name="hk10" type="hidden" value="<%=ZipCode%>">
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
End If

conn.close
set conn=nothing

sub send()
	call goldweb_check_path()

	if request.form("hk1")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in Order Number.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end If
	
	if request.form("hk2")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in Payment Amount.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if request.form("hk3")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please fill in From Bank Name.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if request.form("hk4")="" then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please choose Payment By.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if llen(request.form("hk11"))>200  then
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('The note is too long.');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	dim MailTitle
	dim MailContent
	MailTitle = request.form("hk5") & " has reported remittance on " & siteurl
	MailContent="Order No.: " & request.form("hk1")  & "<br>"
	MailContent=MailContent & "Payment Amount: " & request.form("hk2")  & "<br>"
	MailContent=MailContent & "From Bank Name: " & request.form("hk3")  & "<br>"
	MailContent=MailContent & "Payment by: " & request.form("hk4")  & "<br>"
	MailContent=MailContent & "Full Name: " & request.form("hk5")  & "<br>"
	MailContent=MailContent & "Mobile Number: " & request.form("hk6")  & "<br>"
	MailContent=MailContent & "Email: " & request.form("hk7")  & "<br>"
	MailContent=MailContent & "Address: " & request.form("hk8")  & "<br>"
	MailContent=MailContent & "Country: " & request.form("hk9")  & "<br>"
	MailContent=MailContent & "Post Code: " & request.form("hk10")  & "<br>"
	MailContent=MailContent & "Note: " & request.form("hk11")  & "<br>"
	Call SendOutMail(MailTitle, MailContent, adm_mail, adm_mail, request.form("hk7"))

	sql="select * from hand"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
  	rs("name")="admin"
  	rs("neirong")=MailContent
  	rs("riqi")=DateAdd("h", TimeOffset, now())
	If request.cookies("goldweb")("userid")<>"" Then 
 		rs("fname")=request.cookies("goldweb")("userid")
	Else
 		rs("fname")="Guest"
	End If 
  	rs("zuti")="1"
	rs.update	
	rs.close
	set rs=nothing

	conn.close
	set conn=Nothing

	response.write "<script language='javascript'>"
	response.write "alert('The message has been sent.');"
	response.write "</script>"
	response.write "<meta http-equiv=refresh content='0;URL=my_order_view.asp?OrderNum="&OrderNum&"&MobilePhone="&server.URLEncode(MobilePhone)&"'>"
end sub
%>
