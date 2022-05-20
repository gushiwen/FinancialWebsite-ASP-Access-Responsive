<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/Alipay_Payto.asp" -->
<%
'call aspsql()
'call checklogin()
OrderNum=request("OrderNum")
MobilePhone=request("MobilePhone")
%>

<!DOCTYPE html>
<head id="Head">
	<title>Order Details-<%=ensitename%>-<%=siteurl%></title>
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
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;ORDER DETAILS
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

.order_details{font-size:13px; }
.order_details .details_title{background:#d2364c; color:#fff; display:block; width:100%; height:36px; line-height:36px; padding:0px; text-align:center; font-size:14px; font-weight:600; }
.order_details .details_row{width:100%; margin-top:10px; padding-left:10px; text-align:left; }
.order_details .details_row .input_common{width:90%; height:30px; line-height:30px; text-indent:2px; }
.order_details .details_row .textarea_common{width:90%; height:60px; text-indent:2px; }
.order_details .details_row .btn_pay{display:inline-block; width:90%; height:30px; line-height:30px; font-weight:bold; color:#db3e54; background-color:#fee; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
.order_details .details_row .text_bank{padding:10px 0px 0px 10px; }

.content-box{overflow:hidden; width:100%;max-width: 1000px;margin: 0px auto; margin-top:20px; font-size:13px; }
/*主体*/
.item-content{width:100%; overflow:hidden; position:relative; padding-bottom:15px; border-top:#fff 1px solid; }
.item-content .td {float: left;}
/*商品信息*/
.td.td-item{float:left; width:100%; padding-left:0px ;}
/*图片*/
.td-item .item-info {margin:-3px 0px 0px 91px;}
.item-content .item-pic a {text-align: center;}
.item-content .item-pic {width:80px; height:80px; border:1px solid #EEE; float:left; overflow:hidden; margin-top:20px; margin-left:3px;}	
.td-item .item-basic-info { height:40px; line-height:1.2; margin-top:20px; text-align:left; font-weight:600; }
/*规格*/
.td.td-info{position: absolute;left:0px;top:55px;width:100%;overflow: hidden;padding-left:90px;line-height: 16px;}
.td.td-info .sku-line {float:left;color: #9C9C9C;overflow: hidden;}
/*数量*/
.td.td-amount{position: absolute;left:0px;bottom:10px;width:100%;overflow: hidden;padding-left:90px;line-height: 16px;}
.td.td-amount .sl {float:left; padding:0px; }
.td.td-amount .sl .min{display:inline-block; border:1px solid #ddd; background-color:#ffe; font-size:16px; padding:0px 0px 4px 0px; width:30px; height:30px; line-height:26px; }
.td.td-amount .sl .add{display:inline-block; border:1px solid #ddd; background-color:#ffe; font-size:16px; padding:0px 0px 2px 0px; width:30px; height:30px; line-height:28px; }
.td.td-amount .sl input.text_box{border:1px solid #ddd; font-size:12px; padding:0px; width:40px; height:30px; line-height:30px; text-align:center; }
.td.td-amount .sl .del{display:inline-block; border:1px solid #ddd; background-color:#fee; font-size:12px; padding:0px; width:50px; height:30px; line-height:30px; margin-left:10px;}

.sub-total {width:100%; overflow: hidden; border-top:#fff 1px solid; padding:0px 10px 10px 0px; }
.sub-total .member { font-size:12px; padding-top:5px; }
.sub-total .sub-price {width:100%; text-align:right; padding-top:15px; }
.sub-total .discount-price {width:100%; text-align:right; padding-top:5px; }
.sub-total .total-price {width:100%; text-align:right; font-weight:600; padding-top:5px; }

.customer_details{margin-top:10px; }
.customer_details .details_title{background:#d2364c; color:#fff; display:block; width:100%; height:36px; line-height:36px; padding:0px; text-align:center; font-size:14px; font-weight:600; }
.customer_details .details_row{width:100%; margin-top:10px; text-align:center; }
.customer_details .details_row .input_common{width:90%; height:30px; line-height:30px; text-indent:2px; }
.customer_details .details_row .textarea_common{width:90%; height:60px; text-indent:2px; }
.customer_details .details_row .btn_common{display:inline-block; cursor:pointer; width:90%; height:30px; line-height:30px; color:#888; background-color:#ffe; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
</style>

<%
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
else
%>
		<div class="order_details">
			<div class="mobile-section"><div class="details_title">Order Overview</div></div>
			<div class="details_row">
				<strong><%=rs("OrderNum")%></strong>&nbsp;&nbsp;&nbsp;&nbsp;<%=DatePart("yyyy",rs("OrderTime"))%>-<%=Right("0"&DatePart("m",rs("OrderTime")),2)%>-<%=Right("0"&DatePart("d",rs("OrderTime")),2)%>&nbsp;&nbsp;&nbsp;&nbsp;<%=englobalpriceunit%><%=formatnumber(rs("OrderSum"),2)%>
			</div>
			<div class="details_row">
				Order Status: <% response.write conn.execute("select enStatusDefine from order_type where Status='" & rs("Status") & "'")("enStatusDefine") %>
			</div>
			<div class="details_row">
				Payment Status: <% response.write conn.execute("select enPayTypeDefine from pay_type where PayType='" & rs("PayType") & "'")("enPayTypeDefine") %>
			</div>
			<%
				If rs("PayType")="0" And rs("Status")<>"11" And rs("Status")<>"12" And rs("Status")<>"99" Then   
					ListPaymentType()
				End If 
			%>
		</div>


		<!--购物车内Information -->
		<div class="content-box">
<%
 	Sum = 0
	set rs2=conn.execute("select A.ProdId,A.BuyPrice,A.ProdUnit,B.enProdName,B.enPriceUnit from goldweb_Order A left join goldweb_product B on A.ProdId=B.Prodid where A.OrderNum='"&OrderNum&"' order by A.ID")
	do while not rs2.eof 
		Sum = Sum + FormatNumber(rs2("BuyPrice"),2)*CInt(rs2("ProdUnit"))
%> 
			<ul class="item-content">
				<li class="td td-item">
					<div class="item-pic">
						<a href="product.asp?ProdId=<%=rs2("Prodid")%>"><img src="<%=conn.execute("select ImgPrev from goldweb_product where ProdId='"&rs2("Prodid")&"'")("ImgPrev")%>" width="80" height="80" border="0" onload="DrawImage(this,80,80)" ></a>
					</div>
					<div class="item-info">
						<div class="item-basic-info">
							<a href="product.asp?ProdId=<%=rs2("Prodid")%>" class="item-title J_MakePoint" data-point="tbcart.8.11"><%=rs2("enProdName")%></a>
						</div>
					</div>
				</li>
				<li class="td td-info">
					<span class="sku-line">
						Unit Price <%=rs2("enPriceUnit")%>&nbsp;<%=FormatNumber(rs2("BuyPrice"),2)%>
					</span>
				</li>
				<li class="td td-amount">
					<div class="sl">
						Quantity <%=rs2("ProdUnit")%>x
					</div>
				</li>
			</ul>

<%
	rs2.movenext
	loop
	rs2.close
    set rs2=Nothing
    
	userkou=FormatNumber(rs("thiskou"),2)
    fei=rs("fei")
%> 

			<div class="sub-total">
				<div class="sub-price">
					Sub Total: <%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum,2)%>
				</div>
				<%
					If userkou<>10 then
				%>
				<div class="discount-price">
					After -<%=100-10*userkou%>% Discount: <%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10,2)%>
				</div>
				<%
					End If 
				%>
				<div class="total-price">
					Total Amount: <%=englobalpriceunit%>&nbsp;<%=FormatNumber((Sum*userkou/10+fei),2)%>
				</div>
			</div>

			<div class="clear"></div>
		</div>

<%
rs.movefirst
%>

		<!--客户详细资料-->
		<div class="customer_details">
			<div class="details_title">Customer Details</div> 
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("pei")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecName")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecHomePhone")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecMail")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecAddress")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecCountry")%>" readonly="true" >
			</div>
			<div class="details_row">
				<input class="input_common" type="text" value="<%=rs("RecZipcode")%>" readonly="true" >
			</div>
			<% If Trim(rs("Notes"))<>"" Then %>
			<div class="details_row">
				<textarea class="textarea_common" disabled ><%=rs("Notes")%></textarea>
			</div>
			<% End If %>
			<div class="details_row">
				<textarea class="textarea_common" disabled ><%=rs("memo")%></textarea>
			</div>
		</div>

<%
end If

rs.close
set rs=nothing
%>
<!-- E form -->

</div></div></div></div>

</section>
<!-- E content -->


<!-- S sidebar -->
<section class="sidebar float-left">
<%
If request.cookies("goldweb")("userid")<>"" Then 
%>
<!--#include file="goldweb_usertree.asp"-->
<%
Else
%>
<!--#include file="goldweb_proclasstree.asp"-->
<%
End If 
%>
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
'列出支付方式
Sub ListPaymentType()

INTERFACE_URL="https://www.alipay.com/cooperate/gateway.do?"

' 列出在线支付方式
sql_pay_online = "select * from pay_type where PayTypeClass='online' and Display=true order by PayType desc"
set rs_pay_online=Server.Createobject("ADODB.RecordSet")
rs_pay_online.Open sql_pay_online,conn,1,1
if rs_pay_online.eof and rs_pay_online.bof  then

Else
  do while not rs_pay_online.eof
  if InStr(rs_pay_online("enPayTypeDefine"),"Paypal")>0 then
%>
<div style='padding-top: 5px;'>
<form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="new">
<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="<%=rs_pay_online("Key1")%>">
<input type="hidden" name="item_name" value="<%=siteurl&" online order: "&rs("OrderNum")%>">
<input type="hidden" name="amount" value="<%=rs("OrderSum")%>">
<input type="hidden" name="item_number" value="<%=rs("OrderNum")%>">
<input type="hidden" name="currency_code" value="<%=rs_pay_online("Key5")%>">
<input type="hidden" name="notify_url" value="<%=FullSiteUrl%>/include/Paypal_Notify.asp">
<input type="hidden" name="return" value="<%=FullSiteUrl%>/en/Paypal_Return.asp">
<input type="hidden" name="cancel_return" value="<%=FullSiteUrl%>/en/Paypal_Cancel_Return.asp">
<input type="submit" style="font-size:10px; padding:2px;" value="  <%=rs_pay_online("enPayTypeDefine")%> (<%=englobalpriceunit%> <%=FormatNumber(rs("OrderSum"),2)%>)  ">
</form>
</div>
<%
  elseif InStr(rs_pay_online("enPayTypeDefine"),"Alipay")>0 then
  		'海外商家请求参数设置
		Dim Alipay_gateway,Alipay_service,Alipay_partner,Alipay_input_charset,Alipay_sign_type,Alipay_key
		Dim Alipay_notify_url,Alipay_return_url
		Dim Alipay_subject,Alipay_out_trade_no,Alipay_currency,Alipay_total_fee
		Dim Alipay_Obj,Alipay_itemUrl
		Alipay_gateway = "https://mapi.alipay.com/gateway.do?"
		Alipay_service =	"create_forex_trade"
		Alipay_partner	=	rs_pay_online("Key1")	
		Alipay_input_charset = WebPageCharset
		Alipay_sign_type = "MD5"
		Alipay_key = rs_pay_online("Key2")
		Alipay_notify_url = FullSiteUrl&"/include/Alipay_Notify.asp"
		Alipay_return_url = FullSiteUrl&"/en/Alipay_Return.asp"
		Alipay_subject = siteurl&" online order: "&rs("OrderNum")
		Alipay_out_trade_no = rs("OrderNum")
		Alipay_currency = rs_pay_online("Key5")
		Alipay_total_fee = FormatNumber(FormatNumber(rs("OrderSum"),2)*FormatNumber(rs_pay_online("key6"),2),2)

		'海外商家即时到账接口
		Set Alipay_Obj	= New creatAlipayItemURL
		Alipay_itemUrl=Alipay_Obj.creatAlipayItemURL(Alipay_gateway,Alipay_service,Alipay_partner,Alipay_input_charset,Alipay_sign_type,Alipay_key,Alipay_notify_url,Alipay_return_url,Alipay_subject,Alipay_out_trade_no,Alipay_currency,Alipay_total_fee)

  		'国内商家请求参数设置
		'Dim Alipay_gateway,Alipay_service,Alipay_partner,Alipay_input_charset,Alipay_sign_type,Alipay_key
		'Dim Alipay_notify_url,Alipay_return_url
		'Dim Alipay_out_trade_no,Alipay_subject,Alipay_payment_type,Alipay_total_fee,Alipay_seller_email
		'Dim Alipay_Obj,Alipay_itemUrl
		'Alipay_gateway = "https://mapi.alipay.com/gateway.do?"
		'Alipay_service =	"create_direct_pay_by_user"
		'Alipay_partner	=	rs_pay_online("Key1")	
		'Alipay_input_charset = WebPageCharset
		'Alipay_sign_type = "MD5"
		'Alipay_key = rs_pay_online("Key2")
		'Alipay_notify_url = FullSiteUrl&"/include/Alipay_Notify.asp"
		'Alipay_return_url = FullSiteUrl&"/ch/Alipay_Return.asp"
		'Alipay_out_trade_no = rs("OrderNum")
		'Alipay_subject = siteurl&" 网上订单："&rs("OrderNum")&" 的付款"
		'Alipay_payment_type = "1"
		'Alipay_total_fee = FormatNumber(rs("OrderSum"),2)
		'Alipay_seller_email = rs_pay_online("Key3") 

		'国内商家即时到账接口
		'Set Alipay_Obj	= New creatAlipayItemURL
		'Alipay_itemUrl=Alipay_Obj.creatAlipayItemURL(Alipay_gateway,Alipay_service,Alipay_partner,Alipay_input_charset,Alipay_sign_type,Alipay_key,Alipay_notify_url,Alipay_return_url,Alipay_out_trade_no,Alipay_subject,Alipay_payment_type,Alipay_total_fee,Alipay_seller_email)
%>
    <div style='padding-top: 5px;'>
    <form action="<%=Alipay_itemUrl%>" method="post" target="new">
      <input type="submit" style="font-size:10px; padding:2px;" value="  <%=rs_pay_online("enPayTypeDefine")%>(exchange to RenMinBi ￥<%=FormatNumber(FormatNumber(rs("OrderSum"),2)*FormatNumber(rs_pay_online("key6"),2),2)%>)  ">
	</form>
	</div>
<%
  end If

  rs_pay_online.movenext
  Loop
end If

rs_pay_online.close
set rs_pay_online=nothing

' 列出银行转账方式
sql_pay_bank = "select * from pay_type where PayTypeClass='bank' and Display=true order by PayType"
set rs_pay_bank=Server.Createobject("ADODB.RecordSet")
rs_pay_bank.Open sql_pay_bank,conn,1,1

if rs_pay_bank.eof and rs_pay_bank.bof  then

Else
  response.write "<div style='padding-top: 5px;'>"
  response.write "<div id='pay_bank_1'><input type='button' value='  View Paynow / Bank / Alipay / Wechat Accounts  ' style='font-size:10px; padding:2px;' onclick='javascript:document.getElementById(""pay_bank_1"").style.display=""none"";document.getElementById(""pay_bank_2"").style.display="""";'></div>"
  response.write "<div id='pay_bank_2' style='display:none;'>"
  response.write "<input type='button' value='  Close Paynow / Bank / Alipay / Wechat Accounts  ' style='font-size:10px; padding:2px;' onclick='javascript:document.getElementById(""pay_bank_1"").style.display="""";document.getElementById(""pay_bank_2"").style.display=""none"";'>&nbsp;&nbsp;<input type='button' value='  Remittance Report  ' style='font-size:10px; padding:2px;' onclick='javascript:window.location.href=""my_remittance.asp?OrderNum="&rs("OrderNum")&"&MobilePhone="&server.URLEncode(MobilePhone)&""";'><br><br>"
  response.write "<div style='padding-left:20px;'>"
  do while not rs_pay_bank.eof
	if InStr(rs_pay_bank("enPayTypeDefine"),"Alipay")>0 Or InStr(rs_pay_bank("enPayTypeDefine"),"Wechat")>0 then
		response.write rs_pay_bank("enPayTypeDefine")&FormatNumber(FormatNumber(rs("OrderSum"),2)*FormatNumber(rs_pay_bank("key6"),2),2)&"<br>Account: "&rs_pay_bank("Key1")&"<br>Name: "&rs_pay_bank("Key2")
	Else
		response.write rs_pay_bank("enPayTypeDefine")&" "&englobalpriceunit&FormatNumber(rs("OrderSum"),2)&"<br>Account: "&rs_pay_bank("Key1")&"<br>Name: "&rs_pay_bank("Key2")
	End If 
	if rs_pay_bank("Pic")<>"" then
		response.write "<br><img src='"&rs_pay_bank("Pic")&"' border='0'/>"
	End If 
	response.write "<br><br>"
    rs_pay_bank.movenext
  Loop
  response.write "</div>"
  response.write "</div>"
  response.write "</div>"
End If 
  
rs_pay_bank.close
set rs_pay_bank=nothing

End Sub

conn.close
set conn=nothing
%>
