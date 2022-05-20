<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<!--#include file="../include/goldweb_auto_lock.asp" -->
<!--#include file="chopchar.asp"-->

<!DOCTYPE html>
<head id="Head">
	<title>Shopping Cart-<%=ensitename%>-<%=siteurl%></title>
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

		<script language="JavaScript" type="text/JavaScript">
			function DoDeleteProduct(ProductId) {
				document.shoppingcart.DeleteProduct.value=ProductId;
				document.shoppingcart.action="shoppingcart.asp";
				document.shoppingcart.submit();
 			}
			function DoMinQty(ProductId) {
				if(parseInt(document.getElementById(ProductId).value)-1<1){
					return false;
				}
				else
				{
					document.getElementById(ProductId).value=parseInt(document.getElementById(ProductId).value)-1;
				}
				DoModifyQuantity();
 			}
			function DoAddQty(ProductId) {
				document.getElementById(ProductId).value=parseInt(document.getElementById(ProductId).value)+1;
				DoModifyQuantity();
 			}
			function DoModifyQuantity() {
				document.shoppingcart.ModifyQuantity.value="yes";
				document.shoppingcart.action="shoppingcart.asp";
				document.shoppingcart.submit();
 			}

			function CheckForm()
			{
				if (document.shoppingcart.Post.value == 0) {
					alert("Please arrange the delivery.");
					document.shoppingcart.Post.focus();
					return false;
				}
				if (document.shoppingcart.RecName.value.length <2 || document.shoppingcart.RecName.value.length>30 ) {
					alert("Please check contact name.");
					document.shoppingcart.RecName.focus();
					return false;
				}
				if (document.shoppingcart.HomePhone.value.length < 6 || document.shoppingcart.HomePhone.value.length > 20) {
					alert("Please check mobile phone number.");
					document.shoppingcart.HomePhone.focus();
					return false;
				}
				if (document.shoppingcart.RecMail.value.length <11 || document.shoppingcart.RecMail.value.length > 50) {
					alert("Please check e-mail address.");
					document.shoppingcart.RecMail.focus();
					return false;
				}
				if (document.shoppingcart.RecMail.value.length > 0 && !document.shoppingcart.RecMail.value.match( /^.+@.+$/ ) ) {
 					alert("Please check your e-mail address.");
 					document.shoppingcart.RecMail.focus();
					return false;
				}
				if (document.shoppingcart.address.value.length < 3 || document.shoppingcart.address.value.length > 150) {
					alert("Please check your address.");
					document.shoppingcart.address.focus();
					return false;
				}
				if (document.shoppingcart.Country.value.length < 2 || document.shoppingcart.Country.value.length > 50) {
					alert("Please check your country.");
					document.shoppingcart.address.focus();
					return false;
				}
				if (document.shoppingcart.ZipCode.value.length < 4 || document.shoppingcart.ZipCode.value.length >12) {
					alert("Please check the post code.");
					document.shoppingcart.ZipCode.focus();
					return false;
				}
				document.shoppingcart.Submit.disabled=true;
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
<span>CURRENT: </span><a class="breadcrumbs-home" href="javascript:;">SHOPPING CART</a><i></i>
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

.content-box{overflow:hidden; width:100%; max-width:1000px; margin:0px auto; font-size:13px; }
.content-box .no-item {width:90%; text-align:center; }
/*主体*/
.item-content{width:100%; overflow:hidden; position:relative; padding-bottom:15px; border-bottom:#fff 1px solid; }
.item-content .td {float: left;}
/*商品信息*/
.td.td-item{float:left;width:100%;padding-left:0px ;}
/*图片*/
.td-item .item-info {margin: -3px 0px 0px 91px;}
.item-content .item-pic a {text-align: center;}
.item-content .item-pic {width:80px; height:80px; border: 1px solid #EEE; float: left;overflow: hidden;margin-top:20px;margin-left:3px;}	
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

.sub-total {width:100%; overflow:hidden; padding:0px 10px 10px 0px; }
.sub-total .member {font-size:12px; padding:5px 0px 0px 10px; }
.sub-total .sub-price {width:100%; text-align:right; padding-top:15px; }
.sub-total .discount-price {width:100%; text-align:right; font-weight:600; padding-top:5px; }

.customer_details{margin-top:10px; margin-bottom:50px; }
.customer_details .details_title{background:#d2364c; color:#fff; display:block; width:100%; height:36px; line-height:36px; padding:0px; text-align:center; font-size:14px; font-weight:600; }
.customer_details .details_row{width:100%; margin-top:10px; text-align:center; }
.customer_details .details_row .input_common{width:90%; height:30px; line-height:30px; text-indent:2px; }
.customer_details .details_row .textarea_common{width:90%; height:60px; text-indent:2px; overflow:hidden; }

.float-bar-wrapper{overflow:hidden; width:100%; background:#fff; border-top:1px solid #e5e5e5; }
.price-sum{overflow:hidden; float:right; padding-right:6px; line-height:38px; vertical-align:middle; }
.price-sum .price {color:#F40; font-weight:400; font-size:16px; height:38px; line-height:38px; vertical-align:middle; }
.btn-area{float:right; width:80px;text-decoration: none; text-align: center;}
.btn-submit{display:inline-block; cursor:pointer; width:100%; height:38px; line-height:38px; font-size:16px; background-color:#F40; color: #fff; border:0px; border-radius:0px; }
</style>

<%
'call aspsql()
userkou=10
buylist=request.cookies("goldweb")("cart")

If request.cookies("goldweb")("userid")<>"" Then 
set rs=conn.execute ("select * from goldweb_user where userid='"&request.cookies("goldweb")("userid")&"'")
if not (rs.eof and rs.bof) then
HomePhone=rs("HomePhone")
CompPhone=rs("CompPhone")
Country=rs("Country")
address=rs("Address")
Zipcode=rs("ZipCode")
RecMail=rs("UserMail")
RecName=rs("UserName")
UserKou=rs("UserKou")
UserType=rs("UserType")
end If
set rs=Nothing
End If 

'If IsNull(Country) Then
'	Country="Singapore"
'Elseif Trim(Country)="" Then
'	Country="Singapore"
'End If
Post=1

' Delete product in shopping cart
if trim(request("DeleteProduct"))<>"" then 
buylist=replace(buylist,trim(request("DeleteProduct")),"XXXXXXXX")
response.Cookies("goldweb").path="/"
response.cookies("goldweb")("cart")=buylist

Post=trim(request("Post"))
RecName=trim(request("RecName"))
HomePhone=trim(request("HomePhone"))
RecMail=trim(request("RecMail"))
address=trim(request("address"))
Country=trim(request("Country"))
Zipcode=trim(request("Zipcode"))
Notes=trim(request("Notes"))
end If

' Modify product in shopping cart
If Request("ModifyQuantity") <> "" Then
buylist = ""
buyid = Split(Request("ProdId"), ", ")
For I=0 To UBound(buyid)
	If IsNumeric(request(buyid(I))) Then
		if buylist="" then
			buylist = "'" & buyid(I) & "', '"&CInt(request(buyid(I)))&"'"
		else
			buylist = buylist & ", '" & buyid(I) & "', '"&CInt(request(buyid(I)))&"'"
		End if
	end if
Next
response.Cookies("goldweb").path="/"
response.cookies("goldweb")("cart") = buylist

Post=trim(request("Post"))
GetTime=trim(request("GetTime"))
RecName=trim(request("RecName"))
HomePhone=trim(request("HomePhone"))
CompPhone=trim(request("CompPhone"))
RecMail=trim(request("RecMail"))
address=trim(request("address"))
Country=trim(request("Country"))
Zipcode=trim(request("Zipcode"))
Notes=trim(request("Notes"))
End If

if buylist="" then
	response.Cookies("goldweb").path="/"
	response.cookies("goldweb")("cart") = ""
	response.write "<div class='content-box'><div class='no-item'>Your shopping cart is empty.</div></div>"
else
	' buylist '02337', '1', '02335', '1', '01610', '1' 
	buyprodlist=""
	sp_buylist=split(buylist,", ")
	for i=0 to ubound(sp_buylist) Step 2 
		If buyprodlist="" Then 
			buyprodlist = Trim(sp_buylist(i))
		Else
			buyprodlist = buyprodlist & ", " & Trim(sp_buylist(i))
		End If 
	Next

	' buyprodlist '02337', '02335', '01610' 
	Set rs=conn.execute("select * from goldweb_product where ProdId in ("&buyprodlist&") order by instr ('"&Replace(buyprodlist,"'","")&"',ProdId)")
	if rs.eof and rs.bof then
		response.Cookies("goldweb").path="/"
		response.cookies("goldweb")("cart") = ""
		response.write "<div class=""content-box""><div class='no-item'>Your shopping cart is empty.</div></div>"
		rs.close
		set rs=nothing
	else
%>
		
		<form action="ordersubmit.asp" method="Post" name="shoppingcart" onSubmit="return CheckForm();">
		<!--购物车内信息 -->
		<div class="content-box">

<%
		Sum = 0
		for i=0 to ubound(sp_buylist) Step 2 
			if Trim(Replace(sp_buylist(i),"'",""))=rs("prodid") Then
				Quatity=Trim(Replace(sp_buylist(i+1),"'",""))

				 if not isNumeric(Quatity) then Quatity=1
				Quatity=CInt(Quatity)
				If Quatity <= 0 Then Quatity = 1
				Sum = Sum + FormatNumber(rs("enPriceList"),2)*Quatity
	  
				'判断ProdId="00001", Post=4 送货安排N.A., 数量输入框readonly="true"
				If rs("ProdId")="00001" Then
					Post=4
				End If 
%>
			<ul class="item-content">
				<li class="td td-item">
					<div class="item-pic">
						<a href="product.asp?ProdId=<%=rs("ProdId")%>"><img src="<%=rs("ImgPrev")%>" width="80" height="80" border="0" onload="DrawImage(this,80,80)" ></a><input type="hidden" name="ProdId" value="<%=rs("ProdId")%>">
					</div>
					<div class="item-info">
						<div class="item-basic-info">
							<a href="product.asp?ProdId=<%=rs("ProdId")%>" class="item-title J_MakePoint" data-point="tbcart.8.11"><%=rs("enProdName")%></a>
						</div>
					</div>
				</li>
				<li class="td td-info">
					<span class="sku-line">
						Unit Price <%=rs("enPriceUnit")%>&nbsp;<%=FormatNumber(rs("enPriceList"),2)%>
					</span>
				</li>
				<li class="td td-amount">
					<div class="sl">
						Quantity 
						<input class="min am-btn" name="" type="button" value="-" onclick="javascript:DoMinQty('<%=rs("ProdId")%>');" disabled />
						<input class="text_box" name="<%=rs("ProdId")%>" id="<%=rs("ProdId")%>" type="text" value="<%=Quatity%>" onkeyup="value=value.replace(/[^0-9]/g,'')" onchange="javascript:DoModifyQuantity();" readonly="true" />
						<input class="add am-btn" name="" type="button" value="+" onclick="javascript:DoAddQty('<%=rs("ProdId")%>');" disabled />
						<input class="del am-btn" name="" type="button" value="Delete" onclick="javascript:DoDeleteProduct('<%=rs("ProdId")%>');" />
					</div>
				</li>
			</ul>

	<%
      			rs.MoveNext
				If rs.eof Then Exit For 
			end If
		next

		rs.close
		set rs=nothing
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

			</div>

<input type="hidden" name="DeleteProduct" value="">
<input type="hidden" name="ModifyQuantity" value="">
<input type="hidden" name="Total" value="">
<input type="hidden" name="Pei" value="">
<input type="hidden" name="Fei" value="">

			<div class="clear"></div>
		</div>

		<!--客户详细资料-->
		<div class="customer_details">
			<div class="details_title">Customer Details</div> 

<%
if enmianyoufei<=Sum*userkou/10 then
	enfei1=0
	enpei1=enmianyoufei_msg
else 
	enpei1=enpei1 & " " & englobalpriceunit & enfei1 &""
end if 
if enpei2<>"" then enpei2=enpei2 & " " & englobalpriceunit & enfei2 &""
if enpei3<>"" then enpei3=enpei3 & " " & englobalpriceunit & enfei3 &""
if enpei4<>"" then enpei4=enpei4 & " " & englobalpriceunit & enfei4 &""
if enpei5<>"" then enpei5=enpei5 & " " & englobalpriceunit & enfei5 &""
if enpei6<>"" then enpei6=enpei6 & " " & englobalpriceunit & enfei6 &""

response.write "<div class='details_row'><select class='input_common' name='Post' onchange='javascript:UpdateTotalAmount();'>"
response.write "<option selected value='0'>Please arrange delivery here </option>"
If Post=1 Then
  response.write "<option value='1' selected>"&enpei1&"</option>"
Else
  response.write "<option value='1'>"&enpei1&"</option>"
End If 
if enpei2<>"" then 
If Post=2 Then
  response.write "<option value='2' selected>"&enpei2&"</option>"
Else
  response.write "<option value='2'>"&enpei2&"</option>"
End If 
End If 
if enpei3<>"" then 
If Post=3 Then
  response.write "<option value='3' selected>"&enpei3&"</option>"
Else
  response.write "<option value='3'>"&enpei3&"</option>"
End If 
End If 
if enpei4<>"" then 
If Post=4 Then
  response.write "<option value='4' selected>"&enpei4&"</option>"
Else
  response.write "<option value='4'>"&enpei4&"</option>"
End If 
End If 
if enpei5<>"" then 
If Post=5 Then
  response.write "<option value='5' selected>"&enpei5&"</option>"
Else
  response.write "<option value='5'>"&enpei5&"</option>"
End If 
End If 
if enpei6<>"" then 
If Post=6 Then
  response.write "<option value='6' selected>"&enpei6&"</option>"
Else
  response.write "<option value='6'>"&enpei6&"</option>"
End If 
End If 
response.write "</select></div>"
%>

			<div class="details_row">
				<input class="input_common" type="text" name="RecName" maxlength="30" value="<%=RecName%>" placeholder="Full Name Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="HomePhone" maxlength="20" value="<%=HomePhone%>" placeholder="Mobile Number Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="RecMail" maxlength="50" value="<%=RecMail%>" placeholder="Email Address Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="address" maxlength="150" value="<%=Address%>" placeholder="Detailed Address Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="Country" maxlength="50" value="<%=Country%>" placeholder="Country Required" autocomplete="off">
			</div>
			<div class="details_row">
				<input class="input_common" type="text" name="ZipCode" maxlength="12" value="<%=Zipcode%>" placeholder="Post Code Required" autocomplete="off">
			</div>
			<div class="details_row">
				<textarea class="textarea_common" name="Notes" placeholder="Any Other Requests"></textarea>
			</div>
			<div class="details_row">
				<span style="color:red; ">* All information are non-refundable after purchasing!</span>
			</div>
		</div>

		<div class="float-bar-wrapper">
			<div class="btn-area">
				<input class="btn-submit" type="submit" value="Submit" name="Submit">
			</div>
			<div class="price-sum">
				<span class="txt">Total Amount <span id="Tr_Amount" style="display:none;"></span></span>
				<strong class="price" id="Td_Amount"><%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10,2)%></strong>
			</div>
		</div>
		</form>
<%
  end if
end if
%>

<script language="JavaScript" type="text/JavaScript">
// 函数定义在免邮费比较后面
function UpdateTotalAmount(){
if (document.shoppingcart.Post.value!=0){
  if (document.shoppingcart.Post.value==1)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei1,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei1,2)%>";
    document.shoppingcart.Pei.value="<%=enpei1%>";
    document.shoppingcart.Fei.value="<%=enfei1%>";
	document.getElementById("Tr_Amount").style.display="";
  }
  if (document.shoppingcart.Post.value==2)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei2,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei2,2)%>";
    document.shoppingcart.Pei.value="<%=enpei2%>";
    document.shoppingcart.Fei.value="<%=enfei2%>";
	document.getElementById("Tr_Amount").style.display="none";
  }
  if (document.shoppingcart.Post.value==3)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei3,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei3,2)%>";
    document.shoppingcart.Pei.value="<%=enpei3%>";
    document.shoppingcart.Fei.value="<%=enfei3%>";
	document.getElementById("Tr_Amount").style.display="none";
  }
  if (document.shoppingcart.Post.value==4)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei4,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei4,2)%>";
    document.shoppingcart.Pei.value="<%=enpei4%>";
    document.shoppingcart.Fei.value="<%=enfei4%>";
	document.getElementById("Tr_Amount").style.display="none";
  }
  if (document.shoppingcart.Post.value==5)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei5,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei5,2)%>";
    document.shoppingcart.Pei.value="<%=enpei5%>";
    document.shoppingcart.Fei.value="<%=enfei5%>";
	document.getElementById("Tr_Amount").style.display="none";
  }
  if (document.shoppingcart.Post.value==6)
  {
    document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10+enfei6,2)%>";
    document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10+enfei6,2)%>";
    document.shoppingcart.Pei.value="<%=enpei6%>";
    document.shoppingcart.Fei.value="<%=enfei6%>";
	document.getElementById("Tr_Amount").style.display="none";
  }
} else {
  document.getElementById("Td_Amount").innerHTML="<%=englobalpriceunit%>&nbsp;<%=FormatNumber(Sum*userkou/10,2)%>";
  document.shoppingcart.Total.value="<%=FormatNumber(Sum*userkou/10,2)%>";
  document.shoppingcart.Pei.value="";
  document.shoppingcart.Fei.value="";
  document.getElementById("Tr_Amount").style.display="none";
}
}
// 修改或删除购物车信息后，页面刷新时需要调用
UpdateTotalAmount(); 
</script>
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
