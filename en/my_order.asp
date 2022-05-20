<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
'call aspsql()
call checklogin()
if request("del")<>"" then call delorder()
if request("cancel")<>"" then call cancelorder()
%>

<!DOCTYPE html>
<head id="Head">
	<title>Order List-<%=ensitename%>-<%=siteurl%></title>
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
&nbsp;&nbsp;&raquo;&nbsp;&nbsp;ORDER LIST
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

.list-box { }
.list-box .list_row{width:100%; margin-top:10px; text-align:left; }
.list-box .list_row .no-item {width:100%; text-align:center; }
.list-box .list_row .btn_view{display:inline-block; cursor:pointer; width:78%; height:30px; line-height:30px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding-left:2px; border:1px solid #e6e6e6; text-align:left; border-radius:5px; }
.list-box .list_row .btn_opt{display:inline-block; cursor:pointer; width:18%; height:30px; line-height:30px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; padding:0px; border:1px solid #e6e6e6; text-align:center; border-radius:5px; }
</style>

		<div class="list-box">

<%
sql = "select A.OrderNum,A.OrderTime,A.RecName,A.OrderSum,A.RecName,A.Status,A.PayType,B.enStatusDefine,C.enPayTypeDefine "&_
          " from (goldweb_OrderList A left join order_type B on A.Status = B.Status) left join pay_type C on A.PayType = C.PayType "&_
		  " where A.del<>true and A.UserId='"&request.cookies("goldweb")("userid")&"' order by A.OrderTime desc"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,1

if rs.eof and rs.bof  then
	response.write "<div class='list_row'>You haven't made any order.</div>"
else
	pages=20
	rs.pagesize=pages
	allPages = rs.pageCount	'总页数
	page=request("page")
	If not isNumeric(page) then page=1
	if isEmpty(page) or clng(page) < 1 then
		page = 1
	elseif clng(page) >= allPages then
		page = allPages 
	end if
	rs.AbsolutePage=page

	do while not rs.eof And pages>0
%>

			<div class="list_row">
				<a class="btn_view" href="my_order_view.asp?OrderNum=<%=rs("OrderNum")%>"><%=rs("OrderNum")%>&nbsp;&nbsp;<%=DatePart("yyyy",rs("OrderTime"))%>-<%=Right("0"&DatePart("m",rs("OrderTime")),2)%>-<%=Right("0"&DatePart("d",rs("OrderTime")),2)%>&nbsp;&nbsp;<%=englobalpriceunit%><%=formatnumber(rs("OrderSum"),2)%></a>
				<%
	  				If rs("Status")="0" And rs("PayType")="0" Then 
	    				response.write "<a class=""btn_opt"" href=""my_order.asp?cancel="&rs("OrderNum")&""">Cancel</a>"
      				ElseIf rs("Status")="11" And rs("PayType")="0" then 
	    				response.write "<a class=""btn_opt"" href=""my_order.asp?cancel="&rs("OrderNum")&""">Retrieve</a>"
      				Else 
	    				response.write "<a class=""btn_opt"" href=""my_order_view.asp?OrderNum="&rs("OrderNum")&""">View</a>"
	  				End If 
				%>
			</div>

<%
		pages = pages - 1
		rs.movenext
		if rs.eof then exit do
	Loop
%>

<!-- S pagination -->
<div class="pagination pagination-default">
<%
if page = 1 then
	response.write "<span class=""page-first disabled"">First</span><span class=""page-prev disabled"">Previous</span>"
else
	response.write "<a class=""page-first"" href=""my_order.asp?page=1"">First</a><a class=""page-prev"" href=""my_order.asp?page="&page-1&""">Previous</a>"
end If

response.write "<span class=""current"">" & page & "/" & allpages & "</span>"

if page = allpages then
response.write "<span class=""page-next disabled"">Next</span><span class=""page-last disabled"">Last</span>"
else
response.write " <a class=""page-next"" href=""my_order.asp?page="&page+1&""">Next</a><a class=""page-last"" href=""my_order.asp?page="&allpages&""">Last</a>"
end if
%>
</div>
<!-- E pagination -->

<%
end If

rs.close
set rs=nothing
%>

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
sub cancelorder()
sql="select * from goldweb_OrderList where OrderNum='"&request("cancel")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
		tishi="The order no. is not exist."
	elseif rs("userid")<>request.cookies("goldweb")("userid") then
		tishi="You are not authorized to manage this order."
	Else	
		Recname = rs("Recname")
		Recmail = rs("Recmail")
		TempMailContent = ""

		'Add  order status updation to system message
	    set rs2 = conn.execute("select Status, enStatusDefine from order_type Order by Status asc")
		Do While Not rs2.eof	
		if rs2("Status") = "11" then enStatusDefine1=rs2("enStatusDefine")
		if rs2("Status") = "0"  then enStatusDefine2=rs2("enStatusDefine")
		rs2.movenext
		if rs2.eof then	exit do
		Loop
		rs2.close
	   set rs2=nothing

		if rs("Status")="11" then
		   rs("Status")="0"
		   rs("Memo")  = rs("Memo") & vbCrLf & DateAdd("h", TimeOffset, now()) & " (Web) Order Status from """ & enStatusDefine1 & """ to """ & enStatusDefine2 & """;"
		   TempMailContent = TempMailContent & "Order status has been set to """ & enStatusDefine2 & """<br>" 
		   tishi="The order has been retrieved sucessfully!"
		   rs.update
		elseif rs("Status")="0" then
		   rs("Status")="11"
		   rs("Memo")  = rs("Memo") & vbCrLf & DateAdd("h", TimeOffset, now()) & " (Web) Order Status from """ & enStatusDefine2 & """ to """ & enStatusDefine1 & """;"
		   TempMailContent = TempMailContent & "Order status has been set to """ & enStatusDefine1 & """<br>" 
		   tishi="The order has been cancelled successfully!"
		   rs.update
		else
	  	  tishi="The order can not be auto-cancelled!"
		end if
		
		If TempMailContent <> "" Then 
			dim MailTitle
			dim MailContent
			MailTitle = "Updates for Online Order ("&request("cancel")&")"&" - From "&siteurl
			MailContent = "Dear " & Recname & ",<br><br>Please refer to the updates below for Online Order ("&request("cancel")&").<br><br>"
			MailContent = MailContent & TempMailContent & "<br><br>"
			MailContent = MailContent & "Your sincerely<br>"&ensitename&"<br>"&"Tel: "&adm_tel&"<br>"&FullSiteUrl

			Call SendOutMail(MailTitle, MailContent, adm_mail, Recmail, adm_mail)
		End If 
	end if
rs.close
set rs=nothing
response.write "<script language='javascript'>"
response.write "alert('"&tishi&"');"
response.write "location.href='my_order.asp';"
response.write "</script>"
end sub

sub delorder()
sql="select * from goldweb_OrderList where OrderNum='"&request("Del")&"'"
set rs=Server.Createobject("ADODB.RecordSet")
rs.Open sql,conn,1,3

	if rs.eof and rs.bof then
		tishi="The order no. is not exist."
	elseif rs("userid")<>request.cookies("goldweb")("userid") then
		tishi="You are not authorized to manage this order."
	else
		conn.execute("update goldweb_OrderList set del=true where OrderNum='"&request("Del")&"'")
		tishi="The order has been deleted sucessfully!"
	end if
  		response.write "<script language='javascript'>"
		response.write "alert('"&tishi&"');"
		response.write "location.href='my_order.asp';"
		response.write "</script>"
rs.close
set rs=nothing
end Sub

conn.close
set conn=nothing
%>
