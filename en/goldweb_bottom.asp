<%
Set rsp = conn.Execute("select * from page") 
bottomGuidTitle = "Organization; Newsletter; Cooperation; Membership; Help Desk"
'body99=rsp("body99")
%>

<footer class="footer">
<div class="footer-main">
<div  class="page-width clearfix">

<!-- Start_Module_63238 -->
<div class="module-default">
<div class="module-inner">
<div  class="module-content">
<div class="qhd-module">
<div class="column">

<%
for lie=1 to 5
%>
<div class="col-5-1 <%If lie=5 Then response.write "last"%>">
<div  class="qhd_column_contain"><!-- Start_Module_63224 -->
<div class="module-default">
<div class="module-inner">
<div class="module-title module-title-default clearfix"><div class="module-title-content clearfix"><h3 style="" class=""><%= Split(bottomGuidTitle, ";")(lie-1)%></h3></div></div>
<div  class="module-content"><!-- Start_Module_63224 -->
<!-- S link-line -->
<div class="link link-block">
<ul>
<%
	for hang=1 to help_hang
		ID=lie*10+hang
		ID=cstr(id)
		title="entitle"&id
		url="url"&id
		'body="enbody"&id
		if rsp(url)<>"" then
			response.write "<li><a href='"&rsp(url)&"' target=''><span>"&rsp(title)&"</span></a></li>"
		else
			response.write "<li><a href='page.asp?id="&id&"' target=''><span>"&rsp(title)&"</span></a></li>"
		end if
	next
%>
</ul>
</div>
<!-- E link-line --><!-- End_Module_63224 -->
</div></div></div></div></div>
<%
next
%>

</div></div></div></div>
<div class="module-divider"></div>
</div>
<!-- End_Module_63238 -->

<!-- Start_Module_63239 -->
<div class="module-default">
<div class="module-inner">
<div  class="module-content"><!-- Start_Module_63239 -->
<div class="qhd-module">
<div class="column">

<div class="col-5-3">
<div  class="qhd_column_contain"><!-- Start_Module_63228 -->
<div class="module-default module-no-margin">
<div class="module-inner">
<div  class="module-content"><!-- Start_Module_63228 -->
<div class="qhd-content">
<p><%= enadm_comp%> </p>
<p>Tel: <%= adm_tel%>&nbsp;&nbsp;Email: <%= adm_mail%> </p>
</div><!-- End_Module_63228 -->
</div></div></div></div></div>

<div class="col-5-2 last">
<div  class="qhd_column_contain"><!-- Start_Module_63229 -->
<div class="module-default module-no-margin">
<div class="module-inner">
<div  class="module-content"><!-- Start_Module_63229 -->
<div class="qhd-content">
<p style="text-align: right;"><a style="display:none" href="www.ol.sg">Supported by www.OL.sg</a></p>
</div><!-- End_Module_63229 -->
</div></div></div></div></div>

</div></div><!-- End_Module_63239 -->
</div></div></div>


</div></div>
</footer>

<!--统计访问量开始-->
<script type="text/javascript">
document.write('<scr' + 'ipt type="text/javascript" src="../include/goldweb_tj.asp?where=' + URLencode(document.referrer) + '"><' + '/script>');
</script>
<!--统计访问量结束-->
