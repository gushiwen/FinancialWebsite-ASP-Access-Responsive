<div  class="sidebar-content">
<!-- Start_Module_63269 -->
<div class="module module-hlbg">
<div class="module-inner">
<div class="module-hlbg-title clearfix"><h3>CATEGORY</h3></div>
<div  class="module-hlbg-content">
<!-- Start_Module_63269 -->

<!-- S news classify -->
<div class="category category-p product-category">
<ul>
<% ' 新闻分类
For i=1 To 5
		If newstitle(i)<>"" then
%>
	<li ><a href="newslist.asp?ClassId=<%=i%>"><%=newstitle(i)%></a></li>
<%
		End if
	next
%>
</ul>
</div>
<!-- E news classify -->

<!-- End_Module_63269 -->
</div></div></div></div>
