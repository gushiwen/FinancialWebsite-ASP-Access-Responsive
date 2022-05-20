<div  class="sidebar-content">
<!-- Start_Module_63269 -->
<div class="module module-hlbg">
<div class="module-inner">
<div class="module-hlbg-title clearfix"><h3>CATEGORY</h3></div>
<div  class="module-hlbg-content">
<!-- Start_Module_63269 -->

<!-- S product-category -->
<div class="category category-p product-category">
<ul>
<% ' Information分类
'if LarCode="" then 
	sqlL="select * from goldweb_class where MidSeq=1 order by larseq"
'Else
	'sqlL="select * from goldweb_class where enLarCode='"&LarCode&"' and MidSeq=1"
'End If 
Set rsL=Server.CreateObject("ADODB.Recordset")
rsL.open sqlL,conn,1,3
do while not rsL.eof
%>
	<li class="<% If LarCode = rsL("enLarCode") Then response.write "current" %>" ><a href="productlist.asp?LarCode=<%=Server.URLEncode(rsL("enLarCode"))%>&ClassId=<%=rsL("ClassId")%>"><%=rsL("enLarCode")%></a><i></i>
	</li>
<%
    rsL.movenext
    loop
    rsL.close
    set rsL=Nothing
%>
</ul>
</div>
<script type="text/javascript">
 $(document).ready(function(){
  $('.category-p ul ul').find('li:last').addClass('last');
  
  $('.category-p > ul > li > a').click(function(){
   if( $(this).parent('li').find('ul') ){
    $(this).parent('li').find('ul').slideDown('fast');
    $(this).parent('li').siblings('li').find('ul').slideUp('fast');
    $(this).parent('li').addClass('current').siblings('li').removeClass('current');
   }
  }); 
  
 });
</script>
<!-- E product-category -->

<!-- End_Module_63269 -->
</div></div></div></div>