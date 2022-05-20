<div style="text-align:center; margin-bottom:30px;" class="desktops-section">
	<form name="desktopsearch" action="productsearch.asp" method="post">
		<input name="action" type="hidden" value="topsearch">
		<input name="keywords" value="<%=keywords%>" type="text" placeholder="Search Information" autocomplete="off" style="padding-left:4px; width:80%; height:42px; line-height:42px; BORDER:#CCCCCF 1px solid; COLOR: #666666;">
		<a href="javascript:;" onclick="document.desktopsearch.submit();" style="display:inline-block; width:10%; padding:0 20px; line-height:42px; background-color:#e5e5e5; font-size:16px; font-size:1.6rem; color:#808080; text-decoration:none; text-align:center; cursor:pointer; transition:all 0.5s ease 0s; border-radius:2px; " onMouseOver="this.style.backgroundColor='#d9d9d9';" onMouseOut="this.style.backgroundColor='#e5e5e5';" >GO</a>
	</form>
</div>
