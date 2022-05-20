<!--#include file="goldweb_text.asp"-->
<!--#include file="../include/goldweb_shop_30_conn.asp"-->
<%
buylist = request.cookies("goldweb")("cart")
buyid = Request("Prodid")
if buyid="" Then
	conn.close
	set conn=Nothing
	response.write "<script language='javascript'>"
	response.write "alert('Please choose information first.');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
Else
	If buyid="00001" Then
	'判断ProdId="00001"，升级永久VIP会员
		If request.cookies("goldweb")("userid")="" Then
			'判断会员未登陆，跳转到登陆页面
			conn.close
			set conn=Nothing
			response.write "<script language='javascript'>"
			response.write "alert('Please login in first.');"
			response.write "location.href='login.asp';"							
			response.write "</script>"	
			response.End
		Else
			UserType=conn.execute("select UserType from goldweb_user where UserId='"&request.cookies("goldweb")("userid")&"'")("UserType")
			If UserType<>"1" Then 
				'判断会员类型<>"1"，跳转到会员中心页面
				conn.close
				set conn=Nothing
				response.write "<script language='javascript'>"
				response.write "location.href='user_center.asp';"							
				response.write "</script>"	
				response.End
			Else
				'cart只保留00001，页面跳转到shoppingcart.asp
				buylist = "'00001', '1'"
				conn.close
				set conn=Nothing
				response.Cookies("goldweb").path="/"
				response.cookies("goldweb")("cart") = buylist
				response.redirect "shoppingcart.asp"
			End If 
		End If 
	ElseIf InStr(buylist, "00001")>0 Then
	'判断"00001"是否包含在cart中，防止选择升级会员后继续添加其他信息，去掉00001添加新信息
		buylist = "'" & buyid & "', '1'"
	Else
	'继续执行添加其他信息到cart
		If Len(buylist) = 0 Then
			buylist = "'" & buyid & "', '1'"
		ElseIf InStr(buylist, buyid) <= 0 Then
			buylist = buylist & ", '" & buyid & "', '1'"
		End If
	End If 
End If

conn.close
set conn=Nothing
response.Cookies("goldweb").path="/"
response.cookies("goldweb")("cart") = buylist
response.write "<script language='javascript'>"
response.write "location.href=document.referrer;"	
response.write "</script>"	
response.end
%>
