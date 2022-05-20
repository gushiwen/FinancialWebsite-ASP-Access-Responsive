<%
' 在线支付成功后更新数据库
Sub UpdateDBAfterOnlinePayment(OrderNumber_Paid,PayType_Paid)

		sql="select * from goldweb_OrderList where OrderNum='"&OrderNumber_Paid&"'"
		set rs=Server.Createobject("ADODB.RecordSet")
		rs.Open sql,conn,1,3

		Recname = rs("Recname")
		Recmail = rs("Recmail")
		TempMailContent = ""

		'Add payment status updation to system message
		set rs2 = conn.execute("select PayType, enPayTypeDefine from pay_type Order by PayType asc")
		Do While Not rs2.eof	
			if rs2("PayType") = rs("PayType") then enPayTypeDefine1=rs2("enPayTypeDefine")
			if rs2("PayType") = PayType_Paid  then enPayTypeDefine2=rs2("enPayTypeDefine")
			rs2.movenext
		loop
		rs2.close
		set rs2=Nothing
		rs("Memo")=rs("Memo") & DateAdd("h", TimeOffset, now()) & " (Web) Payment Status from """ & enPayTypeDefine1 & """ to """ & enPayTypeDefine2 & """;" & vbCrLf
		rs("PayType")=PayType_Paid
		TempMailContent = TempMailContent & "Payment status has been set to """ & enPayTypeDefine2 & """<br>" 

		'Add order status updation to system message
		set rs2 = conn.execute("select Status, enStatusDefine from order_type Order by Status asc")
		Do While Not rs2.eof	
			if rs2("Status") = rs("Status")  then enStatusDefine1=rs2("enStatusDefine")
			if rs2("Status") = "99"  then enStatusDefine2=rs2("enStatusDefine")
			rs2.movenext
		loop
		rs2.close
		set rs2=Nothing
		rs("Memo")=rs("Memo") & DateAdd("h", TimeOffset, now()) & " (Web) Order Status from """ & enStatusDefine1 & """ to """ & enStatusDefine2 & """;" & vbCrLf
		rs("Status")="99"
		TempMailContent = TempMailContent & "Order Status has been set to """ & enStatusDefine2 & """<br>" 

		'产品加进Purchased，订阅数量加1
		buylist = ""

		'Add the product quantity accordingly
		set rs2 = conn.execute("select ProdId, ProdUnit from goldweb_order where OrderNum='"&OrderNumber_Paid&"' order by ID")
		Do While Not rs2.eof	
			If buylist="" Then
				buylist = rs2("ProdId")
			Else
				buylist = buylist & ", " & rs2("ProdId")
			End If 
			conn.execute("update goldweb_product set Quantity=Quantity+"&CInt(rs2("ProdUnit"))&" where ProdId='"&rs2("ProdId")&"'")
			rs2.movenext
			if rs2.eof then	exit do
		loop
		rs2.close
		set rs2=Nothing

		'Add ProdId into Purchased
		set rs2=Server.Createobject("ADODB.RecordSet")
		sql2="select * from goldweb_user where UserId='"&rs("UserId")&"'"
		rs2.Open sql2,conn,1,3
		if not (rs2.eof and rs2.bof) then 
			Purchased=rs2("Purchased")
			if isnull(Purchased)=true then Purchased=""

			'Add ProdId into gold_user Purchased '00021', '00022', '00023'
			buyid = Split(buylist, ", ")
			For I=0 To UBound(buyid)
			if Purchased="" then
				Purchased = "'"&buyid(I)&"'"
			ElseIf InStr(Purchased, "'"&buyid(i)&"'") <= 0 Then 'Purchased不包含新订单商品
				Purchased = Purchased & ", '" & buyid(i) &"'"
			ElseIf InStr(Purchased, ", '"&buyid(i)&"'") > 0 Then '第2个开始商品相同   
				Purchased = Replace(Purchased, ", '" & buyid(i) & "'", "")
				Purchased = Purchased & ", '" & buyid(i) &"'"
			ElseIf InStr(Purchased, "'"&buyid(i)&"', ") = 1 Then 'Purchased有2个以上商品，第1个商品相同
				Purchased = Replace(Purchased, "'" & buyid(i) & "', ", "")
				Purchased = Purchased & ", '" & buyid(i) &"'"
			End If  'Purchased有1个商品，若相同无需替换
			Next

			rs2("Purchased") = Purchased
			'Add order amount to user total amount
			rs2("totalsum") = CDbl(rs2("totalsum")) + CDbl(rs("OrderSum"))
			rs2("jifen") = rs2("jifen") + CInt(rs("OrderSum"))
			rs2.update
		End If 
		rs2.close
		set rs2=Nothing

		rs.update
		rs.close
		set rs=Nothing


		'Check whether Upgrade to Permanent VIP Member
		set rs = conn.execute("select * from goldweb_Order where OrderNum='"&OrderNumber_Paid&"' and ProdId='00001'")
		If Not rs.eof then 
			'Update user details
			sql2="select * from goldweb_user where UserId='"&rs("UserId")&"'"
			set rs2=Server.Createobject("ADODB.RecordSet")
			rs2.Open sql2,conn,1,3

			If rs2("UserType")<>"4" Then 
				If rs2("UserType")="1" Then UserTypeText1=enusertype1
				If rs2("UserType")="2" Then UserTypeText1=enusertype2
				If rs2("UserType")="3" Then UserTypeText1=enusertype3
				'If rs2("UserType")="4" Then UserTypeText1=enusertype4
				If rs2("UserType")="5" Then UserTypeText1=enusertype5
				If rs2("UserType")="6" Then UserTypeText1=enusertype6
				UserTypeText2=enusertype4
				rs2("Memo2")  = rs2("Memo2") & DateAdd("h", TimeOffset, now()) & " (Web) UserType from """ & UserTypeText1  & """ to """ & UserTypeText2 & """;" & vbCrLf
				TempMailContent = TempMailContent & "User type has been set to """ & UserTypeText2 & """<br>" 
			End If 

			If FormatNumber(rs2("userkou"),2)<>9.00 Then 
				rs2("Memo2")  = rs2("Memo2") & DateAdd("h", TimeOffset, now()) & " (Web) UserKou from """ & FormatNumber(rs2("userkou"),2)  & """ to """ & "9.00" & """;" & vbCrLf
				TempMailContent = TempMailContent & "User discount has been set to ""-10%""<br>" 
			End If 

			rs2("UserType")="4"
			rs2("UserKou")="9"
			rs2.update
			rs2.close
			set rs2=Nothing

		End If 
		rs.close
		set rs=Nothing
		
		If TempMailContent <> "" Then 
			dim MailTitle
			dim MailContent
			MailTitle = "Updates for Online Order ("&OrderNumber_Paid&")"&" - From "&siteurl
			MailContent = "Dear " & Recname & ",<br><br>Please refer to the updates below for Online Order ("&OrderNumber_Paid&").<br><br>"
			MailContent = MailContent & TempMailContent & "<br><br>"
			MailContent = MailContent & "Your sincerely<br>"&ensitename&"<br>"&"Tel: "&adm_tel&"<br>"&FullSiteUrl

			Call SendOutMail(MailTitle, MailContent, adm_mail, Recmail, adm_mail)
		End If 
End Sub	
%>
