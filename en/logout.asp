<!--#include file="goldweb_text.asp"-->
<% ' 退出登录
Session.abandon
Response.cookies("goldweb").path="/"
response.cookies("goldweb")("userid")=""
response.cookies("goldweb")("userkou")=""
response.cookies("goldweb")("cart")=""
response.cookies("goldweb")("search")=""

response.write "<script language='javascript'>"
response.write "alert('You have logged out sucessfully.');"
response.write "location.href='index.asp';"
response.write "</script>"
response.end
%>
