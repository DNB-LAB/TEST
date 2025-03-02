<%@Language="VBScript" CODEPAGE="65001" %>
<% 
Response.CharSet="utf-8" 
Session.codepage="65001" 
Response.codepage="65001" 
Response.ContentType="text/html;charset=utf-8" 
%>

<!DOCTYPE HTML>

<!-- #include file="inc.asp" -->
<%
if session("id")="" then  
'로그인하여 얻은 세션(id)가 없으면 로그인으로 돌려 보내고 있으면 리스트를 보여준다.
%>

<html>
<head>
<body leftmargin="0" topmargin="0" bgcolor="#D7F1FA">
<script language="javascript">
		alert("로그인이 필요합니다. 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../../Log_in.asp','end','width=310,height=190,top=270, left=350');
</script>

<% else %>
<%
'삭제할 글번호를 전송받는다
Sid = Request.QueryString("sid")


'사용자가 입력한 비밀번호를 전송받는다
Pass = Request.Form("pass")

'디비연결
Set DB = Server.CreateObject("ADODB.Connection")
DB.open ConnString

'데이타베이스의 비밀번호를 가져온다
SQL = " SELECT top 1 pass FROM AL_021_Judge_System" 
SQL = SQL & " WHERE sid =" & sid
Set RS = DB.Execute(SQL)

OldPass=Rs("pass")


RS.Close
Set RS=nothing


'두개의 비밀번호를 확인한다
IF (Pass = OldPass) or (Pass=adminpass) then  'adminpass는 관리자용==> inc.asp에 정의********************************

    Set DB = Server.Createobject("ADODB.Connection")
    DB.open ConnString

   SQL = "DELETE from AL_021_Judge_System" 
   SQL = SQL & " WHERE sid =" & sid
   DB.Execute(SQL)
   
    
    
    SQL = "DELETE from AL_022_Judge_Warehouse" 
    SQL = SQL & " WHERE original_sid =" & sid
    DB.Execute(SQL)
    
    



URL = "list.asp?" & Var2
Response.Redirect URL


Else
%>
<body bgcolor="#D7F1FA">
<script language="javascript">
alert("암호가 틀립니다.-----");
history.back();
</script>
<%
End if
DB.Close
Set DB=nothing
%>
 <% end if %>