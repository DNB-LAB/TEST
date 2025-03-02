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
'내용을 볼 글번호를 전송받는다
Sid= Request.QueryString("sid")
Var3 = Var2 & "&sid=" & Sid

'조회수를 1증가시킨다
Set DB = Server.Createobject("ADODB.connection")
DB.open Connstring

'내용을 볼 레코드의 각 필드값을 가져온다
SQL = "SELECT * FROM " & STable
SQL = SQL & " WHERE sid =" & Sid
Set RS = DB.Execute(SQL)

'한꺼번에 모두 가져와서 변수에 대입
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	Lot_number=RS("Lot_number")
	
	Expiration_Date=RS("Expiration_Date")


RS.Close
Set RS=nothing

DB.Close
Set DB=nothing

%>
<html>
<head>
<title> 적합 로트 리스트(20/01~22/03)</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="basic.css">
</head>
<body bgcolor="#E4F7BA">
  <br><br style="line-height:20px;">
      <table border=0 cellspacing=0 cellpadding=0 width="1024" align=center  style="table-layout:fixed;">
    <tr>
	<td bgcolor="#E4F7BA"  width="512" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:10px; padding-bottom:5px; padding-left:0px;">
         	<b>☞ 선택 제품 로트 및 사용기한 정보</td>
	  <td bgcolor="#E4F7BA" style="text-align:right; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:10px;">
    <td align="right"  width="512" bgcolor="#E4F7BA">
      <a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0"></a>
      &nbsp;
      <a href="list.asp?<%=Var3%>"><img src=images/list.gif border=0></a></td>
    </tr>
        </table>
        

<table border=1 cellspacing=0 cellpadding=0 width="1024"  bgcolor="#FFFFFF" align=center  style="table-layout:fixed;">
      <tr>
       <td  width=120 bgcolor="#F0B6B6" style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <b>제품코드</td> 
       <td width="221" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
                  <%=Product_Code%>&nbsp;</td>
   <td width="120"bgcolor="#F0B6B6" style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
        
           <b>제 품 명</b></font></td>
        <td width="563" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
             <%=Product_Name_DZ%></td>
								
   
      </tr>
       <tr>
        <td bgcolor="#F0B6B6" style="text-align:center; text-indent:0; margin:0; padding-top:15px; padding-right:0px; padding-bottom:15px; padding-left:0px;">
       <b>로트 번호</b></font></td>
              <td colspan=3 style="text-align:left; text-indent:0; margin:0; padding-top:7px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
              <%=Lot_number%>&nbsp;</td>
        </tr>
        <tr>
        <td bgcolor="#F0B6B6" style="text-align:center; text-indent:0; margin:0; padding-top:15px; padding-right:0px; padding-bottom:15px; padding-left:0px;">
       <b>사용 기한</b></font></td>
              <td colspan=3 style="text-align:left; text-indent:0; margin:0; padding-top:7px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
              <%=Expiration_Date%>&nbsp;</td>
        </tr>
     
      </table>  
     <table align="center" width=1024 align="center"  style='table-layout:fixed;'>
         <tr height=50>
             <td  align=center>
             <a href= "javascript:window.open('about:blank', '_self').close();">
             <img src="images/close.gif" border="0"></a></td>
          </tr>
        </table>


  <br><br style="line-height:50px;">
</body>
</html>
