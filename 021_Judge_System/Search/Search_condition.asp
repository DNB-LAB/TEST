<%@Language="VBScript" CODEPAGE="65001" %>
<% 
Response.CharSet="utf-8" 
Session.codepage="65001" 
Response.codepage="65001" 
Response.ContentType="text/html;charset=utf-8" 
%>
<!DOCTYPE HTML>

<!-- #include file="inc.asp" -->

<html>
<title>시장출하 적부 판정 기록 자료</title>
<link rel="stylesheet" href="basic.css">
<body bgcolor="#D7F1FA">
<meta http-equiv="content-type" content="text/html; charset=utf-8">
</head>

<body bgcolor="#D7F1FA">

    <table border="0" width="810" align="center"  bgcolor="#D7F1FA">
      <tr><td align=right  bgcolor="#D7F1FA">
      <a href="javascript:history.go(-1)"><img src=../images/back.gif border=0></a>&nbsp;&nbsp;
      <a href="../list.asp"><img src=../images/list.gif border=0></a>&nbsp;&nbsp;</td>
      </tr>
    <tr>
         <td  valign="top" bgcolor="#D7F1FA">
         <iframe src="01_P_Code.asp" frameborder="0" width="810" height=120 marginwidth="1" marginheight="1" name="1" scrolling="no" allowtransparency="true"></iframe>
        </td>
    </tr>
     <tr>
         <td  valign="top" bgcolor="#D7F1FA">
         <iframe src="02_P_name.asp" frameborder="0" width="810" height=120 marginwidth="1" marginheight="1" name="1" scrolling="no" allowtransparency="true"></iframe>
        </td>
    </tr>
   
	<tr>
         <td  valign="top" bgcolor="#D7F1FA">
         <iframe src="03_Maker.asp" frameborder="0" width="810" height=120 marginwidth="1" marginheight="1" name="1" scrolling="no" allowtransparency="true"></iframe>
        </td>
    </tr>
    <tr>
         <td  valign="top" bgcolor="#D7F1FA">
         <iframe src="04_Warehouse.asp" frameborder="0" width="810" height=120 marginwidth="1" marginheight="1" name="1" scrolling="no" allowtransparency="true"></iframe>
        </td>
    </tr>
   <tr>
         <td  valign="top" bgcolor="#D7F1FA">
         <iframe src="05_Certi.asp" frameborder="0" width="810" height=120 marginwidth="1" marginheight="1" name="1" scrolling="no" allowtransparency="true"></iframe>
        </td>
    </tr>
    
     <tr>
         <td  valign="top" bgcolor="#D7F1FA">
         <iframe src="06_Note.asp" frameborder="0" width="810" height=120 marginwidth="1" marginheight="1" name="1" scrolling="no" allowtransparency="true"></iframe>
        </td>
    </tr>
        
</table>

<br><br><br><br><br><br>
</body>
</form>
</html>
