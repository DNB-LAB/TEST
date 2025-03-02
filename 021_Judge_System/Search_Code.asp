<%@Language="VBScript" CODEPAGE="65001" %>
<% 
Response.CharSet="utf-8" 
Session.codepage="65001" 
Response.codepage="65001" 
Response.ContentType="text/html;charset=utf-8" 
%>

<!DOCTYPE HTML>

<html>
<head>
<title>시장출하 적부 판정 기록을  위한 제품 코드 검색</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="special.css">
<script language="JavaScript">
function send()
{
if (document.form.Product_Code.value.length==0) 
{
alert("제품 코드를 입력하세요.");
document.form.Product_Code.focus();
return false;
}
}
</script>



<body bgcolor="#D7F1FA" align="center" topmargin=100 onLoad="document.form.Product_Code.focus();">
 <table Cellspacing=0 cellpadding=0 width="600" border="0" align="center">
    <tr height="30"> 
      <td width=300 bgcolor="#D7F1FA"></td>
      <td width=300 bgcolor="#D7F1FA" style="text-align:right; text-indent:0; margin:0; padding-top:0px; padding-right:0px; padding-bottom:0px; padding-left:5px; border-width:0pt; border-color:white; border-style:none;"> 
            <a href="javascript:history.go(-1)"><img src="images/back.gif" border="0"></a>&nbsp;
            <a href="list.asp?<%=Var3%>"><a href="list.asp"><img src="images/list.gif" border="0"></a></td>
    </tr>
  </table>
<table cellspacing=0 cellpadding=0 border=1 align="center" width="600">
      <tr  height="40"> 
        <th align="center">
          <b><font color=black>제품 코드 검색</b></th>
      </tr>
<form name=form action="Search_Code_ok.asp?Product_Code=<%=Product_Code%>" method=get onSubmit="return send()">
     <tr align="center" height="80">
      <td>
      <table cellspacing=0 cellpadding=0 border=0 align="center" width="490">
       <tr align="center" height="80">
      <td width=400  style="text-align:center; text-indent:0; margin:0; padding-top:10px; padding-right:0px; padding-bottom:8px; padding-left:0px;">
       <input name="Product_Code" type="text" size=25  maxsize=30>
      </td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:7px; padding-right:0px; padding-bottom:5px; padding-left:10px;">
    <input type="image" img src="images/Choice.gif" border="0"></td>
     </tr>
     </table>
     
  </form>

</td>
</tr>
</table>
<br><br><br><br><br><br><br><br><br><br>
</body>
</html>

