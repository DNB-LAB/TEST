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
<title>시장출하 적부 판정 기록을 위한 제품명 검색 </title>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="special_name.css">
<script language="JavaScript">
function send()
{
if (document.form.Product_Name_DZ.value.length==0) 
{
alert("제품명을 입력하세요.");
document.form.P_name.focus();
return false;
}
}
</script>
</head>
<body bgcolor="#D7F1FA" align="center" topmargin=30 onLoad="document.form.P_name.focus();">
 <table Cellspacing=0 cellpadding=0 width="800" border="0" align="center">
    <tr height="40"> 
      <td width=400 bgcolor="#D7F1FA"></td>
      <td width=400 bgcolor="#D7F1FA" style="text-align:right; text-indent:0; margin:0; padding-top:0px; padding-right:0px; padding-bottom:0px; padding-left:5px; border-width:0pt; border-color:white; border-style:none;"> 
            <a href="javascript:history.go(-1)"><img src="images/back.gif" border="0"></a>&nbsp;
             <a href="list.asp?<%=Var3%>"><a href="list.asp"><img src="images/list.gif" border="0"></a></td>
    </tr>
  </table>
 
<table cellspacing=0 cellpadding=0 border=1 align="center" width="800">
      <tr align="center" height="40">
        <td align="center" bgcolor="#47FF9C" >
          <b><font color=black>제품명 검색</b></th>
      </tr>
       <form name=form action="Search_Name_ok.asp?P_name=<%=P_name%>" method=get onSubmit="return send()">

     <tr align="center" height="80">
      <td>
      <table cellspacing=0 cellpadding=0 border=0 align="center" width="690">
       <tr align="center" height="80">
      <td width=600 align=center >
       <input name="Product_Name_DZ" type="text" size=72  maxsize=100>
      </td>
        <td   style="text-align:left; text-indent:0; margin:0; padding-top:7px; padding-right:0px; padding-bottom:5px; padding-left:10px;">
  <input type="image" img src="images/Choice.gif" border="0"></td>
    </tr>
     </table>
     
  </form>

</td>
</tr>
</table>
<br><br><br><br><br>
</body>
</html>

