<html>
<head>
<title>위,수탁제품 및 외주가공 제품, 기타 제품 시험 의뢰를 위한 제품 코드 입력</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="special.css">
<script language="JavaScript">
function send()
{
if (document.form.P_code.value.length==0) 
{
alert("제품 코드를 입력하세요.");
document.form.P_code.focus();
return false;
}
}
</script>
</head>
<body bgcolor="#D7F1FA" align="center" topmargin=50 onLoad="document.form.p_code.focus();">
  <table Cellspacing=0 cellpadding=0 width="500" border="0" align="center">
    <tr height="30"> 
      <td width=200 bgcolor="#D7F1FA"></td>
      <td width=300 bgcolor="#D7F1FA" style="text-align:right; text-indent:0; margin:0; padding-top:0px; padding-right:0px; padding-bottom:0px; padding-left:5px; border-width:0pt; border-color:white; border-style:none;"> 
            <a href="javascript:history.go(-1)"><img src="images/back.gif" border="0"></a>&nbsp;
             <a href="list.asp?<%=Var3%>"><a href="list.asp"><img src="images/list.gif" border="0"></a></td>
    </tr>
  </table>
<table cellspacing=0 cellpadding=0 border=1 align="center" width="500">
      <tr  height="30"> 
        <th align="center">
          <b><font color=black>위,수탁제품 및 외주가공 제품, 기타 제품 시험 의뢰를 위한 제품 코드 입력</b></th>
      </tr>
<form name=form action="Search_write_ok.asp?P_code=<%=P_code%>" method=get onSubmit="return send()">
     <tr align="center" height="80">
      <td >
      <table cellspacing=0 cellpadding=0 border=0 align="center" width="490">
       <tr align="center" height="80">
      <td width=400 align=center >
       <input name="p_code" type="text" size=13 maxlength="15" maxsize=20>
      </td>
      <td  align=left><input type="image" img src="images/Choice.gif" border="0"></td>
     </tr>
  </form>
</table>
</body>
</html>

