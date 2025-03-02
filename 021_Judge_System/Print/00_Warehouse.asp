<!DOCTYPE html>
<% @LANGUAGE='VBSCRIPT' CODEPAGE='949' %> 
<% session.codepage = 949 %>
<% Response.ChaRset = "euc-kr" %>

<!-- #include file="inc.asp" -->

<%
to_year = year(date())

start_year = Cint(to_year)-20
end_year = Cint(to_year)+1

sdate=date()

%>

<html>
<head>
<title>시장출하 적부 판정 기록</title>
<link rel="stylesheet" href="basic.css">
<meta name="author" content="이청희">

<script language="javascript">
function Check() {
form.target = 'newW'; //타겟지정, 아래에는 url을 적지않는다
window.open ;
}
</script>

</head>
<body bgcolor="#D7F1FA">
    <table border="1" width="800" align="center" cellspacing="0">
        <form method=get action="00_Warehouse_Ok.asp?<%=Var5%>"  name="form" onSubmit="Check();return true;">
        <tr height="30">
            <td width="600"  bgcolor="#FF9999" align="center">
          <font face="돋움"><b>시장출하 적부 판정 기록 일지</td>
            <td width="200" bgcolor="#FF9999"  align="center">
                <font face="돋움"><b>조 회 </td>
        </tr>
        <tr height="50">
            <td  align=center  bgcolor="#FFFFFF">
&nbsp;
<select name="Syear" id="Syear" style="width:80">
                <% for k=start_year to end_year %>
                <option value="<%=k%>"  <%if k=year(sdate) then%> selected <% end if %>><%=k%>년</option>
								<% next%>
      </select>
              <select name="Smonth" id="Smonth"  style="width:60">
							  <% for k=1 to 12 %>
                <option value="<%=k%>"  <%if k=month(sdate) then%> selected <% end if %>><%=k%>월</option>
								<% next%>
              </select>
              <select name="Sday" id="Sday"  style="width:60">
							  <% for k=1 to 31 %>
                <option value="<%=k%>"  <%if k=day(sdate) then%> selected <% end if %>><%=k%>일</option>
	<% next%> </select>
	&nbsp;&nbsp;&nbsp;&nbsp;판정결과: &nbsp;
	<select name="Judge_Result">
                                     <option value="">전체</option>
                                     <option value="적합">적합</option>
                                     <option value="부적합">부적합</option>
                                     <option value="보류">보류</option>
                                     <option value="기입고">기입고</option>
                                     </select></td>
            <td  align=center  bgcolor="#FFFFFF">
                  <span style="font-size:10pt;" ><input type="submit" value="조 회"></span></td>
        </tr>
</table>
 
</body>
</form>
</html>
