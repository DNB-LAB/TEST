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
to_year = year(date())

start_year = Cint(to_year)-3
end_year = Cint(to_year)+10

stdate=date()

%>
<%
to_year = year(date())

start_year = Cint(to_year)-3
end_year = Cint(to_year)+10

Fdate=date()

%>

<html>
<head>
<title>2. 시장출하 적부 판정 기록 기간별 조회 [ 제품명]</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="content-type" content="text/html; charset=utf-8">

<script language="javascript">
function Check() {
form.target = 'newW'; //타겟지정, 아래에는 url을 적지않는다
window.open ;
}
</script>

</head>
<body bgcolor="#D7F1FA">
    <table border="1" width="800" align="center" cellspacing="0">
        <form method=get action="02_P_name_ok.asp?<%=Var5%>"  name="form" onSubmit="Check();return true;">
        <tr height="30">
            <td width="600"  bgcolor="#FFACB7" align="center">
          <font face="돋움"><b>2.  시장출하 적부 판정 기록 기간별 조회 [ 제품명]</td>
            <td width="200" bgcolor="#FFACB7"  align="center">
               <span style="font-size:10pt;"><b>조&nbsp;회</b></span></td>
        </tr>
        <tr>
            <td  height="80" align=center bgcolor="#FFFFFF">
&nbsp;
입고일 : <select name="Styear" id="Styear" style="width:70">
                <% for k=start_year to end_year %>
                <option value="<%=k%>"  <%if k=year(stdate) then%> selected <% end if %>><%=k%>년</option>
								<% next%>
      </select>
              <select name="Stmonth" id="Stmonth"  style="width:60">
							  <% for k=1 to 12 %>
                <option value="<%=k%>"  <%if k=month(stdate) then%> selected <% end if %>><%=k%>월</option>
								<% next%>
              </select>
              <select name="Stday" id="Stday"  style="width:60">
							  <% for k=1 to 31 %>
                <option value="<%=k%>"  <%if k=day(stdate) then%> selected <% end if %>><%=k%>일</option>
	<% next%> </select> ~&nbsp;&nbsp;
	
	
	
	<select name="Fyear" id="Fyear" style="width:70">
                <% for k=start_year to end_year %>
                <option value="<%=k%>"  <%if k=year(Fdate) then%> selected <% end if %>><%=k%>년</option>
								<% next%>
      </select>
              <select name="Fmonth" id="Fmonth"  style="width:60">
							  <% for k=1 to 12 %>
                <option value="<%=k%>"  <%if k=month(Fdate) then%> selected <% end if %>><%=k%>월</option>
								<% next%>
              </select>
              <select name="Fday" id="Fday"  style="width:60">
							  <% for k=1 to 31 %>
                <option value="<%=k%>"  <%if k=day(Fdate) then%> selected <% end if %>><%=k%>일</option>
	<% next%> </select>
	
	
	<br><br style=line-height:5px;><b>제품명 : <input type="text" size=40 name="Product_Name_DZ"></center>
</td>
           <td  height="45"  align=center bgcolor="#FFFFFF">
                <span style="font-size:10pt;" ><input type="submit" value="조 회">
                </span> </td>
        </tr>
      
        
</table>
</body>
</form>
</html>
