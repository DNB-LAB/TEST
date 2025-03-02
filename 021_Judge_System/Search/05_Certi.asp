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
<title>5. 시장출하 적부 판정 기록 기간별 조회 [성적서]</title>
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
        <form method=get action="05_Certi_Ok.asp?<%=Var5%>"  name="form" onSubmit="Check();return true;">
        <tr height="30">
            <td width="600"  bgcolor="#EBFFEB" align="center">
          <font face="돋움"><b>5. 시장출하 적부 판정 기록 기간별 조회 [성적서]</td>
            <td width="200" bgcolor="#EBFFEB"  align="center">
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
	
	
	<br><br><b>성적서 :  <select name="COA_Obtain"><option></option>
                                     <option value="입수">입수</option>
                                     <option value="미입수">미입수</option>
                                     <option value="기입수">기입수</option></select>&nbsp;&nbsp;&nbsp;※ 미선택시 전체가 출력됩니다.</td>
           <td  height="45"  align=center bgcolor="#FFFFFF">
                <span style="font-size:10pt;" ><input type="submit" value="조 회">
                </span> </td>
        </tr>
      
        
</table>
</body>
</form>
</html>
