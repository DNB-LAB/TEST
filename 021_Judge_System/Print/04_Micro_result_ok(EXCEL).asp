<%
Response.ContentType = "application/vnd.ms-excel"
%>
<%
if session("id")="" or session("Aceess_RR")<>"P" then  
'로그인하여 얻은 세션(id)가 없으면 로그인으로 돌려 보내고 있으면 리스트를 보여준다.
%>
<html>
<head>
<body  bgcolor="#D7F1FA">
<script language="javascript">
		alert("조회 권한이 있는 사번으로 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../../Log_in_B.asp','end','width=310,height=190,top=270, left=350');
</script>
<% else %>

<!-- #include file="inc.asp" -->
<%

Set RS = SERVER.CreateObject("ADODB.Recordset")
RS.CursorType = 3

Syear = request("Syear")
Smonth = request("Smonth")
Sday = request("Sday")
Microbial_judge = request("Microbial_judge")

Print_date = syear & "-" & smonth & "-" & sday
Print_date = CDATE(Print_date) 

sql = " select ERP_No,P_division, Mark_amount, Mf_unit, Mf_amount, Mark_unit, Syear,smonth,sday, STime_request,Micro_name, Micro_remarks, AL_046_OEM_Product_Microorganism_2018.Micro_result, AL_046_OEM_Product_2018.Sid,AL_046_OEM_Product_2018.grp, AL_046_OEM_Product_2018.lev,P_code, AL_046_OEM_Product_2018.Seq,AL_046_OEM_Product_2018.Visit,AL_046_OEM_Product_2018.Mf_amount, AL_046_OEM_Product_2018.Mf_unit, AL_046_OEM_Product_2018.STime, AL_046_OEM_Product_2018.uTime, Microbial_judge, Subject, Lot_number, Good_class, Registor_name, Produce_date,  Fungus_result, Remarks from AL_046_OEM_Product_2018 left outer JOIN AL_046_OEM_Product_Fungus_2018 ON AL_046_OEM_Product_2018.sid=AL_046_OEM_Product_Fungus_2018.original_sid   left outer join     AL_046_OEM_Product_Microorganism_2018 ON AL_046_OEM_Product_Microorganism_2018.original_sid=AL_046_OEM_Product_2018.sid      "
SQL = SQL & " WHERE Micro_date  = '" & Print_date & "' "  '앞페이지에서 날라온 조회일자와 등록일자가 동일한 자료만 선택



'응답형에 맞게 정렬한다
SQL = SQL & " ORDER BY Sid ASC"

RS.Open SQL, ConnString

IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize=1000 '한페이지에 100개씩 보여준다
	TotPage   = RS.PageCount
End if



%>
<html>
<head>
<title>위,수탁 ODM 및 외주가공 제품/기타 미생물 시험결과 인쇄하기</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">


<center><span style="font-size:14pt;">
<font face="돋움">
<b><%=Print_date%>

<%if weekday(Print_date)=1 then %>(일)
          <%elseif weekday(Print_date)=2 then %>(월)
          <%elseif weekday(Print_date)=3 then %>(화)
          <%elseif weekday(Print_date)=4 then %>(수)
          <%elseif weekday(Print_date)=5 then %>(목)
          <%elseif weekday(Print_date)=6 then %>(금)
          <%elseif weekday(Print_date)=7 then %>(토)<% end if %> 
&nbsp;&nbsp;<b>위탁 ODM 및 외주가공 제품/기타 미생물 시험결과<br></b></font></span>
<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720">
   <tr>
       <td colspan=5   align=left>
       </td>
        <td colspan=4  align=right>조회시간: <%=now%>&nbsp;</td>
    </tr>
</table>


   
   <table border=1 cellspacing="0" cellpadding="0" width="720" align="center" bgcolor="#FAFDF6"  style='table-layout:fixed;'>
      <tr height="40">
        <td align="center" width="30" bgcolor="#CCCCCC"><b>번호</b></td>
        <td align="center" width="50" bgcolor="#CCCCCC"><b><font color="black">의뢰일</font></b></td>
        <td align="center" width="180" bgcolor="#CCCCCC"><b><font color="black">제품코드 / ERP No.<br>제 품 명</font></b></td>
        <td align="center" width="90" bgcolor="#CCCCCC"><font color="black"><b>제조일<br>Lot No</b></font></td>
        <td align="center" width="90" bgcolor="#CCCCCC"><b><font color="black">사업부</font></b></td>
        <td align="center" width="60" bgcolor="#CCCCCC"><font color="black"><b>생산량<br>종류</b></font></td>
        <td align="center" width="60" bgcolor="#CCCCCC"><font color="black"><b>판정</td>
        <td align="center" width="60" bgcolor="#CCCCCC"><font color="black"><b>시험자</b></font></td>
        <td align="center" width="100" bgcolor="#CCCCCC"><b><font color="black">비 &nbsp;고</font></b></td>
      </tr> 
     
<center>


    <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=40> <td colspan=9 align=center>"
	Response.Write "조회 날짜로 등록된 수탁, 외주가공 제품 미생물 판정결과가 없습니다."
	Response.Write "</td></tr>"

  S_number=0 '순번 초기 값
  
 Else
	RS.AbsolutePage = IPage '해당 페이지의 첫번째 레코드로 이동한다
	RCount = RS.pageSize
	Do while (NOT RS.EOF) and (RCount > 0 )


   Sid=RS("sid")
	Grp=RS("grp")
	Seq=RS("seq")
	Lev=RS("lev")
	Subject=RS("Subject")
	P_code=RS("P_code")
	ERP_No=RS("ERP_No")
	P_division=RS("P_division")
	
	Lot_number=RS("Lot_number")
	Mf_amount=RS("Mf_amount")
	Mf_unit=RS("Mf_unit")
	
	Mark_amount=RS("Mark_amount")
	Mark_unit=RS("Mark_unit")
	
	
	good_class=RS("good_class")
	
	Microbial_judge=RS("Microbial_judge")
	Registor_name=RS("Registor_name")
	Syear=RS("Syear")
	smonth=RS("smonth")
	sday=RS("sday")
	Remarks=RS("Remarks")
	  
	STime=RS("stime")
	STime_request=RS("STime_request")
	UTime=RS("utime")
	Visit=RS("visit")
	
	Micro_result=RS("Micro_result")
	Micro_name=RS("Micro_name")
	S_number=S_number+1
	

	
	
%> 
    <tr>
        <td align="center" width="" bgcolor="white"  align="center"><%=S_number%></td>
        <td align="center" width="" bgcolor="white" align="center"><%=mid(left(Stime_request,10),6)%></td>
        <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:5px; padding-bottom:3px; padding-left:5px;">
            <%=P_code%> / <%=ERP_No%> <br>
         <%=Subject%>&nbsp;[ <%=Mark_amount%>  <%=Mark_unit%> ]</td>
        <td align="center" width="" bgcolor="white"><%=Syear%>/<%=Smonth%>/<%=Sday%> <br><%=Lot_number%></td>
        <td align="center" width="" bgcolor="white"><%=P_division%></td>
        <td align="center" width="" bgcolor="white"><%=formatnumber(Mf_amount,0)%> <%=Mf_unit%> <br><% if good_class="위탁제품" then %>
        ODM
        <% else %>
        <%=good_class%><% end if %></td>
       
        <td align="center" width="" bgcolor="white"><%=Micro_result%></td>
        <td align="center" width="" bgcolor="white"><%=Micro_name%></td>
        <td align="left" width="" bgcolor="white"><%=Micro_remarks%>&nbsp;</td>
      
      </tr>  
      
      

  
    <%
	RS.MoveNext
	RCount = RCount -1
	Loop
End if


RS.Close
Set RS=nothing

%> 

  </table>
       
 <% end if %>
</body>
</html>
