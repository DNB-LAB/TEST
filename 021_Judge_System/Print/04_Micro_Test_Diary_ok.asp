<!-- #include file="inc.asp" -->
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
<%

Set RS = SERVER.CreateObject("ADODB.Recordset")
RS.CursorType = 3

Syear = request("Syear")
Smonth = request("Smonth")
Sday = request("Sday")

Microbial_judge = request("Microbial_judge")

Print_date = syear & "-" & smonth & "-" & sday
Print_date = CDATE(Print_date) 



sql = " select ERP_No,P_division, Mark_amount, Mf_unit, Mf_amount, Mark_unit, Syear,smonth,sday, STime_request,Micro_name, AL_046_OEM_Product_Microorganism_2018.Micro_result, AL_046_OEM_Product_2018.Sid,AL_046_OEM_Product_2018.grp, AL_046_OEM_Product_2018.lev,P_code, AL_046_OEM_Product_2018.Seq,AL_046_OEM_Product_2018.Visit,AL_046_OEM_Product_2018.Mf_amount, AL_046_OEM_Product_2018.Mf_unit, AL_046_OEM_Product_2018.STime, AL_046_OEM_Product_2018.uTime, Microbial_judge, Subject, Lot_number, Good_class, Registor_name, Produce_date,  Fungus_result, Remarks from AL_046_OEM_Product_2018 left outer JOIN AL_046_OEM_Product_Fungus_2018 ON AL_046_OEM_Product_2018.sid=AL_046_OEM_Product_Fungus_2018.original_sid   left outer join     AL_046_OEM_Product_Microorganism_2018 ON AL_046_OEM_Product_Microorganism_2018.original_sid=AL_046_OEM_Product_2018.sid      "
SQL = SQL & " WHERE  left(AL_046_OEM_Product_2018.Stime, 10)  = '" & Print_date & "' "  '앞페이지에서 날라온 조회일자와 등록일자가 동일한 자료만 선택





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
<title>위탁 ODM 및 외주가공 제품/기타 미생물 시험일지 조회</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">


<center><span style="font-size:14pt;"><b>
<font face="돋움"><%=Print_date%>

<%if weekday(Print_date)=1 then %>(일)
          <%elseif weekday(Print_date)=2 then %>(월)
          <%elseif weekday(Print_date)=3 then %>(화)
          <%elseif weekday(Print_date)=4 then %>(수)
          <%elseif weekday(Print_date)=5 then %>(목)
          <%elseif weekday(Print_date)=6 then %>(금)
          <%elseif weekday(Print_date)=7 then %>(토)<% end if %>

         
          위탁 ODM 및 외주가공 제품/기타 미생물 시험일지<br></b></font></span>
   
          
<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720">
    <tr>
       <td width="60"  align=left>
       <a href="04_Micro_Test_Diary_ok_Print.asp?Syear=<%=Syear%>&Smonth=<%=Smonth%>&Sday=<%=Sday%>&Microbial_judge=<%=Microbial_judge%>">
       <img src="../images/print.gif" border=0></a></td>
        <td width="360"  align=left></td>
        <td width="300"  align=right>조회시간: <%=now%>&nbsp;</td>
    </tr>
     <tr>
      <td  style="text-align:left; text-indent:0; margin:0; padding-top:20px; padding-right:5px; padding-bottom:0px; padding-left:0px;">            
      <form method="post" action="04_Micro_Test_Diary_ok_Print_Control.asp?sid=<%=Sid%>" name="form" OnSubmit="return Send()">
       <input type="image"   img src="../images/print.gif" border=0></td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:15px; padding-right:5px; padding-bottom:0px; padding-left:0px;">            
        (출력시간, 결재란 날자 有, ISO 번호, 회사명 수정 출력)</a>
        <input type=hidden name=Syear value="<%=Syear%>">
        <input type=hidden name=Smonth value="<%=Smonth%>">
        <input type=hidden name=Sday value="<%=Sday%>">
        <input type=hidden name=Microbial_judge value="1"></td>
      
         <td   style="text-align:right; text-indent:0; margin:0; padding-top:5px; padding-right:5px; padding-bottom:0px; padding-left:0px;">            
   <input type="text" name="New_Now" size="40" style="text-align:right;" value="출력시간 : <%=Syear%>-<%=Smonth%>-<%=Sday%>&nbsp;<%=RIGHT(now,11)%>" ></td>
        </tr>
</table>


<table style="border-collapse:collapse;" align="center" cellspacing="0" width="720" height="93%">
    <tr>
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" valign="top">
<table border=1 cellspacing="0" cellpadding="0" width="720" align="center" bgcolor="#FAFDF6" style='table-layout:fixed;'>
      <tr height="22">
        <td align="center" width="30" bgcolor="#CCCCCC"  rowspan=2><b>번호</b></center></td>
        <td align="center" width="320" bgcolor="#CCCCCC" rowspan=2  ><b><font color="black">제품명</font></b></td>
        <td align="center" width="70" bgcolor="#CCCCCC"  rowspan=2><b>생산일</b></td>
        <td align="center" width="70" bgcolor="#CCCCCC"  rowspan=2 ><b><font color="black">Lot No</font></b></td>
        <td align="center" width="120" bgcolor="#CCCCCC"  colspan=3><font color="black"><b>결과(일)</b></font></td>
        <td align="center" width="60" bgcolor="#CCCCCC"  rowspan=2><font color="black"><b>판정</b></font></td>
        <td align="center" width="40" bgcolor="#CCCCCC"  rowspan=2><b><font color="black">비고</font></b></td>
      </tr>  
      <tr height="22"> 
      <td align="center" width="40" bgcolor="#CCCCCC" ><font color="black"><b>1</b></font></td>
      <td align="center" width="40" bgcolor="#CCCCCC" ><font color="black"><b>2</b></font></td>
      <td align="center" width="40" bgcolor="#CCCCCC"><font color="black"><b>3</b></font></td>
      </tr>  


    <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=40> <td colspan=9 align=center>"
	Response.Write "조회 날짜로 등록된 의뢰가 없습니다."
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
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        <%=S_number%></td>
        
        <td style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;">            
       <%=Subject%> &nbsp;[ <%=Mark_amount%>  <%=Mark_unit%> ]
        
        </td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        <%=Syear%>/<%=Smonth%>/<%=Sday%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        <%=Lot_number%></td>
       <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:0px;">            
        &nbsp;</td>
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
        </td>
    </tr>
    <tr>
        <td   align=right style="border-width:0%; border-color:white; border-style:none;" height="100">

       결재란 <select name="Approval" size="1" style="width:150px">
                    <option value="New">날짜 有</option>
                    <option value="Old">날짜 無</option>
                    </select></td>
    </tr>
    <tr>
     
        <td width="720" style="border-width:0%; border-color:white; border-style:none;" height="100">

<table width="720" cellspacing=0 cellpadding=0 border="0"  height="10">
<tr>
  <td align=left><input type="text" name="New_QI_No" size="30"></td>
  <td align=right><select name="New_Company" size="1" style="width:150px">
                    <option value="(주)잇츠한불">(주)잇츠한불</option>
                    <option value="한불화장품(주)">한불화장품(주)</option>
                    </select></td>
</tr>

</table>
<% end if %>
</body>
</html>
