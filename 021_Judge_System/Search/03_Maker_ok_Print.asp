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

Set RS = SERVER.CreateObject("ADODB.Recordset")
RS.CursorType = 3

Styear = request("Styear")
Stmonth = request("Stmonth")
Stday = request("Stday")

Supplier = request("Supplier")

Fyear = request("Fyear")
Fmonth = request("Fmonth")
Fday = request("Fday")

Stdate = Styear & "-" & Stmonth & "-" & Stday
Stdate = CDATE(Stdate) 

Fdate = Fyear & "-" & Fmonth & "-" & Fday
Fdate = CDATE(Fdate) 



if Supplier = ""  then

sql="SELECT AL_021_Judge_System.sid, Product_Code,Product_Name_DZ, Delivery_Amount,Lot_number_01,Lot_number_02, Lot_number_03, Lot_number_04, Lot_number_05, Lot_number_06, Lot_number_07, Lot_number_08, Expiration_Date_01,Expiration_Date_02, Expiration_Date_03, Expiration_Date_04, Expiration_Date_05, Expiration_Date_06, Expiration_Date_07, Expiration_Date_08, Lot_No_Divide, Good_class,  Judge_Result, Supplier, Warehouse, Manage_No, COA_Obtain, Warehouse_Date, Remarks, Registor_name, Visit,Sdate, W_Registor, W_Judge_Method, W_Judge_Result, W_Remarks  from AL_021_Judge_System left outer join  AL_022_Judge_Warehouse ON AL_021_Judge_System.sid=AL_022_Judge_Warehouse.original_sid "
SQL = SQL & " WHERE Warehouse_Date  between '"& Stdate & "' and '" & Fdate & "'"   '지정기간 출력


elseif Supplier <> "" then

sql="SELECT AL_021_Judge_System.sid, Product_Code,Product_Name_DZ, Delivery_Amount,Lot_number_01,Lot_number_02, Lot_number_03, Lot_number_04, Lot_number_05, Lot_number_06, Lot_number_07, Lot_number_08, Expiration_Date_01,Expiration_Date_02, Expiration_Date_03, Expiration_Date_04, Expiration_Date_05, Expiration_Date_06, Expiration_Date_07, Expiration_Date_08, Lot_No_Divide, Good_class,  Judge_Result, Supplier, Warehouse, Manage_No, COA_Obtain, Warehouse_Date, Remarks, Registor_name, Visit,Sdate, W_Registor, W_Judge_Method, W_Judge_Result, W_Remarks  from AL_021_Judge_System left outer join  AL_022_Judge_Warehouse ON AL_021_Judge_System.sid=AL_022_Judge_Warehouse.original_sid "
SQL = SQL & " WHERE Warehouse_Date  between '"& Stdate & "' and '" & Fdate & "'"   '지정기간 출력
SQL = SQL & " and Supplier LIKE  '%" & Supplier & "%' "  

end if 

'응답형에 맞게 정렬한다
SQL = SQL & " ORDER BY Warehouse_Date ASC, sid"

RS.Open SQL, ConnString

IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize=100000 '한페이지에 10000개씩 보여준다
	TotPage   = RS.PageCount
End if
%>



<html>
<head>
<title>3. 시장출하 적부 판정 기록 기간별 조회 [ 제조사] 인쇄</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="content-type" content="text/html; charset=utf-8">

<script language="javascript">
function printWindow() {
factory.printing.header = ""
factory.printing.footer = ""   
factory.printing.portrait = false    
factory.printing.leftMargin = 0.0   
factory.printing.topMargin = 15.0    
factory.printing.rightMargin = 0.0
factory.printing.bottomMargin = 0.0
factory.printing.Print(true, window)
}
</script>


</head>
<body id="b1" onload="javascript:printWindow();">

<object id="factory" style="display:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" 
codebase="http://www.meadroid.com/scriptx/ScriptX.cab#Version=6,1,429,14">
</object>
<table style="border-collapse:collapse;" align="center" cellspacing="0" width="1080">
   <tr>
       <td width="1080"  align=left>
입고일로<b> <%=mid(Stdate,3)%></b>부터 <b><%=mid(Fdate,3)%></b>까지&nbsp;&nbsp;&nbsp;&nbsp;
<% if Supplier="" then %>
          제조사로 <b>[]</font></b>을(를) 검색한 결과
<% elseif Supplier <> "" then %>          
       제조사로  <b>[ <%=Supplier%> ]</font></b>을(를) 검색한 결과</span>
          
  <% end if %>        
</td>
    </tr>
</table>
 
 
<table style="border-collapse:collapse;" align="center" cellspacing="0" width="1080">
   <tr>
       <td width="360"  align=left></td>
        <td width=""  align=right>조회시간: <%=now%>&nbsp;</td>
    </tr>
</table>    

<table style="border-collapse:collapse;" align="center" cellspacing="0" width="1080" height="93%">
    <tr>
        <td width="1080" style="border-width:0%; border-color:white; border-style:none;" valign="top">
<table border=1 cellspacing="0" cellpadding="0" width="1080" align="center"   style='table-layout:fixed;'>
        <!---게시판 각열의 이름 출력--->
         
       <tr height="42">
        <td align="center" width="40"  bgcolor="#CCCCCC"><b>번호</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b><b>입고일</b></td>
        <td align="center" width="90" bgcolor="#CCCCCC"><b>품목코드</b></td>
        <td align="center" width="270" bgcolor="#CCCCCC"><b>제 품 명</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b><b>입고수량</b></td>
        <td align="center" width="70"  bgcolor="#CCCCCC"><b>Lot No</b></td>
        <td align="center" width="80"  bgcolor="#CCCCCC"><b>사용기한</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b>관리품<br>판정</b></td>
        <td align="center" width="60"  bgcolor="#CCCCCC"><b>물류<br>QC</b></td>
        <td align="center" width="150"  bgcolor="#CCCCCC"><b>입고처<br>출고처</b></td>
        <td align="center" width="50"  bgcolor="#CCCCCC"><b>관리<br>번호</b></td>
        <td align="center" width="90" bgcolor="#CCCCCC"><b>비&nbsp;고</b></td>
      </tr> 
          <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr> <td colspan=12  align=center height=50 bgcolor=white><font color=red>"
	Response.Write "검색 조건에 해당되는 자료가없습니다.&nbsp;&nbsp;&nbsp;&nbsp;다른 조건으로 검색해보기 바랍니다."
	Response.Write "</td></tr>"

S_number=0 '순번 초기 값

Else

	
	RS.AbsolutePage = IPage '해당 페이지의 첫번째 레코드로 이동한다
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '레코드번호로 사용할 임시번호
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'이 페이지에서 사용할 필드의 값을 전부 한꺼번에 가져와서 변수에 대입해 둔다.
'이렇게 한번에 가져오는 게 좋은 습관이다.

	Sid=RS("sid")
	
	Warehouse_Date=RS("Warehouse_Date")
	Manage_No=RS("Manage_No")
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	Delivery_Amount=RS("Delivery_Amount")
	Lot_No_Divide=RS("Lot_No_Divide")
	Good_class=RS("Good_class")
	Judge_Result=RS("Judge_Result")
	Supplier=RS("Supplier")
	Warehouse=RS("Warehouse")
	
	
	
	
	Lot_number_01=RS("Lot_number_01")
	Lot_number_02=RS("Lot_number_02")
	Lot_number_03=RS("Lot_number_03")
	Lot_number_04=RS("Lot_number_04")
	Lot_number_05=RS("Lot_number_05")
	Lot_number_06=RS("Lot_number_06")
	Lot_number_07=RS("Lot_number_07")
	Lot_number_08=RS("Lot_number_08")
	
	Expiration_Date_01=RS("Expiration_Date_01")
	Expiration_Date_02=RS("Expiration_Date_02")
	Expiration_Date_03=RS("Expiration_Date_03")
	Expiration_Date_04=RS("Expiration_Date_04")
	Expiration_Date_05=RS("Expiration_Date_05")
	Expiration_Date_06=RS("Expiration_Date_06")
	Expiration_Date_07=RS("Expiration_Date_07")
	Expiration_Date_08=RS("Expiration_Date_08")
	
	Remarks=RS("Remarks")
	Registor_name=RS("Registor_name")
	
	W_Registor=RS("W_Registor")
	W_Judge_Method=RS("W_Judge_Method")
	W_Judge_Result=RS("W_Judge_Result")
	W_Remarks=RS("W_Remarks")
	
 
S_number=S_number+1
 
 
 'Remarks = replace (Remarks,"&","&amp;")
 'Remarks = replace (Remarks,"<","&lt;")
 'Remarks = replace (Remarks,">","&gt;")
 'Remarks = replace (Remarks,Chr(32),"&nbsp;") '공백(스페이스)
 'Remarks = replace (Remarks,Chr(13),"<br>") '줄바꿈
 
 
   %>
 <tr>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
          <%=S_number%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <%=mid(Warehouse_Date,3)%></td>
        <td style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:5px; padding-bottom:1px; padding-left:5px;">
           <%=Product_Code%></td>
        <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:5px; padding-bottom:1px; padding-left:5px;">
           <%=Product_Name_DZ%>
        </td>
       <td style="text-align:right; text-indent:0; margin:0; padding-top:3px; padding-right:5px; padding-bottom:1px; padding-left:0px;">
           <%=formatnumber(Delivery_Amount,0)%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
          <%=Lot_number_01%>
          <% if Lot_number_02 <>"" then%><br><%=Lot_number_02%><% else %><% end if%>
          <% if Lot_number_03 <>"" then%><br><%=Lot_number_03%><% else %><% end if%>
          <% if Lot_number_04 <>"" then%><br><%=Lot_number_04%><% else %><% end if%>
          <% if Lot_number_05 <>"" then%><br><%=Lot_number_05%><% else %><% end if%>
          <% if Lot_number_06 <>"" then%><br><%=Lot_number_06%><% else %><% end if%>
          <% if Lot_number_07 <>"" then%><br><%=Lot_number_07%><% else %><% end if%>
          <% if Lot_number_08 <>"" then%><br><%=Lot_number_08%><% else %><% end if%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <%=Expiration_Date_01%>
          <% if Expiration_Date_02 <>"" then%><br><%=Expiration_Date_02%><% else %><% end if%>
          <% if Expiration_Date_03 <>"" then%><br><%=Expiration_Date_03%><% else %><% end if%>
          <% if Expiration_Date_04 <>"" then%><br><%=Expiration_Date_04%><% else %><% end if%>
          <% if Expiration_Date_05 <>"" then%><br><%=Expiration_Date_05%><% else %><% end if%>
          <% if Expiration_Date_06 <>"" then%><br><%=Expiration_Date_06%><% else %><% end if%>
          <% if Expiration_Date_07 <>"" then%><br><%=Expiration_Date_07%><% else %><% end if%>
          <% if Expiration_Date_08 <>"" then%><br><%=Expiration_Date_08%><% else %><% end if%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
        <%=Judge_Result%></td>
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
        <% if W_Judge_Result<>"" then %><%=W_Judge_Result%><% else %>&nbsp;<% end if %></td>
       
       
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
        <%=Supplier%><br><br style="line-height:3pt;"> <%=Warehouse%></td>
       
       
       
        <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
           <% if Manage_No<> "" then %><%=Manage_No%><% else %>&nbsp;<% end if %></td>

         <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
       <%=Remarks%>&nbsp;</td>
      </tr>  

          <%
	RS.MoveNext

	imsiNO=imsiNO-1
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
        <td width="1080" style="border-width:0%; border-color:white; border-style:none;" height="100">
        </td>
    </tr>
</table>


</body>
</html>
