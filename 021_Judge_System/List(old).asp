<% @LANGUAGE='VBSCRIPT' CODEPAGE='949' %> 
<% session.codepage = 949 %>
<% Response.ChaRset = "euc-kr" %>

<!-- #include file="inc.asp" -->

<html>
<title>시장출하 적부 판정 기록 자료실</title>
<link rel="stylesheet" href="basic.css">
<body bgcolor="#D7F1FA">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<center>

<%
Set RS = SERVER.CreateObject("ADODB.Recordset")
RS.CursorType = 3

SQL="SELECT AL_021_Judge_System.Sid as HD_sid, AL_021_Judge_System.Grp as HD_Grp,AL_021_Judge_System.Seq as HD_Seq, Product_Code,Product_Name_DZ, Delivery_Amount,Lot_number_01,Lot_number_02,Lot_number_03,Lot_number_04,Lot_number_05,Lot_number_06,Lot_number_07,Lot_number_08,Judge_Result,Supplier,Good_class,COA_Obtain,Registor_name, Warehouse,Visit,Sdate, Manage_No, W_Judge_Result,Warehouse_Date from AL_021_Judge_System left outer join  AL_022_Judge_Warehouse ON AL_021_Judge_System.sid=AL_022_Judge_Warehouse.original_sid "


'검색이냐 아니냐에 따라
if Str<>"" then 
SQL = SQL & " WHERE " & Field
SQL = SQL & " LIKE '%" & Str & "%'"

end if

'순서형/응답형에 따라 정렬순서를 다르게 한다
IF SMode = "qa" Then
	SQL = SQL & " ORDER BY HD_sid DESC"
Else 
	SQL = SQL & " ORDER BY HD_Grp DESC,HD_Seq"
END IF 

RS.Open SQL, ConnString




IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize = Pagesize '한페이지에 보여줄 레코드개수(inc.asp에서 정의)
	TotPage   = RS.PageCount
End if
%>
  <table cellspacing=0 cellpadding=0 border=0 width="1024" align="center" valign=top>
  <tr>
  <td align="center" bgcolor="#D7F1FA">

      <table cellspacing=0 cellpadding=2 border=0 width="1024"  align=center  style='table-layout:fixed;'>
          <tr height="40"> 
            <td width=424 align="left">
                           <b>▶ 
              <%if Str<>"" then%>
              검색결과:</b>&nbsp;
                <% if field="Product_Name_DZ"  then %>완제품명으로
                <% elseif field="Product_Code"  then %>제품 코드로
                <% elseif field="Lot_number_All"  then %>로트번호 전체로
                <% elseif field="Lot_number_01" then %>로트번호 1로
                <% elseif field="Lot_number_02" then %>로트번호 2로
                <% elseif field="Lot_number_03" then %>로트번호 3으로
                <% elseif field="Lot_number_04" then %>로트번호 4로
                <% elseif field="Lot_number_05" then %>로트번호 5로
                <% elseif field="Lot_number_06" then %>로트번호 6으로
                <% elseif field="Lot_number_07" then %>로트번호 7로
                <% elseif field="Lot_number_08" then %>로트번호 8로
                <% elseif field="Good_class" then %>제품 종류로
                <% elseif field="Judge_Result" then %>판정 결과로
                <% elseif field="COA_Obtain" then %>성적서 입수로
                <% elseif field="Supplier" then %>제조사
                <% elseif field="Warehouse" then %>입고처로
                <% elseif field="Manage_No" then %>관리번호로
                <% elseif field="Registor_name" then %>판정자로
                
                <% end if %>
              <b><font color=#000088><%=Str%></font></b>(을)를 검색한 결과
              <% else %>
            시장출하 적부 판정 기록
             
              <%end if %>
            </td>
           
           
              <td  width=200 align="left" width="360"">&nbsp;&nbsp;
             Total : <%=TotRecord%> &nbsp;<%=IPage%>/<%=TotPage%> Pages</td>
             
             <td width="400" align="right"> 
             <a href="Search_Code.asp?<%=Var1%>"><img src="images/Code_Serch.gif"  border="0" valign=bottom align="top"></a>
             &nbsp;&nbsp;&nbsp;&nbsp;
             <a href="Search_Name.asp?<%=Var1%>"><img src="images/Name_Serch.gif"  border="0" valign=bottom align="top"></a>
             &nbsp;&nbsp;&nbsp;&nbsp;
           
            <a href="list.asp?<%=Var1%>"><img src="images/list.gif"  border="0"></a></td>
            </tr>
            </table>
      
      
        <table cellspacing=0 width="1024" cellpadding="3" align=center bgcolor=#FFFFFF align=center  style='table-layout:fixed;'>
          <!---게시판 각열의 이름 출력--->
          <tr  height=2> 
          <td bgcolor=#336600></td>
          </tr>
           </table>
      
          <table cellspacing=0 width="1024" cellpadding="3" align=center bgcolor=#FFFFFF align=center  style='table-layout:fixed;'>
          <tr height="50" align="center"> 
      <td width="50"  bgcolor="#CCFFCC"><b>번호</td>
      <td width="100" bgcolor="#CCFFCC"><b>완제품코드</td>
      <td width="319" bgcolor="#CCFFCC"><b>제 품 명</td>
      <td width="50"  bgcolor="#CCFFCC"><b>수량</td>
      <td width="90"  bgcolor="#CCFFCC"><b>LOT</td>
      <td width="50"  bgcolor="#CCFFCC"><b>관리품</td>
      <td width="55"  bgcolor="#CCFFCC"><b>물류QC</td>
      
      <td width="150"  bgcolor="#CCFFCC"  align="center"><b>제조사<br><br style="line-height:3pt;"> 입고처</td>
      <td width="50"  bgcolor="#CCFFCC"  align="center"><b>성적서</td>
      <td width="35"  bgcolor="#CCFFCC"><b>조회</td>
      <td width="75"  bgcolor="#CCFFCC"><b>입고일<br><br style="line-height:3pt;"> 등록일</td>
    </tr>
    <tr  height=1> 
          <td bgcolor=#336600 colspan="11"></td>
          </tr>
          <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=11 align=center height=30><font color=red>"
	Response.Write "검색 조건에 해당되는 자료가없습니다.&nbsp;&nbsp;&nbsp;&nbsp;다른 조건으로 검색해보기 바랍니다."
	Response.Write "</td></tr>"


Else

	
	RS.AbsolutePage = IPage '해당 페이지의 첫번째 레코드로 이동한다
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '레코드번호로 사용할 임시번호
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'이 페이지에서 사용할 필드의 값을 전부 한꺼번에 가져와서 변수에 대입해 둔다.
'이렇게 한번에 가져오는 게 좋은 습관이다.

	HD_sid=RS("HD_sid")
	HD_Grp=RS("HD_Grp")
	HD_Seq=RS("HD_Seq")
	
	
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	Delivery_Amount=RS("Delivery_Amount")
	
	Lot_number_01=RS("Lot_number_01")
	Lot_number_02=RS("Lot_number_02")
	Lot_number_03=RS("Lot_number_03")
	Lot_number_04=RS("Lot_number_04")
	Lot_number_05=RS("Lot_number_05")
	Lot_number_06=RS("Lot_number_06")
	Lot_number_07=RS("Lot_number_07")
	Lot_number_08=RS("Lot_number_08")
	
	Judge_Result=RS("Judge_Result")
	COA_Obtain=RS("COA_Obtain")
	Manage_No=RS("Manage_No")
	COA_Obtain=RS("COA_Obtain")
	
	
	
	Supplier=RS("Supplier")
	Good_class=RS("Good_class")
	Warehouse=RS("Warehouse")
	
	Visit=RS("Visit")
	Sdate=RS("Sdate")
  Registor_name=RS("Registor_name")

	W_Judge_Result=RS("W_Judge_Result")
 
  Warehouse_Date=RS("Warehouse_Date")
 
 
    '반제품명의 길이가 너무 길어 두줄로 넘어가는걸 방지하기 위한 생략 방법
   If len(Product_Name_DZ)>50 then
      str1=".."
      else
      str1=""
     end if

        
 
    '작성된지 24시간이내라면 new!라는 메시지를 준비한다
    Ndate=date()        '현재 날자
       IF datediff("h",Stime,Ndate) < 24 Then       
		  Msg2="<font color=red>ⓝ</font>"
       Else
		  Msg2=""
       End if
 
       
  '미생물 판정 여부에 따라서 다른 색상으로 list에 뿌려준다    
     IF Micro_result="적합" Then     '적합이 아니라면
   		  Msg3="<font color=blue>적합</font>"
     elseif Micro_result="보류" Then
       Msg3="<font color=red>보류</font>"
      elseif  Micro_result="부적합" Then                 
	   	  Msg3="<font color=red>부적합</font>"
	   else                  
	   	  Msg3=""
       end If
 
    %>
     <tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffff00'" >
  
         
           <!--번호를 출력한다-->
            <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
            <%=imsiNO%></td>
              <!--완제품 코드를 출력한다-->
            <td  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:5px;">
          <% if Good_class="화장품" then %>
          <font color=blue><%=Product_Code%></font>
          <% else %>
         <font color=green><%=Product_Code%></a>
          <% end if %></td>
            <td  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:5px;">
          <font color="black">
              <%  
	           IF SMode = "qa" Then
	           For I=1 To Lev
			       Response.Write("&nbsp;")
		         next
             End if 
              %>
             <%=Msg1%>
            <a href="view.asp?<%=Var2%>&sid=<%=HD_sid%>">
             <%=Product_Name_DZ%></a></td>
   
    
    <td style="text-align:right; text-indent:0; margin:0; padding-top:8px; padding-right:5px; padding-bottom:3px; padding-left:0px;">
          <%=formatnumber(Delivery_Amount,0)%></td>
     
      <!--로트넘버를 출력한다--> 
     <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
     <%=Lot_number_01%></td>
     
    <!--관리품 판정결과를 출력한다-->
     <!--물류 판정결과를 출력한다-->
     <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
      <% if Judge_Result="적합"  then %>
            <img src="images/OK.gif">
         <% elseif Judge_Result="부적합"   then %>
          <img src="images/No.gif">
          <% elseif Judge_Result="보류"  then %>
          <img src="images/Hold.gif">
      <% elseif Judge_Result="기입고"  then %>
          <img src="images/Aready.gif">
    
    

          <% else %>
          <img src="images/Not_Regist.gif">
          
          <% end if %></td>
          
         <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
     
          
           <% if W_Judge_Result="적합"  then %>
            <img src="images/OK.gif">
         <% elseif W_Judge_Result="부적합"   then %>
          <img src="images/No.gif">
          <% elseif W_Judge_Result="보류"  then %>
          <img src="images/Hold.gif">
    
          <% else %>
          &nbsp;
          
          <% end if %></td>
     <!--제조사 출력한다-->
    <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
       <%=Supplier%><br><br style="line-height:3pt;">
       <%=Warehouse%></td>
       
       <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
        <% if COA_Obtain="입수"  then %>
            <img src="images/icon_accept.gif">
         <% elseif COA_Obtain="미입수"   then %>
          <img src="images/action_stop.gif">
          <% elseif COA_Obtain="기입수"  then %>
          <img src="images/icon_package_get.gif">
    
          <% else %>
          &nbsp;
          
          <% end if %> </td>
       
          <!--입고일 출력한다-->
     <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
     <%=Visit%></td>
   <td  style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:0px;">
   <%=mid(Warehouse_Date,3)%>
   
    <br><br style="line-height:3pt;">
     <%=mid(Sdate,3)%></td>
  </tr>
          <tr>
            <td colspan="11" height="1" bgcolor="#BBBBBB"></td>
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
          <tr> 
            <td colspan="11" height=1></td>
          </tr>
          <tr  height=1> 
          <td bgcolor=#336600 colspan="11"></td>
          </tr>
        </table>
        <table width="1024" cellpadding="2" cellspacing="0" border="0" align=center>
          <tr> 
            <td height="50" align="center"> 
<%
TAr = Var1 & "&Page=1"
%>
              [<a href="list.asp?<%=Tar%>">&lt;</a>] 
<%
	Gsize = Groupsize 
	PreGNum = (IPage - 1) \ GSize
	EndPNum  = PreGNum * GSize
		Tar = Var1 & "&Page=" & EndPNum
	IF ( EndPNum > 0 ) Then
%>
              [<a href="list.asp?<%=Tar%>">이 전</a>] 

<% 
End IF
%>
[<%
	LCount = GSize
	intI = EndPNum + 1
	Do While (LCount > 0) and (intI <= TotPage)

    '현재 페이지의 색상 표현	 
	if intI = Ipage then
	   intI2 = "<font color=tomato>" & intI & "</font>"
    else
	   intI2 = intI
	end if
		Tar = Var1 & "&Page=" & intI
%> &nbsp;<a href="list.asp? <%=TAr%>" ><%=intI2%></a>
<%
	intI = intI + 1
	LCount = LCount - 1
	Loop
%>
&nbsp;]
<%
	intI = EndPNum + (GSize + 1)
		TAr = Var1 & "&Page=" & intI
	If (intI <= TotPage ) Then
%>
[<a href="list.asp?<%=TAr%>">다음</a>] 
<% 
End if 
TAr = Var1 & "&Page=" & Totpage
%>
              [<a href="list.asp?<%=Tar%>">&gt;</a>] </td>
          </tr>
        </table>
              
        <table width="1024" cellpadding="2" cellspacing="0" border="0" align=center>
          <form method="Get" Action="list.asp">
            <tr> 
              <td align="right">
               <input  type="hidden" name="Table" value="<%=Stable%>">
                <select name="Field" style="width:110">
                  <option value="Product_Name_DZ" <% if field="Product_Name_DZ" then %> selected <% end if %>>완제품명
                  <option value="Product_Code"    <% if field="Product_Code" then %> selected <% end if %>>제품 코드
                  
                  
                  <option value="Lot_number_01"  <% if field="Lot_number_01" then %> selected <% end if %>>로트번호1
                  <option value="Lot_number_02"  <% if field="Lot_number_02" then %> selected <% end if %>>로트번호2
                  <option value="Lot_number_03"  <% if field="Lot_number_03" then %> selected <% end if %>>로트번호3
                  <option value="Lot_number_04"  <% if field="Lot_number_04" then %> selected <% end if %>>로트번호4
                  <option value="Lot_number_05"  <% if field="Lot_number_05" then %> selected <% end if %>>로트번호5
                  <option value="Lot_number_06"  <% if field="Lot_number_06" then %> selected <% end if %>>로트번호6
                  <option value="Lot_number_07"  <% if field="Lot_number_07" then %> selected <% end if %>>로트번호7
                  <option value="Lot_number_08"  <% if field="Lot_number_08" then %> selected <% end if %>>로트번호8
                                    
                  <option value="Good_class"     <% if field="Good_class" then %> selected <% end if %>>제품 종류
                  <option value="Judge_Result"   <% if field="Judge_Result" then %> selected <% end if %>>판정 결과
                  <option value="COA_Obtain"     <% if field="COA_Obtain" then %> selected <% end if %>>성적서 입수
                  <option value="Supplier"       <% if field="Supplier" then %> selected <% end if %>>제조사
                  <option value="Warehouse"      <% if field="Warehouse" then %> selected <% end if %>>입고처
                  <option value="Manage_No"      <% if field="Manage_No" then %> selected <% end if %>>관리번호
                  <option value="Registor_name"  <% if field="Registor_name" then %> selected <% end if %>>등록자</select> 
                  
                  <input type="text" name="Str" value="<%=Str%>" style="width:400">
                  <input  type="submit" value="검색"> 
                  <% if str<>"" then %>
                 <input  type="button" onClick="document.location.href='list.asp?Table=<%=Stable%>'" value="취소">
                <%end if%></td>
              <td align="right" width="170"><a href="Search/Search_condition.asp"><img src="images/inquiry.gif"  border="0"></a></td>
              <td align="center" width="80"><a href="Print/Print_condition.asp"><img src="images/icon_calendar.png"  border="0"></a></td>
            </tr>
          </form>
        </table>
        
        




<script language="javascript">
function Send() {
	var vA = document.form.LOT_NO_All.value;

	if (vA == "") {
		alert("로트번호를 입력하세요.\n");
		document.form.LOT_NO_All.focus();
		return false;
		}



return true;
} // end function
//  -->
</script>


        <form method=get action="List_LOT_Search_Ok.asp?<%=Var5%>"  name="form" onSubmit="return Send()" target="_blank">
        <br style="line-height:8pt;"> 
        <table width="1024" cellpadding="2" cellspacing="0" border="0" align=center>
          <tr> 
              <td align="right">
              <select name="" style="width:150"> 
              <option value="">로트번호 전체 검색</option></td>
              <td align="left" width="190"><input type="text" name="LOT_NO_All" size="25" maxlength="20"></td>
             <td align="left" width="50"><input type="image" img src="images/btn_search.gif" border="0"></td>
              <td width="450" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:3px; padding-left:5px;">
              ← 로트번호 전체 검색은 정확히 일치하는 번호가 검색됩니다!(띠어쓰기 포함)</td>
            </tr>
             </form>
         </table>
        
	  </td>
    </tr>
   
  </table>
<br><br style="line-height:50pt;"> 

</center>
</body>
</html>
