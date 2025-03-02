<%@Language="VBScript" CODEPAGE="65001" %>
<% 
Response.CharSet="utf-8" 
Session.codepage="65001" 
Response.codepage="65001" 
Response.ContentType="text/html;charset=utf-8" 
%>

<!DOCTYPE HTML>

<!-- #include file="Inc_Lot_Search.asp" -->
<%

Set RS = SERVER.CreateObject("ADODB.Recordset")
RS.CursorType = 3
SQL = "SELECT * FROM " & STable

'검색이냐 아니냐에 따라
if Str<>"" then 
SQL = SQL & " WHERE " & Field
SQL = SQL & " LIKE '%" & Str & "%'"
end if


'순서형/응답형에 따라 정렬순서를 다르게 한다
IF SMode = "qa" Then
	SQL = SQL & " ORDER BY grp DESC,seq"
Else 
	SQL = SQL & " ORDER BY sid DESC"
END IF 

RS.Open SQL, ConnString


IF (RS.BOF and RS.EOF) Then
	TotRecord = 0 
	TotPage   = 0
Else
	TotRecord = RS.RecordCount
	Rs.pagesize = 15  '한페이지에 보여줄 레코드개수(inc.asp에서 정의)
	TotPage   = RS.PageCount
End if
%>
<html>
<head>
<title> 적합 로트 (20/01~22/03) 자료실</title>
<link rel="stylesheet" href="basic.css">
<meta http-equiv="content-type" content="text/html; charset=utf-8">

</head>

<body bgcolor="#E4F7BA">
  <table cellspacing=0 cellpadding=0 border=0 width="1024" align="center" valign=top>
  <tr>
  <td align="center" bgcolor="#E4F7BA">
  
  
<table cellspacing=0 cellpadding=0 border=0 width="1024" align=center   style="table-layout:fixed;">
    <tr height=50> 
      <td align="left" width=""  bgcolor="#E4F7BA"><b>▶
      <%if Str<>"" then%>
               </b>
                <% if field="Product_Name_DZ"  then %>완제품명으로
                <% elseif field="Product_Code"  then %>제품 코드로
                <% elseif field="Lot_number"  then %>로트번호로
                 <% elseif field="Expiration_Date" then %>사용기한으로
                <% end if %>
              <b><font color=#000088 size=3><%=Str%></font></b>(을)를 검색한 결과
              <% else %> 적합 로트 자료실(20/01~22/03) <% end if %>
     </b></td>
     <td width="200" align="left" class="f8" bgcolor="#E4F7BA" align="center">
     Total : <%=TotRecord%> &nbsp;<%=IPage%>/<%=TotPage%> Pages</td>
      <td width="150"  align="right" bgcolor="#E4F7BA">
      <a href="Lot_Search.asp?<%=Var1%>"><img src="IMAGES/list.gif"  border="0"></a></td>
    </tr>
</table>

  <!--목록을 보여준다.. -->

     <table cellspacing=0 width="1024" cellpadding="0" align=center bgcolor=#FFFFFF align=center  style='table-layout:fixed;'>
          <!---게시판 각열의 이름 출력--->
          <tr  height=2> 
          <td bgcolor=FF7A85></td>
          </tr>
           </table>
      
          <table cellspacing=0 width="1024" cellpadding="0" align=center bgcolor=#FFFFFF align=center  style='table-layout:fixed;'>
          <tr height="30" align="center"> 
      <td width="60"  bgcolor="#FFF0F5"><b>번호</td>
      <td width="100" bgcolor="#FFF0F5"><b>완제품코드</td>
      <td width="444" bgcolor="#FFF0F5"><b>제 품 명</td>
      <td width="190"  bgcolor="#FFF0F5"><b>LOT</td>
      <td width="230"  bgcolor="#FFF0F5"><b>사용기한</td>
    </tr>
    <tr  height=1> 
          <td bgcolor=FF7A85 colspan="5"></td>
          </tr>
          <%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=5 align=center height=30><font color=red>"
	Response.Write "검색 조건에 해당되는 자료가없습니다.&nbsp;&nbsp;&nbsp;&nbsp;다른 조건으로 검색해보기 바랍니다."
	Response.Write "</td></tr>"


Else

	
	RS.AbsolutePage = IPage '해당 페이지의 첫번째 레코드로 이동한다
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '레코드번호로 사용할 임시번호
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'이 페이지에서 사용할 필드의 값을 전부 한꺼번에 가져와서 변수에 대입해 둔다.
'이렇게 한번에 가져오는 게 좋은 습관이다.

	Sid=RS("Sid")
	Grp=RS("Grp")
	Seq=RS("Seq")
	
	
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	Lot_number=RS("Lot_number")
	
	Expiration_Date=RS("Expiration_Date")
	
 
    '로트의 길이가 너무 길어 두줄로 넘어가는걸 방지하기 위한 생략 방법
   If len(Lot_number)>25 then
      str1="..."
      else
      str1=""
     end if


   '사용기한의 길이가 너무 길어 두줄로 넘어가는걸 방지하기 위한 생략 방법
   If len(Expiration_Date)>35 then
      str2="..."
      else
      str2=""
     end if
        

 
    %>
     <tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffff00'" >
  
         
           <!--번호를 출력한다-->
            <td  style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
            <%=imsiNO%></td>
              <!--완제품 코드를 출력한다-->
            <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=Product_Code%></td>
            <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
          <font color="black">
              <%  
	           IF SMode = "qa" Then
	           For I=1 To Lev
			       Response.Write("&nbsp;")
		         next
             End if 
              %>
             <%=Msg1%>
            <a href="Lot_View.asp?<%=Var2%>&sid=<%=sid%>">
             <%=Product_Name_DZ%></a></td>
   
    
   
     
      <!--로트넘버를 출력한다--> 
     <td  style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
     <%=left(Lot_number, 25)%><%=str1%></td>
      <td  style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
   <%=left(Expiration_Date, 35)%><%=str2%></td>
  </tr>
          <tr>
            <td colspan="5" height="1" bgcolor="#BBBBBB"></td>
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
            <td colspan="5" height=1></td>
          </tr>
          <tr  height=1> 
          <td bgcolor=FF7A85 colspan="11"></td>
          </tr>
        </table>
        <table width="1024" cellpadding="2" cellspacing="0" border="0" align=center>
          <tr> 
            <td height="50" align="center"> 
<%
TAr = Var1 & "&Page=1"
%>
              [<a href="Lot_Search.asp?<%=Tar%>">&lt;</a>] 
<%
	Gsize = Groupsize 
	PreGNum = (IPage - 1) \ GSize
	EndPNum  = PreGNum * GSize
		Tar = Var1 & "&Page=" & EndPNum
	IF ( EndPNum > 0 ) Then
%>
              [<a href="Lot_Search.asp?<%=Tar%>">이 전</a>] 

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
%> &nbsp;<a href="Lot_Search.asp? <%=TAr%>" ><%=intI2%></a>
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
[<a href="Lot_Search.asp?<%=TAr%>">다음</a>] 
<% 
End if 
TAr = Var1 & "&Page=" & Totpage
%>
              [<a href="Lot_Search.asp?<%=Tar%>">&gt;</a>] </td>
          </tr>
        </table>
              
        <table width="1024" cellpadding="2" cellspacing="0" border="0" align=center>
          <form method="Get" Action="Lot_Search.asp">
            <tr> 
              <td align="center">
               <input  type="hidden" name="Table" value="<%=Stable%>">
                <select name="Field" style="width:110">
                   <option value="Lot_number"  <% if field="Lot_number" then %> selected <% end if %>>로트번호
                  <option value="Product_Name_DZ" <% if field="Product_Name_DZ" then %> selected <% end if %>>완제품명
                  <option value="Product_Code"    <% if field="Product_Code" then %> selected <% end if %>>제품 코드
                  
                  
                 
                  <option value="Expiration_Date"  <% if field="Expiration_Date" then %> selected <% end if %>>사용기한
                </select> 
                  
                  <input type="text" name="Str" value="<%=Str%>" style="width:400">
                  <input  type="submit" value="검색"> 
                  <% if str<>"" then %>
                 <input  type="button" onClick="document.location.href='Lot_Search.asp?Table=<%=Stable%>'" value="취소">
                <%end if%></td>
            </tr>
          </form>
        </table>  </td>
    </tr>
   
</table>

   <table align="center" width=1024 align="center"  style='table-layout:fixed;'>
         <tr height=50>
             <td  align=center>
             <a href= "javascript:window.open('about:blank', '_self').close();">
             <img src="images/close.gif" border="0"></a></td>
          </tr>
        </table>
<br><br style="line-height:50pt;"> 

</body>
</html>
