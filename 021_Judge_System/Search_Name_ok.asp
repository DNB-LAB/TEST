<%@Language="VBScript" CODEPAGE="65001" %>
<% 
Response.CharSet="utf-8" 
Session.codepage="65001" 
Response.codepage="65001" 
Response.ContentType="text/html;charset=utf-8" 
%>

<!DOCTYPE HTML>

<!--#include file="inc_Finsh_Goods.asp"-->

<html>
<head>
<title>제품 코드 검색 결과</title>


<%
Product_Name_DZ=request("Product_Name_DZ")

set DB=server.createobject("adodb.connection")
DB.open connstring

sql="select * from AL_032_Finsh_Goods where Product_Name_DZ LIKE '%" & Product_Name_DZ & "%'"
set RS=DB.execute(sql)

'응답형에 맞게 정렬한다

SQL = SQL & " ORDER BY Product_Name_DZ DESC,sid DESC"

%>

<meta http-equiv="content-type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="basic-1.css" type="text/css">
</head>
<br>
<br style="line-height:10px;">
<body topmargin=0 marginheight=0 leftmargin="0"  bgcolor="#D7F1FA">
  <table cellspacing=0 cellpadding=0 border=0 width="1024" align="center" valign=top>
  <tr>
  <td align="center" bgcolor="#D7F1FA">
  
<table width=1024 border=0 cellspacing=0 cellpadding=0  align="center">
<tr height="35"> 
<td bgcolor="#D7F1FA">  <b>☞ 완제품 Db(화장품)에서 완제품 명(더존)으로 [ <font color="red">  <u><%=Product_Name_DZ%></u> </font> ] (을)를  검색한 결과</b></td>
 <td align=right bgcolor="#D7F1FA"  width=100>
           <a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0"></a></td>
      </tr>
</table>

  <table cellspacing=1 width="1024" cellpadding="0" align=center style="table-layout:fixed;">
    <tr height=33> 
      <th width="50">번호</th>
      <th width="110">완제품 코드</th>
      <th width="644">완제품 명(더존)</th>
      <th width="50">구성</th>
      <th width="40">조회</td>
      <th width="60">등록자</th>
      <th width="70">등록일</th>
    </tr>
<%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=7 align=center>"
	Response.Write "완제품 Db(화장품) 검색 결과가 없습니다."
	Response.Write "</td></tr>"
	
Else
 TotRecord = RS.RecordCount
	Rs.pagesize = 200 '한페이지에 보여줄 레코드개수(inc.asp에서 정의)
	TotPage   = RS.PageCount
	
  S_number=0 '순번 초기 값
	
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '레코드번호로 사용할 임시번호
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'이 페이지에서 사용할 필드의 값을 전부 한꺼번에 가져와서 변수에 대입해 둔다.
'이렇게 한번에 가져오는 게 좋은 습관이다.


	Sid=RS("sid")

	Registor=RS("Registor")
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	
	Product_Name_KFDA_01=RS("Product_Name_KFDA_01")
	Product_Name_KFDA_02=RS("Product_Name_KFDA_02")
	Product_Name_KFDA_03=RS("Product_Name_KFDA_03")
	Product_Name_KFDA_04=RS("Product_Name_KFDA_04")
	Product_Name_KFDA_05=RS("Product_Name_KFDA_05")
	Product_Name_KFDA_06=RS("Product_Name_KFDA_06")
	Product_Name_KFDA_07=RS("Product_Name_KFDA_07")
	Product_Name_KFDA_08=RS("Product_Name_KFDA_08")
	
	STime=RS("stime")
	Visit=RS("visit")


'작성된지 12시간이내라면 new!라는 메시지를 준비한다
       Sdate=date()        '현재 날자
       IF datediff("h",Stime,Sdate) < 4 Then       
		  Msg1="<font color=red>ⓝ</font>"
       Else
		  Msg1=""
       End if

'상품명의 길이가 너무 길어 두줄로 넘어가는걸 방지하기 위한 생략 방법
   If len(Product_Name_DZ)>60 then
      str1="..."
      else
      str1=""
     end if
	
	'일련번호 매기기
  S_number = S_number+1
	%> 
	
	<tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffdee9'" >
                <!--등록 번호를 출력한다-->
                <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;" valign="middle">
                  <%=S_number%></td>
               <!--Code를 출력한다-->
                <td  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;" valign="middle">
                  <a href="product_detail.asp?<%=Var2%>&sid=<%=Sid%>"><strong><%=Product_Code%></strong></a></td>
                <!--Product_Name_DZ을 출력한다-->
                <td  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;" valign="middle">
                   <a href="Write_Finsh_Good.asp?<%=Var2%>&sid=<%=Sid%>"><%=left(Product_Name_DZ,60)%><%=str1%><%=Msg1%></a></td>
               <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">            
     <% if Product_Name_KFDA_02 <> ""  then %>
     <font color=red>복합</font>
     
     <% else%>
     <font color=blue>단일</font>
     <% end if %>



    </td>
     
     <!--조회수를 출력한다-->
      <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">            
      <%=Visit%></td>
       <!--등록자를 출력한다-->
      <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">            
      <%=Registor%></td>
      <!--날짜를 출력한다-->
      <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">            
      <%=MID(left(STime,10), 3)%></td>
    </tr>
       
<%
	RS.MoveNext
	RCount = RCount -1
	Loop

End if

RS.Close
Set RS=nothing

%>
             <tr> 
                <td height="1" colspan="7" align="center" bgcolor=#BBBBBB></td>
              </tr>
</table>



<br>
<br style="line-height:30px;">


<!--#include file="inc_Finsh_Goods_Other.asp"-->

<%
Product_Name_DZ=request("Product_Name_DZ")

set DB=server.createobject("adodb.connection")
DB.open connstring

sql="select * from AL_034_Finsh_Goods_Other where Product_Name_DZ LIKE '%" & Product_Name_DZ & "%'"
set RS=DB.execute(sql)

'응답형에 맞게 정렬한다

SQL = SQL & " ORDER BY Product_Name_DZ DESC,sid DESC"

%>




  <table cellspacing=1 width="1024" cellpadding="0" align=center style="table-layout:fixed;">
    <tr height=33> 
     <td align=left bgcolor="#D7F1FA"><b><font color="#0066FF">☞ 완제품 Db(기타) 에서
     
     완제품 명(더존)으로 [ <font color="red"><u><%=Product_Name_DZ%></u>  ] </font> (을)를 검색한 결과</b></td>
     
     
     </td>
     <td align=right bgcolor="#D7F1FA" width=100><a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0"></a></td>
     </tr>
  </table>
  
  <table cellspacing=1 width="1024" cellpadding="0" align=center style="table-layout:fixed;">
    <tr height=33> 
                <td width="50" bgcolor="0A82FF" align=center><b>번호</td>
                <td width="120" bgcolor="0A82FF" align=center><b>완제품(기타) 코드</td>
                <td width="634" bgcolor="0A82FF" align=center><b>완제품(기타) 명(더존)</td>
                <td width="50" bgcolor="0A82FF" align=center><b>구성</td>
                <td width="40" bgcolor="0A82FF" align=center><b>조회</td>
                <td width="60" bgcolor="0A82FF" align=center><b>등록자</td>
                <td width="70" bgcolor="0A82FF" align=center><b><b>등록일</td>
              </tr>
              
<%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=7 align=center>"
	Response.Write "완제품 Db(기타) 검색 결과가 없습니다."
	Response.Write "</td></tr>"
	
Else
 TotRecord = RS.RecordCount
	Rs.pagesize = 200 '한페이지에 보여줄 레코드개수(inc.asp에서 정의)
	TotPage   = RS.PageCount
	
  S_number=0 '순번 초기 값
	
	RCount = RS.pageSize
	
	imsiNO=totrecord-(ipage-1)*(Rcount) '레코드번호로 사용할 임시번호
	
	Do while (NOT RS.EOF) and (RCount > 0 )

'이 페이지에서 사용할 필드의 값을 전부 한꺼번에 가져와서 변수에 대입해 둔다.
'이렇게 한번에 가져오는 게 좋은 습관이다.


	Sid=RS("sid")

	Registor=RS("Registor")
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	
	Product_Name_KFDA_01=RS("Product_Name_KFDA_01")
	Product_Name_KFDA_02=RS("Product_Name_KFDA_02")
	Product_Name_KFDA_03=RS("Product_Name_KFDA_03")
	Product_Name_KFDA_04=RS("Product_Name_KFDA_04")
	Product_Name_KFDA_05=RS("Product_Name_KFDA_05")
	Product_Name_KFDA_06=RS("Product_Name_KFDA_06")
	Product_Name_KFDA_07=RS("Product_Name_KFDA_07")
	Product_Name_KFDA_08=RS("Product_Name_KFDA_08")
	
	STime=RS("stime")
	Visit=RS("visit")
	
	'작성된지 12시간이내라면 new!라는 메시지를 준비한다
       Sdate=date()        '현재 날자
       IF datediff("h",Stime,Sdate) < 4 Then       
		  Msg1="<font color=red>ⓝ</font>"
       Else
		  Msg1=""
       End if

'상품명의 길이가 너무 길어 두줄로 넘어가는걸 방지하기 위한 생략 방법
   If len(Product_Name_DZ)>60 then
      str1="..."
      else
      str1=""
     end if
     
	'일련번호 매기기
  S_number = S_number+1
	%> 
	
	<tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffdee9'" height="28" >
                <!--등록 번호를 출력한다-->
                  <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;" valign="middle">
              <%=S_number%></td>
               <!--Code를 출력한다-->
               <td  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;" valign="middle">
                   <a href="product_detail.asp?<%=Var2%>&sid=<%=Sid%>"><%=Product_Code%></a></td>
                   
               <td  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;" valign="middle">
                  <a href="Write_Finsh_Goods_Other.asp?<%=Var2%>&sid=<%=Sid%>"><%=left(Product_Name_DZ,60)%><%=str1%><%=Msg1%></a></td>
                 <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">            
     <% if Product_Name_KFDA_02 <> ""  then %>
     <font color=red>복합</font>
     
     <% else%>
     <font color=blue>단일</font>
     <% end if %>



    </td>
     
     <!--조회수를 출력한다-->
      <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">            
      <%=Visit%></td>
       <!--등록자를 출력한다-->
      <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">            
      <%=Registor%></td>
      <!--날짜를 출력한다-->
      <td style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">            
      <%=MID(left(STime,10), 3)%></td>
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
<tr><td colspan="7" height=1  bgcolor=6699FF></td></tr>
  </table>

<br>
<br style="line-height:50px;">
</body>
</html>
