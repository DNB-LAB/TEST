<!-- #include file="inc.asp" -->
<%
if session("id")="" or session("Aceess_RR")<>"P" then  
'로그인하여 얻은 세션(id)가 없으면 로그인으로 돌려 보내고 있으면 리스트를 보여준다.
%>

<html>
<head>
<body bgcolor="#D7F1FA">
<script language="javascript">
		alert("조회 권한이 있는 사번으로 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../Log_in_B.asp','end','width=310,height=190,top=270, left=350');
</script>

<% else %>

<!--#include file="inc_Outsourcing_Good.asp"-->

<html>
<head>
<title>제품 코드 검색 결과</title>


<%
P_Code=request("P_Code")

set DB=server.createobject("adodb.connection")
DB.open connstring

sql="select * from AL_038_Outsourcing_Good where P_Code LIKE '%" & P_Code & "%'"
set RS=DB.execute(sql)

'응답형에 맞게 정렬한다

SQL = SQL & " ORDER BY P_Code DESC,sid DESC"

%>

<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="basic.css" type="text/css">
</head>
<br>
<br style="line-height:10px;">
<body topmargin=0 marginheight=0 leftmargin="0"  bgcolor="#D7F1FA">

<table width=720 border=0 cellspacing=0 cellpadding=0  align="center">
<tr height="35"> 
<td valign=middle> <b><font color="#0066FF">☞ <u>외주 제품 Db 검색 결과</u>  </font></b></td>
 <td height=25 valign="bottom" align=right >
           <a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0" valign=bottom align="top"></a>&nbsp;
           <a href="list.asp"><img src=images/list.gif border=0></a></td>
      </tr>
</table>
<table width="720"  border="0" cellpadding="0" cellspacing="0" align=center bgcolor=#FFFFFF> 
       <tr> <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
              <tr height="24"> 
                <td width="40" align="center" valign="middle" bgcolor="#E5B2B2"><b>번호</td>
                <td width="120" align="center" bgcolor="#E5B2B2"><b>제품코드</td>
                <td width="360" align="center" bgcolor="#E5B2B2"><b>제품명</td>
                <td width="70" align="center" bgcolor="#E5B2B2"><b>ERP No</td>
                <td width="60" align="center" bgcolor="#E5B2B2"><b>등록자</td>
                <td width="70" align="center" bgcolor="#E5B2B2"><b>등록일</td>
              </tr>
              <tr> 
                <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
<%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=6 align=center>"
	Response.Write "외주 제품 Db 검색 결과가 없습니다."
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
	P_code=RS("P_code")
	Registor =RS("Registor")
	P_name=RS("P_name")
	ERP_No=RS("ERP_No")
	Visit=RS("visit")
  STime=RS("Stime")
	
	'일련번호 매기기
  S_number = S_number+1
	%> 
	
	<tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffdee9'" >
                <!--등록 번호를 출력한다-->
                <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                  <%=S_number%></td>
                <!--성명을 출력한다-->
                <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                  <a href="product_detail.asp?<%=Var2%>&sid=<%=Sid%>"><strong><%=P_code%></strong></a></td>
                <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                   <a href="Write_Outsourcing_Good.asp?<%=Var2%>&sid=<%=Sid%>"><%=P_name%></a></td>
                   <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=ERP_No%></td>
                  <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=Registor%></td>
                 <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=left(STime,10)%></td>
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
                <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
</table>





<br>
<br style="line-height:30px;">


<!--#include file="inc_Packing_Info.asp"-->

<%
P_Code=request("P_Code")

set DB=server.createobject("adodb.connection")
DB.open connstring

sql="select * from AL_034_Packing_Info where P_Code LIKE '%" & P_Code & "%'"
set RS=DB.execute(sql)

'응답형에 맞게 정렬한다

SQL = SQL & " ORDER BY P_Code DESC,sid DESC"

%>



<table width=720 border=0 cellspacing=0 cellpadding=0  align="center">
<tr height="35"> 
<td valign=middle> <b><font color="#0066FF">☞ <u>충전 포장 제품 Db 검색 결과</u>  </font></b></td>
 <td height=25 valign="bottom" align=right >
           <a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0" valign=bottom align="top"></a>&nbsp;
           <a href="list.asp"><img src=images/list.gif border=0></a></td>
      </tr>
</table>
<table width="720"  border="0" cellpadding="0" cellspacing="0" align=center bgcolor=#FFFFFF> 
       <tr> <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
              <tr height="28"> 
                <td width="40" align="center" valign="middle" bgcolor="#008080"><b>번호</td>
                <td width="120" align="center" bgcolor="#008080"><b>제품코드</td>
                <td width="360" align="center" bgcolor="#008080"><b>제품명</td>
                <td width="70" align="center" bgcolor="#008080"><b>ERP No</td>
                <td width="60" align="center" bgcolor="#008080"><b>등록자</td>
                <td width="70" align="center" bgcolor="#008080"><b>등록일</td>
              </tr>
              <tr> 
                <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
<%
IF (RS.BOF and RS.EOF) Then
	Response.Write "<tr height=50> <td colspan=6 align=center>"
	Response.Write "충전 포장 제품 Db 검색 결과가 없습니다."
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
	P_code=RS("P_code")
	Registor =RS("Registor")
	P_name=RS("P_name")
	ERP_No=RS("ERP_No")
	Visit=RS("visit")
  STime=RS("Stime")
	
	'일련번호 매기기
  S_number = S_number+1
	%> 
	
	<tr onMouseOut="this.style.background='#FFFFFF'" onMouseOver="this.style.background='#ffdee9'" height="28" >
                <!--등록 번호를 출력한다-->
                  <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=S_number%></td>
                <!--성명을 출력한다-->
               <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                  <a href="product_detail.asp?<%=Var2%>&sid=<%=Sid%>"><strong><%=P_code%></strong></a></td>
               <td  style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
                   <a href="Write_Packing_Good.asp?<%=Var2%>&sid=<%=Sid%>"><%=P_name%></a></td>
                   <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=ERP_No%></td>
                  <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=Registor%></td>
                 <td style="text-align:center; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:3px; padding-left:5px;" valign="middle">
              <%=left(STime,10)%></td>
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
                <td height="1" colspan="6" align="center" bgcolor=#BBBBBB></td>
              </tr>
</table>

<br>
<br style="line-height:50px;">
</body>
</html>
       <% end if %>   
