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
if session("id")="" then  
'로그인하여 얻은 세션(id)가 없으면 로그인으로 돌려 보내고 있으면 리스트를 보여준다.
%>

<html>
<head>
<body leftmargin="0" topmargin="0" bgcolor="#D7F1FA">
<script language="javascript">
		alert("로그인이 필요합니다. 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../../Log_in.asp','end','width=310,height=190,top=270, left=350');
</script>

<% else %>


<!--#include file="inc_Finsh_Goods_Other.asp"-->



<%
'내용을 볼 글번호를 전송받는다
Sid= Request.QueryString("sid")
Var3 = Var2 & "&sid=" & Sid


Set DB = Server.Createobject("ADODB.connection")
DB.open Connstring

sql="select * from AL_034_Finsh_Goods_Other"
SQL = SQL & " WHERE sid =" & Sid
Set RS = DB.Execute(SQL)
Sid=RS("sid")

	Registor=RS("Registor")
	Product_Code=RS("Product_Code")
	Product_Name_DZ=RS("Product_Name_DZ")
	Product_Name_Final=RS("Product_Name_Final")
	
	Product_Name_KFDA_01=RS("Product_Name_KFDA_01")
	Product_Name_KFDA_02=RS("Product_Name_KFDA_02")
	Product_Name_KFDA_03=RS("Product_Name_KFDA_03")
	Product_Name_KFDA_04=RS("Product_Name_KFDA_04")
	Product_Name_KFDA_05=RS("Product_Name_KFDA_05")
	Product_Name_KFDA_06=RS("Product_Name_KFDA_06")
	Product_Name_KFDA_07=RS("Product_Name_KFDA_07")
	Product_Name_KFDA_08=RS("Product_Name_KFDA_08")
	
	
	Bulk_Code_01= RS("Bulk_Code_01")
	Product_Name_KFDA_01= RS("Product_Name_KFDA_01")
	P_Class_01= RS("P_Class_01")
	P_Capacity_01= RS("P_Capacity_01")
	P_Capacity_Unit_01= RS("P_Capacity_Unit_01")
	Period_of_Usage_01= RS("Period_of_Usage_01")
	Manufacturer_01= RS("Manufacturer_01")
	
	
	Bulk_Code_02= RS("Bulk_Code_02")
	Product_Name_KFDA_02= RS("Product_Name_KFDA_02")
	P_Class_02= RS("P_Class_02")
	P_Capacity_02= RS("P_Capacity_02")
	P_Capacity_Unit_02= RS("P_Capacity_Unit_02")
	Period_of_Usage_02= RS("Period_of_Usage_02")
	Manufacturer_02= RS("Manufacturer_02")
	
	
	Bulk_Code_03= RS("Bulk_Code_03")
	Product_Name_KFDA_03= RS("Product_Name_KFDA_03")
	P_Class_03= RS("P_Class_03")
	P_Capacity_03= RS("P_Capacity_03")
	P_Capacity_Unit_03= RS("P_Capacity_Unit_03")
	Period_of_Usage_03= RS("Period_of_Usage_03")
	Manufacturer_03= RS("Manufacturer_03")
	
	
  Bulk_Code_04= RS("Bulk_Code_04")
	Product_Name_KFDA_04= RS("Product_Name_KFDA_04")
	P_Class_04= RS("P_Class_04")
	P_Capacity_04= RS("P_Capacity_04")
	P_Capacity_Unit_04= RS("P_Capacity_Unit_04")
	Period_of_Usage_04= RS("Period_of_Usage_04")
	Manufacturer_04= RS("Manufacturer_04")
	

  Bulk_Code_05= RS("Bulk_Code_05")
	Product_Name_KFDA_05= RS("Product_Name_KFDA_05")
	P_Class_05= RS("P_Class_05")
	P_Capacity_05= RS("P_Capacity_05")
	P_Capacity_Unit_05= RS("P_Capacity_Unit_05")
	Period_of_Usage_05= RS("Period_of_Usage_05")
	Manufacturer_05= RS("Manufacturer_05")
	

  Bulk_Code_06= RS("Bulk_Code_06")
	Product_Name_KFDA_06= RS("Product_Name_KFDA_06")
	P_Class_06= RS("P_Class_06")
	P_Capacity_06= RS("P_Capacity_06")
	P_Capacity_Unit_06= RS("P_Capacity_Unit_06")
	Period_of_Usage_06= RS("Period_of_Usage_06")
	Manufacturer_06= RS("Manufacturer_06")
	
	
	Bulk_Code_07= RS("Bulk_Code_07")
	Product_Name_KFDA_07= RS("Product_Name_KFDA_07")
	P_Class_07= RS("P_Class_07")
	P_Capacity_07= RS("P_Capacity_07")
	P_Capacity_Unit_07= RS("P_Capacity_Unit_07")
	Period_of_Usage_07= RS("Period_of_Usage_07")
	Manufacturer_07= RS("Manufacturer_07")
	
	
	Bulk_Code_08= RS("Bulk_Code_08")
	Product_Name_KFDA_08= RS("Product_Name_KFDA_08")
	P_Class_08= RS("P_Class_08")
	P_Capacity_08= RS("P_Capacity_08")
	P_Capacity_Unit_08= RS("P_Capacity_Unit_08")
	Period_of_Usage_08= RS("Period_of_Usage_08")
	Manufacturer_08= RS("Manufacturer_08")
	
	
	
	STime=RS("stime")
	Visit=RS("visit")
  
   Cosmetic_Class_Big_01 = LEFT(P_Class_01,1)
   
   '유형
  IF Cosmetic_Class_Big_01="가" Then            
   		  Msg_01="영ㆍ유아용(만 3세이하의 어린이용을 말한다.) 제품류"
     elseif Cosmetic_Class_Big_01="나" Then     
         Msg_01="목욕용 제품류"
     elseif  Cosmetic_Class_Big_01="다" Then  
	   	   Msg_01="인체 세정용 제품류(화장비누 제외)"
	  elseif  Cosmetic_Class_Big_01="라" Then
	   	   Msg_01="눈 화장용 제품류"
    elseif  Cosmetic_Class_Big_01="마" Then
	   	   Msg_01="방향용 제품류"
	  elseif  Cosmetic_Class_Big_01="바" Then
	   	   Msg_01="두발 염색용 제품류"   
	  elseif  Cosmetic_Class_Big_01="사" Then
	   	   Msg_01="색조 화장용 제품"
	  elseif  Cosmetic_Class_Big_01="아" Then
	   	   Msg_01="두발용 제품류"   
	  elseif  Cosmetic_Class_Big_01="자" Then
	   	   Msg_01="손발톱용 제품류"
	  elseif  Cosmetic_Class_Big_01="차" Then
	   	   Msg_01="면도용 제품류"   
	  elseif  Cosmetic_Class_Big_01="카" Then
	   	   Msg_01="기초화장용 제품류"
	  elseif  Cosmetic_Class_Big_01="타" Then
	   	   Msg_01="체취 방지용 제품류" 
	 elseif  Cosmetic_Class_Big_01="파" Then
	   	   Msg_01="체모 제거용 제품류"
	   	   
	 else
	   	  Msg_01=""
     end If	   	   
	
	
	IF 	P_Class_01="가1" Then 
      MsgP_01="1) 영ㆍ유아용 샴푸, 린스"
	  
	    elseif P_Class_01="가2" Then 
	    MsgP_01="2) 영ㆍ유아용 로션, 크림"
	    
	    elseif P_Class_01="가3" Then 
	    MsgP_01="3) 영ㆍ유아용 오일"
	    
	    elseif P_Class_01="가4" Then 
	    MsgP_01="4) 영ㆍ유아용 인체 세정용 제품"
	    
	    elseif P_Class_01="가5" Then 
	    MsgP_01="5) 영ㆍ유아용 목욕용 제품"
	    
	    elseif P_Class_01="나1" Then 
	    MsgP_01="1) 목욕용 오일ㆍ정제ㆍ캡슐"
	    
	    elseif P_Class_01="나2" Then 
	    MsgP_01="2) 목욕용 소금류"
	    
	    elseif P_Class_01="나3" Then 
	    MsgP_01="3) 버블 배스"
	    
	    elseif P_Class_01="나4" Then 
	    MsgP_01="4) 그 밖의 목욕용 제품류"
	    
	    elseif P_Class_01="다1" Then 
	    MsgP_01="1) 폼 클렌저"
	    
	    elseif P_Class_01="다2" Then 
	    MsgP_01="2) 바디 클렌저"
	    
	    elseif P_Class_01="다3" Then 
	    MsgP_01="3) 액상비누"
	    
	    elseif P_Class_01="다4" Then 
	    MsgP_01="4) 외음부 세정제"
	    
	    elseif P_Class_01="다5" Then 
	    MsgP_01="5) 그 밖의 인체 세정용 제품류"
	    
	    elseif P_Class_01="라1" Then 
	    MsgP_01="1) 아이브로 펜슬"
	    
	    elseif P_Class_01="라2" Then 
	    MsgP_01="2) 아이 라이너"
	    
	    elseif P_Class_01="라3" Then 
	    MsgP_01="3) 아이 섀도"
	    
	    elseif P_Class_01="라4" Then 
	    MsgP_01="4) 마스카라"
	    
	    elseif P_Class_01="라5" Then 
	    MsgP_01="5) 아이 메이크업 리무버"
	    
	    elseif P_Class_01="라6" Then 
	    MsgP_01="6) 그 밖의 눈화장용 제품류"
	    
	    elseif P_Class_01="마1" Then 
	    MsgP_01="1) 향수"
	    
	    elseif P_Class_01="마2" Then 
	    MsgP_01="2) 분말향"
	    
	    elseif P_Class_01="마3" Then 
	    MsgP_01="3) 향낭"
	    
	    elseif P_Class_01="마4" Then 
	    MsgP_01="4) 코롱"
	    
	    elseif P_Class_01="마5" Then 
	    MsgP_01="5) 그 밖의 방향용 제품류"
	    
	    elseif P_Class_01="바1" Then 
	    MsgP_01="1) 헤어 틴트"
	    
	    elseif P_Class_01="바2" Then 
	    MsgP_01="2) 헤어 칼라스프레이"
	    
	    elseif P_Class_01="바3" Then 
	    MsgP_01="3) 그 밖의 염모용 제품류"
	    
	    elseif P_Class_01="사1" Then 
	    MsgP_01="1) 볼연지"
	    
	    elseif P_Class_01="사2" Then 
	    MsgP_01="2) 페이스 파우더, 페이스 케익"
	    
	    elseif P_Class_01="사3" Then 
	    MsgP_01="3) 리퀴드, 크림, 케익 파운데이션"
	    
	    elseif P_Class_01="사4" Then 
	    MsgP_01="4) 메이크업 베이스"
	    
	    elseif P_Class_01="사5" Then 
	    MsgP_01="5) 메이크업 픽서티브"
	    
	    elseif P_Class_01="사6" Then 
	    MsgP_01="6) 립스틱, 립라이너"
	    
	    elseif P_Class_01="사7" Then 
	    MsgP_01="7) 립글로스, 립밤"
	    
	    elseif P_Class_01="사8" Then 
	    MsgP_01="8) 바디페인팅, 분장용 제품"
	    
	    elseif P_Class_01="사9" Then 
	    MsgP_01="9) 그 밖의 메이크업 제품류"
	    
	    elseif P_Class_01="아1" Then 
	    MsgP_01="1) 헤어 컨디셔너"
	    
	    elseif P_Class_01="아2" Then 
	    MsgP_01="2) 헤어 토닉"
	    
	    elseif P_Class_01="아3" Then 
	    MsgP_01="3) 헤어 그루밍에이드"
	    
	    elseif P_Class_01="아4" Then 
	    MsgP_01="4) 헤어 크림, 로션"
	    
	    elseif P_Class_01="아5" Then 
	    MsgP_01="5) 헤어 오일"
	    
	    elseif P_Class_01="아6" Then 
	    MsgP_01="6) 포마드"
	    
	    elseif P_Class_01="아7" Then 
	    MsgP_01="7) 헤어 스프레이ㆍ무스ㆍ왁스ㆍ젤"
	    
	    elseif P_Class_01="아8" Then 
	    MsgP_01="8) 샴푸, 린스"
	    
	    elseif P_Class_01="아9" Then 
	    MsgP_01="9) 퍼머넌트 웨이브"
	    
	    elseif P_Class_01="아10" Then 
	    MsgP_01="10) 헤어 스트레이트너"
	    
	    elseif P_Class_01="아11" Then 
	    MsgP_01="11) 그 밖의 두발용 제품류"
	    
	    elseif P_Class_01="자1" Then 
	    MsgP_01="1) 베이스코트, 언더코트"
	    
	    elseif P_Class_01="자2" Then 
	    MsgP_01="2) 네일폴리시, 네일에나멜"
	    
	    elseif P_Class_01="자3" Then 
	    MsgP_01="3) 탑코트"
	    
	    elseif P_Class_01="자4" Then 
	    MsgP_01="4) 네일 크림ㆍ로션ㆍ에센스"
	    
	    elseif P_Class_01="자5" Then 
	    MsgP_01="5) 네일폴리시ㆍ네일에나멜 리무버"
	    
	    elseif P_Class_01="자6" Then 
	    MsgP_01="6) 그 밖의 손발톱용 제품류"
	    
	    elseif P_Class_01="차1" Then 
	    MsgP_01="1) 애프터셰이브 로션"
	    
	    elseif P_Class_01="차2" Then 
	    MsgP_01="2) 남성용 탤컴"
	    
	    elseif P_Class_01="차3" Then 
	    MsgP_01="3) 프리셰이브 로션"
	    
	    elseif P_Class_01="차4" Then 
	    MsgP_01="4) 셰이빙 크림"
	    
	    elseif P_Class_01="차5" Then 
	    MsgP_01="5) 셰이빙 폼"
	    
	    elseif P_Class_01="차6" Then 
	    MsgP_01="6) 그 밖의 면도용 제품류"
	    
	    elseif P_Class_01="카1" Then 
	    MsgP_01="1) 수렴ㆍ유연ㆍ영양화장수"
	    
	    elseif P_Class_01="카2" Then 
	    MsgP_01="2) 마사지 크림"
	    
	    elseif P_Class_01="카3" Then 
	    MsgP_01="3) 에센스, 오일"
	    
	    elseif P_Class_01="카4" Then 
	    MsgP_01="4) 파우더"
	    
	    elseif P_Class_01="카5" Then 
	    MsgP_01="5) 바디 제품"
	    
	    elseif P_Class_01="카6" Then 
	    MsgP_01="6) 팩, 마스크"
	    
	    elseif P_Class_01="카7" Then 
	    MsgP_01="7) 눈 주위 제품"
	    
	    elseif P_Class_01="카8" Then 
	    MsgP_01="8) 로션, 크림"
	    
	    elseif P_Class_01="카9" Then 
	    MsgP_01="9) 손ㆍ발의 피부연화 제품"
	    
	    elseif P_Class_01="카10" Then 
	    MsgP_01="10) 클렌징워터ㆍ클렌징오일ㆍ클렌징로션ㆍ클렌징크림 등 메이크업 리무버"
	    
	    elseif P_Class_01="카11" Then 
	    MsgP_01="11) 그 밖의 기초화장용 제품류"
	    
	    elseif P_Class_01="타1" Then 
	    MsgP_01="1) 데오도런트"
	    
	    elseif P_Class_01="타2" Then 
	    MsgP_01="2) 그 밖의 체취 방지용 제품류"
	    
	    else                  
	   	  MsgP_01=""
      end If
	 
	 
	 
  IF      Functional_01="F1" Then 
          MsgF_01="미백"
  elseif  Functional_01="F2" Then
	   	    MsgF_01="주름개선"
  elseif  Functional_01="F3" Then
	   	    MsgF_01="자외선 차단"
	elseif  Functional_01="F4" Then
	   	    MsgF_01="미백, 주름개선"
	elseif  Functional_01="F5" Then
	   	    MsgF_01="미백, 자외선 차단"
  elseif  Functional_01="F6" Then
	   	    MsgF_01="주름개선, 자외선 차단"
  elseif  Functional_01="F7" Then
	   	    MsgF_01="미백, 주름개선, 자외선 차단"
  elseif  Functional_01="F8" Then
	   	    MsgF_01="염모"
  elseif  Functional_01="F9" Then
	   	    MsgF_01="제모"
  elseif  Functional_01="F10" Then
	   	    MsgF_01="탈모 완화"
  elseif  Functional_01="F11" Then
	   	    MsgF_01="여드름성 피부 완화"
  elseif  Functional_01="F12" Then
	   	    MsgF_01="아토피성 피부 보습"
  elseif  Functional_01="F13" Then
	   	    MsgF_01="튼살로 인한 붉은 선 완화"
  elseif  Functional_01="F14" Then
	   	    MsgF_01="기타 복합유형"
  elseif  Functional_01="일반" Then
	   	    MsgF_01="일반"
  else                  
	   	  MsgF_01=""
     end If
     
     
     
     
     
     
     
     Cosmetic_Class_Big_02 = LEFT(P_Class_02,1)

'유형
  IF Cosmetic_Class_Big_02="가" Then            
   		  Msg_02="영ㆍ유아용(만 3세이하의 어린이용을 말한다.) 제품류"
     elseif Cosmetic_Class_Big_02="나" Then     
         Msg_02="목욕용 제품류"
     elseif  Cosmetic_Class_Big_02="다" Then  
	   	   Msg_02="인체 세정용 제품류(화장비누 제외)"
	  elseif  Cosmetic_Class_Big_02="라" Then
	   	   Msg_02="눈 화장용 제품류"
    elseif  Cosmetic_Class_Big_02="마" Then
	   	   Msg_02="방향용 제품류"
	  elseif  Cosmetic_Class_Big_02="바" Then
	   	   Msg_02="두발 염색용 제품류"   
	  elseif  Cosmetic_Class_Big_02="사" Then
	   	   Msg_02="색조 화장용 제품"
	  elseif  Cosmetic_Class_Big_02="아" Then
	   	   Msg_02="두발용 제품류"   
	  elseif  Cosmetic_Class_Big_02="자" Then
	   	   Msg_02="손발톱용 제품류"
	  elseif  Cosmetic_Class_Big_02="차" Then
	   	   Msg_02="면도용 제품류"   
	  elseif  Cosmetic_Class_Big_02="카" Then
	   	   Msg_02="기초화장용 제품류"
	  elseif  Cosmetic_Class_Big_02="타" Then
	   	   Msg_02="체취 방지용 제품류" 
	 elseif  Cosmetic_Class_Big_02="파" Then
	   	   Msg_02="체모 제거용 제품류"
	   	   
	 else
	   	  Msg_02=""
     end If	   	   
	
	
	IF 	P_Class_02="가1" Then 
      MsgP_02="1) 영ㆍ유아용 샴푸, 린스"
	  
	    elseif P_Class_02="가2" Then 
	    MsgP_02="2) 영ㆍ유아용 로션, 크림"
	    
	    elseif P_Class_02="가3" Then 
	    MsgP_02="3) 영ㆍ유아용 오일"
	    
	    elseif P_Class_02="가4" Then 
	    MsgP_02="4) 영ㆍ유아용 인체 세정용 제품"
	    
	    elseif P_Class_02="가5" Then 
	    MsgP_02="5) 영ㆍ유아용 목욕용 제품"
	    
	    elseif P_Class_02="나1" Then 
	    MsgP_02="1) 목욕용 오일ㆍ정제ㆍ캡슐"
	    
	    elseif P_Class_02="나2" Then 
	    MsgP_02="2) 목욕용 소금류"
	    
	    elseif P_Class_02="나3" Then 
	    MsgP_02="3) 버블 배스"
	    
	    elseif P_Class_02="나4" Then 
	    MsgP_02="4) 그 밖의 목욕용 제품류"
	    
	    elseif P_Class_02="다1" Then 
	    MsgP_02="1) 폼 클렌저"
	    
	    elseif P_Class_02="다2" Then 
	    MsgP_02="2) 바디 클렌저"
	    
	    elseif P_Class_02="다3" Then 
	    MsgP_02="3) 액상비누"
	    
	    elseif P_Class_02="다4" Then 
	    MsgP_02="4) 외음부 세정제"
	    
	    elseif P_Class_02="다5" Then 
	    MsgP_02="5) 그 밖의 인체 세정용 제품류"
	    
	    elseif P_Class_02="라1" Then 
	    MsgP_02="1) 아이브로 펜슬"
	    
	    elseif P_Class_02="라2" Then 
	    MsgP_02="2) 아이 라이너"
	    
	    elseif P_Class_02="라3" Then 
	    MsgP_02="3) 아이 섀도"
	    
	    elseif P_Class_02="라4" Then 
	    MsgP_02="4) 마스카라"
	    
	    elseif P_Class_02="라5" Then 
	    MsgP_02="5) 아이 메이크업 리무버"
	    
	    elseif P_Class_02="라6" Then 
	    MsgP_02="6) 그 밖의 눈화장용 제품류"
	    
	    elseif P_Class_02="마1" Then 
	    MsgP_02="1) 향수"
	    
	    elseif P_Class_02="마2" Then 
	    MsgP_02="2) 분말향"
	    
	    elseif P_Class_02="마3" Then 
	    MsgP_02="3) 향낭"
	    
	    elseif P_Class_02="마4" Then 
	    MsgP_02="4) 코롱"
	    
	    elseif P_Class_02="마5" Then 
	    MsgP_02="5) 그 밖의 방향용 제품류"
	    
	    elseif P_Class_02="바1" Then 
	    MsgP_02="1) 헤어 틴트"
	    
	    elseif P_Class_02="바2" Then 
	    MsgP_02="2) 헤어 칼라스프레이"
	    
	    elseif P_Class_02="바3" Then 
	    MsgP_02="3) 그 밖의 염모용 제품류"
	    
	    elseif P_Class_02="사1" Then 
	    MsgP_02="1) 볼연지"
	    
	    elseif P_Class_02="사2" Then 
	    MsgP_02="2) 페이스 파우더, 페이스 케익"
	    
	    elseif P_Class_02="사3" Then 
	    MsgP_02="3) 리퀴드, 크림, 케익 파운데이션"
	    
	    elseif P_Class_02="사4" Then 
	    MsgP_02="4) 메이크업 베이스"
	    
	    elseif P_Class_02="사5" Then 
	    MsgP_02="5) 메이크업 픽서티브"
	    
	    elseif P_Class_02="사6" Then 
	    MsgP_02="6) 립스틱, 립라이너"
	    
	    elseif P_Class_02="사7" Then 
	    MsgP_02="7) 립글로스, 립밤"
	    
	    elseif P_Class_02="사8" Then 
	    MsgP_02="8) 바디페인팅, 분장용 제품"
	    
	    elseif P_Class_02="사9" Then 
	    MsgP_02="9) 그 밖의 메이크업 제품류"
	    
	    elseif P_Class_02="아1" Then 
	    MsgP_02="1) 헤어 컨디셔너"
	    
	    elseif P_Class_02="아2" Then 
	    MsgP_02="2) 헤어 토닉"
	    
	    elseif P_Class_02="아3" Then 
	    MsgP_02="3) 헤어 그루밍에이드"
	    
	    elseif P_Class_02="아4" Then 
	    MsgP_02="4) 헤어 크림, 로션"
	    
	    elseif P_Class_02="아5" Then 
	    MsgP_02="5) 헤어 오일"
	    
	    elseif P_Class_02="아6" Then 
	    MsgP_02="6) 포마드"
	    
	    elseif P_Class_02="아7" Then 
	    MsgP_02="7) 헤어 스프레이ㆍ무스ㆍ왁스ㆍ젤"
	    
	    elseif P_Class_02="아8" Then 
	    MsgP_02="8) 샴푸, 린스"
	    
	    elseif P_Class_02="아9" Then 
	    MsgP_02="9) 퍼머넌트 웨이브"
	    
	    elseif P_Class_02="아10" Then 
	    MsgP_02="10) 헤어 스트레이트너"
	    
	    elseif P_Class_02="아11" Then 
	    MsgP_02="11) 그 밖의 두발용 제품류"
	    
	    elseif P_Class_02="자1" Then 
	    MsgP_02="1) 베이스코트, 언더코트"
	    
	    elseif P_Class_02="자2" Then 
	    MsgP_02="2) 네일폴리시, 네일에나멜"
	    
	    elseif P_Class_02="자3" Then 
	    MsgP_02="3) 탑코트"
	    
	    elseif P_Class_02="자4" Then 
	    MsgP_02="4) 네일 크림ㆍ로션ㆍ에센스"
	    
	    elseif P_Class_02="자5" Then 
	    MsgP_02="5) 네일폴리시ㆍ네일에나멜 리무버"
	    
	    elseif P_Class_02="자6" Then 
	    MsgP_02="6) 그 밖의 손발톱용 제품류"
	    
	    elseif P_Class_02="차1" Then 
	    MsgP_02="1) 애프터셰이브 로션"
	    
	    elseif P_Class_02="차2" Then 
	    MsgP_02="2) 남성용 탤컴"
	    
	    elseif P_Class_02="차3" Then 
	    MsgP_02="3) 프리셰이브 로션"
	    
	    elseif P_Class_02="차4" Then 
	    MsgP_02="4) 셰이빙 크림"
	    
	    elseif P_Class_02="차5" Then 
	    MsgP_02="5) 셰이빙 폼"
	    
	    elseif P_Class_02="차6" Then 
	    MsgP_02="6) 그 밖의 면도용 제품류"
	    
	    elseif P_Class_02="카1" Then 
	    MsgP_02="1) 수렴ㆍ유연ㆍ영양화장수"
	    
	    elseif P_Class_02="카2" Then 
	    MsgP_02="2) 마사지 크림"
	    
	    elseif P_Class_02="카3" Then 
	    MsgP_02="3) 에센스, 오일"
	    
	    elseif P_Class_02="카4" Then 
	    MsgP_02="4) 파우더"
	    
	    elseif P_Class_02="카5" Then 
	    MsgP_02="5) 바디 제품"
	    
	    elseif P_Class_02="카6" Then 
	    MsgP_02="6) 팩, 마스크"
	    
	    elseif P_Class_02="카7" Then 
	    MsgP_02="7) 눈 주위 제품"
	    
	    elseif P_Class_02="카8" Then 
	    MsgP_02="8) 로션, 크림"
	    
	    elseif P_Class_02="카9" Then 
	    MsgP_02="9) 손ㆍ발의 피부연화 제품"
	    
	    elseif P_Class_02="카10" Then 
	    MsgP_02="10) 클렌징워터ㆍ클렌징오일ㆍ클렌징로션ㆍ클렌징크림 등 메이크업 리무버"
	    
	    elseif P_Class_02="카11" Then 
	    MsgP_02="11) 그 밖의 기초화장용 제품류"
	    
	    elseif P_Class_02="타1" Then 
	    MsgP_02="1) 데오도런트"
	    
	    elseif P_Class_02="타2" Then 
	    MsgP_02="2) 그 밖의 체취 방지용 제품류"
	    
	    else                  
	   	  MsgP_02=""
      end If
	 
	 
	 
  IF      Functional_02="F1" Then 
          MsgF_02="미백"
  elseif  Functional_02="F2" Then
	   	    MsgF_02="주름개선"
  elseif  Functional_02="F3" Then
	   	    MsgF_02="자외선 차단"
	elseif  Functional_02="F4" Then
	   	    MsgF_02="미백, 주름개선"
	elseif  Functional_02="F5" Then
	   	    MsgF_02="미백, 자외선 차단"
  elseif  Functional_02="F6" Then
	   	    MsgF_02="주름개선, 자외선 차단"
  elseif  Functional_02="F7" Then
	   	    MsgF_02="미백, 주름개선, 자외선 차단"
  elseif  Functional_02="F8" Then
	   	    MsgF_02="염모"
  elseif  Functional_02="F9" Then
	   	    MsgF_02="제모"
  elseif  Functional_02="F10" Then
	   	    MsgF_02="탈모 완화"
  elseif  Functional_02="F11" Then
	   	    MsgF_02="여드름성 피부 완화"
  elseif  Functional_02="F12" Then
	   	    MsgF_02="아토피성 피부 보습"
  elseif  Functional_02="F13" Then
	   	    MsgF_02="튼살로 인한 붉은 선 완화"
  elseif  Functional_02="F14" Then
	   	    MsgF_02="기타 복합유형"
  elseif  Functional_02="일반" Then
	   	    MsgF_02="일반"
  else                  
	   	  MsgF_02=""
     end If
     
     
     
     
     
     
     
     Cosmetic_Class_Big_03 = LEFT(P_Class_03,1)

'유형
  IF Cosmetic_Class_Big_03="가" Then            
   		  Msg_03="영ㆍ유아용(만 3세이하의 어린이용을 말한다.) 제품류"
     elseif Cosmetic_Class_Big_03="나" Then     
         Msg_03="목욕용 제품류"
     elseif  Cosmetic_Class_Big_03="다" Then  
	   	   Msg_03="인체 세정용 제품류(화장비누 제외)"
	  elseif  Cosmetic_Class_Big_03="라" Then
	   	   Msg_03="눈 화장용 제품류"
    elseif  Cosmetic_Class_Big_03="마" Then
	   	   Msg_03="방향용 제품류"
	  elseif  Cosmetic_Class_Big_03="바" Then
	   	   Msg_03="두발 염색용 제품류"   
	  elseif  Cosmetic_Class_Big_03="사" Then
	   	   Msg_03="색조 화장용 제품"
	  elseif  Cosmetic_Class_Big_03="아" Then
	   	   Msg_03="두발용 제품류"   
	  elseif  Cosmetic_Class_Big_03="자" Then
	   	   Msg_03="손발톱용 제품류"
	  elseif  Cosmetic_Class_Big_03="차" Then
	   	   Msg_03="면도용 제품류"   
	  elseif  Cosmetic_Class_Big_03="카" Then
	   	   Msg_03="기초화장용 제품류"
	  elseif  Cosmetic_Class_Big_03="타" Then
	   	   Msg_03="체취 방지용 제품류" 
	 elseif  Cosmetic_Class_Big_03="파" Then
	   	   Msg_03="체모 제거용 제품류"
	   	   
	 else
	   	  Msg_03=""
     end If	   	   
	
	
	IF 	P_Class_03="가1" Then 
      MsgP_03="1) 영ㆍ유아용 샴푸, 린스"
	  
	    elseif P_Class_03="가2" Then 
	    MsgP_03="2) 영ㆍ유아용 로션, 크림"
	    
	    elseif P_Class_03="가3" Then 
	    MsgP_03="3) 영ㆍ유아용 오일"
	    
	    elseif P_Class_03="가4" Then 
	    MsgP_03="4) 영ㆍ유아용 인체 세정용 제품"
	    
	    elseif P_Class_03="가5" Then 
	    MsgP_03="5) 영ㆍ유아용 목욕용 제품"
	    
	    elseif P_Class_03="나1" Then 
	    MsgP_03="1) 목욕용 오일ㆍ정제ㆍ캡슐"
	    
	    elseif P_Class_03="나2" Then 
	    MsgP_03="2) 목욕용 소금류"
	    
	    elseif P_Class_03="나3" Then 
	    MsgP_03="3) 버블 배스"
	    
	    elseif P_Class_03="나4" Then 
	    MsgP_03="4) 그 밖의 목욕용 제품류"
	    
	    elseif P_Class_03="다1" Then 
	    MsgP_03="1) 폼 클렌저"
	    
	    elseif P_Class_03="다2" Then 
	    MsgP_03="2) 바디 클렌저"
	    
	    elseif P_Class_03="다3" Then 
	    MsgP_03="3) 액상비누"
	    
	    elseif P_Class_03="다4" Then 
	    MsgP_03="4) 외음부 세정제"
	    
	    elseif P_Class_03="다5" Then 
	    MsgP_03="5) 그 밖의 인체 세정용 제품류"
	    
	    elseif P_Class_03="라1" Then 
	    MsgP_03="1) 아이브로 펜슬"
	    
	    elseif P_Class_03="라2" Then 
	    MsgP_03="2) 아이 라이너"
	    
	    elseif P_Class_03="라3" Then 
	    MsgP_03="3) 아이 섀도"
	    
	    elseif P_Class_03="라4" Then 
	    MsgP_03="4) 마스카라"
	    
	    elseif P_Class_03="라5" Then 
	    MsgP_03="5) 아이 메이크업 리무버"
	    
	    elseif P_Class_03="라6" Then 
	    MsgP_03="6) 그 밖의 눈화장용 제품류"
	    
	    elseif P_Class_03="마1" Then 
	    MsgP_03="1) 향수"
	    
	    elseif P_Class_03="마2" Then 
	    MsgP_03="2) 분말향"
	    
	    elseif P_Class_03="마3" Then 
	    MsgP_03="3) 향낭"
	    
	    elseif P_Class_03="마4" Then 
	    MsgP_03="4) 코롱"
	    
	    elseif P_Class_03="마5" Then 
	    MsgP_03="5) 그 밖의 방향용 제품류"
	    
	    elseif P_Class_03="바1" Then 
	    MsgP_03="1) 헤어 틴트"
	    
	    elseif P_Class_03="바2" Then 
	    MsgP_03="2) 헤어 칼라스프레이"
	    
	    elseif P_Class_03="바3" Then 
	    MsgP_03="3) 그 밖의 염모용 제품류"
	    
	    elseif P_Class_03="사1" Then 
	    MsgP_03="1) 볼연지"
	    
	    elseif P_Class_03="사2" Then 
	    MsgP_03="2) 페이스 파우더, 페이스 케익"
	    
	    elseif P_Class_03="사3" Then 
	    MsgP_03="3) 리퀴드, 크림, 케익 파운데이션"
	    
	    elseif P_Class_03="사4" Then 
	    MsgP_03="4) 메이크업 베이스"
	    
	    elseif P_Class_03="사5" Then 
	    MsgP_03="5) 메이크업 픽서티브"
	    
	    elseif P_Class_03="사6" Then 
	    MsgP_03="6) 립스틱, 립라이너"
	    
	    elseif P_Class_03="사7" Then 
	    MsgP_03="7) 립글로스, 립밤"
	    
	    elseif P_Class_03="사8" Then 
	    MsgP_03="8) 바디페인팅, 분장용 제품"
	    
	    elseif P_Class_03="사9" Then 
	    MsgP_03="9) 그 밖의 메이크업 제품류"
	    
	    elseif P_Class_03="아1" Then 
	    MsgP_03="1) 헤어 컨디셔너"
	    
	    elseif P_Class_03="아2" Then 
	    MsgP_03="2) 헤어 토닉"
	    
	    elseif P_Class_03="아3" Then 
	    MsgP_03="3) 헤어 그루밍에이드"
	    
	    elseif P_Class_03="아4" Then 
	    MsgP_03="4) 헤어 크림, 로션"
	    
	    elseif P_Class_03="아5" Then 
	    MsgP_03="5) 헤어 오일"
	    
	    elseif P_Class_03="아6" Then 
	    MsgP_03="6) 포마드"
	    
	    elseif P_Class_03="아7" Then 
	    MsgP_03="7) 헤어 스프레이ㆍ무스ㆍ왁스ㆍ젤"
	    
	    elseif P_Class_03="아8" Then 
	    MsgP_03="8) 샴푸, 린스"
	    
	    elseif P_Class_03="아9" Then 
	    MsgP_03="9) 퍼머넌트 웨이브"
	    
	    elseif P_Class_03="아10" Then 
	    MsgP_03="10) 헤어 스트레이트너"
	    
	    elseif P_Class_03="아11" Then 
	    MsgP_03="11) 그 밖의 두발용 제품류"
	    
	    elseif P_Class_03="자1" Then 
	    MsgP_03="1) 베이스코트, 언더코트"
	    
	    elseif P_Class_03="자2" Then 
	    MsgP_03="2) 네일폴리시, 네일에나멜"
	    
	    elseif P_Class_03="자3" Then 
	    MsgP_03="3) 탑코트"
	    
	    elseif P_Class_03="자4" Then 
	    MsgP_03="4) 네일 크림ㆍ로션ㆍ에센스"
	    
	    elseif P_Class_03="자5" Then 
	    MsgP_03="5) 네일폴리시ㆍ네일에나멜 리무버"
	    
	    elseif P_Class_03="자6" Then 
	    MsgP_03="6) 그 밖의 손발톱용 제품류"
	    
	    elseif P_Class_03="차1" Then 
	    MsgP_03="1) 애프터셰이브 로션"
	    
	    elseif P_Class_03="차2" Then 
	    MsgP_03="2) 남성용 탤컴"
	    
	    elseif P_Class_03="차3" Then 
	    MsgP_03="3) 프리셰이브 로션"
	    
	    elseif P_Class_03="차4" Then 
	    MsgP_03="4) 셰이빙 크림"
	    
	    elseif P_Class_03="차5" Then 
	    MsgP_03="5) 셰이빙 폼"
	    
	    elseif P_Class_03="차6" Then 
	    MsgP_03="6) 그 밖의 면도용 제품류"
	    
	    elseif P_Class_03="카1" Then 
	    MsgP_03="1) 수렴ㆍ유연ㆍ영양화장수"
	    
	    elseif P_Class_03="카2" Then 
	    MsgP_03="2) 마사지 크림"
	    
	    elseif P_Class_03="카3" Then 
	    MsgP_03="3) 에센스, 오일"
	    
	    elseif P_Class_03="카4" Then 
	    MsgP_03="4) 파우더"
	    
	    elseif P_Class_03="카5" Then 
	    MsgP_03="5) 바디 제품"
	    
	    elseif P_Class_03="카6" Then 
	    MsgP_03="6) 팩, 마스크"
	    
	    elseif P_Class_03="카7" Then 
	    MsgP_03="7) 눈 주위 제품"
	    
	    elseif P_Class_03="카8" Then 
	    MsgP_03="8) 로션, 크림"
	    
	    elseif P_Class_03="카9" Then 
	    MsgP_03="9) 손ㆍ발의 피부연화 제품"
	    
	    elseif P_Class_03="카10" Then 
	    MsgP_03="10) 클렌징워터ㆍ클렌징오일ㆍ클렌징로션ㆍ클렌징크림 등 메이크업 리무버"
	    
	    elseif P_Class_03="카11" Then 
	    MsgP_03="11) 그 밖의 기초화장용 제품류"
	    
	    elseif P_Class_03="타1" Then 
	    MsgP_03="1) 데오도런트"
	    
	    elseif P_Class_03="타2" Then 
	    MsgP_03="2) 그 밖의 체취 방지용 제품류"
	    
	    else                  
	   	  MsgP_03=""
      end If
	 
	 
	 
  IF      Functional_03="F1" Then 
          MsgF_03="미백"
  elseif  Functional_03="F2" Then
	   	    MsgF_03="주름개선"
  elseif  Functional_03="F3" Then
	   	    MsgF_03="자외선 차단"
	elseif  Functional_03="F4" Then
	   	    MsgF_03="미백, 주름개선"
	elseif  Functional_03="F5" Then
	   	    MsgF_03="미백, 자외선 차단"
  elseif  Functional_03="F6" Then
	   	    MsgF_03="주름개선, 자외선 차단"
  elseif  Functional_03="F7" Then
	   	    MsgF_03="미백, 주름개선, 자외선 차단"
  elseif  Functional_03="F8" Then
	   	    MsgF_03="염모"
  elseif  Functional_03="F9" Then
	   	    MsgF_03="제모"
  elseif  Functional_03="F10" Then
	   	    MsgF_03="탈모 완화"
  elseif  Functional_03="F11" Then
	   	    MsgF_03="여드름성 피부 완화"
  elseif  Functional_03="F12" Then
	   	    MsgF_03="아토피성 피부 보습"
  elseif  Functional_03="F13" Then
	   	    MsgF_03="튼살로 인한 붉은 선 완화"
  elseif  Functional_03="F14" Then
	   	    MsgF_03="기타 복합유형"
  elseif  Functional_03="일반" Then
	   	    MsgF_03="일반"
  else                  
	   	  MsgF_03=""
     end If	  
          
     
     
     
     
     
     
     Cosmetic_Class_Big_04 = LEFT(P_Class_04,1)

'유형
  IF Cosmetic_Class_Big_04="가" Then            
   		  Msg_04="영ㆍ유아용(만 3세이하의 어린이용을 말한다.) 제품류"
     elseif Cosmetic_Class_Big_04="나" Then     
         Msg_04="목욕용 제품류"
     elseif  Cosmetic_Class_Big_04="다" Then  
	   	   Msg_04="인체 세정용 제품류(화장비누 제외)"
	  elseif  Cosmetic_Class_Big_04="라" Then
	   	   Msg_04="눈 화장용 제품류"
    elseif  Cosmetic_Class_Big_04="마" Then
	   	   Msg_04="방향용 제품류"
	  elseif  Cosmetic_Class_Big_04="바" Then
	   	   Msg_04="두발 염색용 제품류"   
	  elseif  Cosmetic_Class_Big_04="사" Then
	   	   Msg_04="색조 화장용 제품"
	  elseif  Cosmetic_Class_Big_04="아" Then
	   	   Msg_04="두발용 제품류"   
	  elseif  Cosmetic_Class_Big_04="자" Then
	   	   Msg_04="손발톱용 제품류"
	  elseif  Cosmetic_Class_Big_04="차" Then
	   	   Msg_04="면도용 제품류"   
	  elseif  Cosmetic_Class_Big_04="카" Then
	   	   Msg_04="기초화장용 제품류"
	  elseif  Cosmetic_Class_Big_04="타" Then
	   	   Msg_04="체취 방지용 제품류" 
	 elseif  Cosmetic_Class_Big_04="파" Then
	   	   Msg_04="체모 제거용 제품류"
	   	   
	 else
	   	  Msg_04=""
     end If	   	   
	
	
	IF 	P_Class_04="가1" Then 
      MsgP_04="1) 영ㆍ유아용 샴푸, 린스"
	  
	    elseif P_Class_04="가2" Then 
	    MsgP_04="2) 영ㆍ유아용 로션, 크림"
	    
	    elseif P_Class_04="가3" Then 
	    MsgP_04="3) 영ㆍ유아용 오일"
	    
	    elseif P_Class_04="가4" Then 
	    MsgP_04="4) 영ㆍ유아용 인체 세정용 제품"
	    
	    elseif P_Class_04="가5" Then 
	    MsgP_04="5) 영ㆍ유아용 목욕용 제품"
	    
	    elseif P_Class_04="나1" Then 
	    MsgP_04="1) 목욕용 오일ㆍ정제ㆍ캡슐"
	    
	    elseif P_Class_04="나2" Then 
	    MsgP_04="2) 목욕용 소금류"
	    
	    elseif P_Class_04="나3" Then 
	    MsgP_04="3) 버블 배스"
	    
	    elseif P_Class_04="나4" Then 
	    MsgP_04="4) 그 밖의 목욕용 제품류"
	    
	    elseif P_Class_04="다1" Then 
	    MsgP_04="1) 폼 클렌저"
	    
	    elseif P_Class_04="다2" Then 
	    MsgP_04="2) 바디 클렌저"
	    
	    elseif P_Class_04="다3" Then 
	    MsgP_04="3) 액상비누"
	    
	    elseif P_Class_04="다4" Then 
	    MsgP_04="4) 외음부 세정제"
	    
	    elseif P_Class_04="다5" Then 
	    MsgP_04="5) 그 밖의 인체 세정용 제품류"
	    
	    elseif P_Class_04="라1" Then 
	    MsgP_04="1) 아이브로 펜슬"
	    
	    elseif P_Class_04="라2" Then 
	    MsgP_04="2) 아이 라이너"
	    
	    elseif P_Class_04="라3" Then 
	    MsgP_04="3) 아이 섀도"
	    
	    elseif P_Class_04="라4" Then 
	    MsgP_04="4) 마스카라"
	    
	    elseif P_Class_04="라5" Then 
	    MsgP_04="5) 아이 메이크업 리무버"
	    
	    elseif P_Class_04="라6" Then 
	    MsgP_04="6) 그 밖의 눈화장용 제품류"
	    
	    elseif P_Class_04="마1" Then 
	    MsgP_04="1) 향수"
	    
	    elseif P_Class_04="마2" Then 
	    MsgP_04="2) 분말향"
	    
	    elseif P_Class_04="마3" Then 
	    MsgP_04="3) 향낭"
	    
	    elseif P_Class_04="마4" Then 
	    MsgP_04="4) 코롱"
	    
	    elseif P_Class_04="마5" Then 
	    MsgP_04="5) 그 밖의 방향용 제품류"
	    
	    elseif P_Class_04="바1" Then 
	    MsgP_04="1) 헤어 틴트"
	    
	    elseif P_Class_04="바2" Then 
	    MsgP_04="2) 헤어 칼라스프레이"
	    
	    elseif P_Class_04="바3" Then 
	    MsgP_04="3) 그 밖의 염모용 제품류"
	    
	    elseif P_Class_04="사1" Then 
	    MsgP_04="1) 볼연지"
	    
	    elseif P_Class_04="사2" Then 
	    MsgP_04="2) 페이스 파우더, 페이스 케익"
	    
	    elseif P_Class_04="사3" Then 
	    MsgP_04="3) 리퀴드, 크림, 케익 파운데이션"
	    
	    elseif P_Class_04="사4" Then 
	    MsgP_04="4) 메이크업 베이스"
	    
	    elseif P_Class_04="사5" Then 
	    MsgP_04="5) 메이크업 픽서티브"
	    
	    elseif P_Class_04="사6" Then 
	    MsgP_04="6) 립스틱, 립라이너"
	    
	    elseif P_Class_04="사7" Then 
	    MsgP_04="7) 립글로스, 립밤"
	    
	    elseif P_Class_04="사8" Then 
	    MsgP_04="8) 바디페인팅, 분장용 제품"
	    
	    elseif P_Class_04="사9" Then 
	    MsgP_04="9) 그 밖의 메이크업 제품류"
	    
	    elseif P_Class_04="아1" Then 
	    MsgP_04="1) 헤어 컨디셔너"
	    
	    elseif P_Class_04="아2" Then 
	    MsgP_04="2) 헤어 토닉"
	    
	    elseif P_Class_04="아3" Then 
	    MsgP_04="3) 헤어 그루밍에이드"
	    
	    elseif P_Class_04="아4" Then 
	    MsgP_04="4) 헤어 크림, 로션"
	    
	    elseif P_Class_04="아5" Then 
	    MsgP_04="5) 헤어 오일"
	    
	    elseif P_Class_04="아6" Then 
	    MsgP_04="6) 포마드"
	    
	    elseif P_Class_04="아7" Then 
	    MsgP_04="7) 헤어 스프레이ㆍ무스ㆍ왁스ㆍ젤"
	    
	    elseif P_Class_04="아8" Then 
	    MsgP_04="8) 샴푸, 린스"
	    
	    elseif P_Class_04="아9" Then 
	    MsgP_04="9) 퍼머넌트 웨이브"
	    
	    elseif P_Class_04="아10" Then 
	    MsgP_04="10) 헤어 스트레이트너"
	    
	    elseif P_Class_04="아11" Then 
	    MsgP_04="11) 그 밖의 두발용 제품류"
	    
	    elseif P_Class_04="자1" Then 
	    MsgP_04="1) 베이스코트, 언더코트"
	    
	    elseif P_Class_04="자2" Then 
	    MsgP_04="2) 네일폴리시, 네일에나멜"
	    
	    elseif P_Class_04="자3" Then 
	    MsgP_04="3) 탑코트"
	    
	    elseif P_Class_04="자4" Then 
	    MsgP_04="4) 네일 크림ㆍ로션ㆍ에센스"
	    
	    elseif P_Class_04="자5" Then 
	    MsgP_04="5) 네일폴리시ㆍ네일에나멜 리무버"
	    
	    elseif P_Class_04="자6" Then 
	    MsgP_04="6) 그 밖의 손발톱용 제품류"
	    
	    elseif P_Class_04="차1" Then 
	    MsgP_04="1) 애프터셰이브 로션"
	    
	    elseif P_Class_04="차2" Then 
	    MsgP_04="2) 남성용 탤컴"
	    
	    elseif P_Class_04="차3" Then 
	    MsgP_04="3) 프리셰이브 로션"
	    
	    elseif P_Class_04="차4" Then 
	    MsgP_04="4) 셰이빙 크림"
	    
	    elseif P_Class_04="차5" Then 
	    MsgP_04="5) 셰이빙 폼"
	    
	    elseif P_Class_04="차6" Then 
	    MsgP_04="6) 그 밖의 면도용 제품류"
	    
	    elseif P_Class_04="카1" Then 
	    MsgP_04="1) 수렴ㆍ유연ㆍ영양화장수"
	    
	    elseif P_Class_04="카2" Then 
	    MsgP_04="2) 마사지 크림"
	    
	    elseif P_Class_04="카3" Then 
	    MsgP_04="3) 에센스, 오일"
	    
	    elseif P_Class_04="카4" Then 
	    MsgP_04="4) 파우더"
	    
	    elseif P_Class_04="카5" Then 
	    MsgP_04="5) 바디 제품"
	    
	    elseif P_Class_04="카6" Then 
	    MsgP_04="6) 팩, 마스크"
	    
	    elseif P_Class_04="카7" Then 
	    MsgP_04="7) 눈 주위 제품"
	    
	    elseif P_Class_04="카8" Then 
	    MsgP_04="8) 로션, 크림"
	    
	    elseif P_Class_04="카9" Then 
	    MsgP_04="9) 손ㆍ발의 피부연화 제품"
	    
	    elseif P_Class_04="카10" Then 
	    MsgP_04="10) 클렌징워터ㆍ클렌징오일ㆍ클렌징로션ㆍ클렌징크림 등 메이크업 리무버"
	    
	    elseif P_Class_04="카11" Then 
	    MsgP_04="11) 그 밖의 기초화장용 제품류"
	    
	    elseif P_Class_04="타1" Then 
	    MsgP_04="1) 데오도런트"
	    
	    elseif P_Class_04="타2" Then 
	    MsgP_04="2) 그 밖의 체취 방지용 제품류"
	    
	    else                  
	   	  MsgP_04=""
      end If
	 
	 
	 
  IF      Functional_04="F1" Then 
          MsgF_04="미백"
  elseif  Functional_04="F2" Then
	   	    MsgF_04="주름개선"
  elseif  Functional_04="F3" Then
	   	    MsgF_04="자외선 차단"
	elseif  Functional_04="F4" Then
	   	    MsgF_04="미백, 주름개선"
	elseif  Functional_04="F5" Then
	   	    MsgF_04="미백, 자외선 차단"
  elseif  Functional_04="F6" Then
	   	    MsgF_04="주름개선, 자외선 차단"
  elseif  Functional_04="F7" Then
	   	    MsgF_04="미백, 주름개선, 자외선 차단"
  elseif  Functional_04="F8" Then
	   	    MsgF_04="염모"
  elseif  Functional_04="F9" Then
	   	    MsgF_04="제모"
  elseif  Functional_04="F10" Then
	   	    MsgF_04="탈모 완화"
  elseif  Functional_04="F11" Then
	   	    MsgF_04="여드름성 피부 완화"
  elseif  Functional_04="F12" Then
	   	    MsgF_04="아토피성 피부 보습"
  elseif  Functional_04="F13" Then
	   	    MsgF_04="튼살로 인한 붉은 선 완화"
  elseif  Functional_04="F14" Then
	   	    MsgF_04="기타 복합유형"
  elseif  Functional_04="일반" Then
	   	    MsgF_04="일반"
  else                  
	   	  MsgF_04=""
     end If
  
       
     
     
     
     
     
     
     Cosmetic_Class_Big_05 = LEFT(P_Class_05,1)

'유형
  IF Cosmetic_Class_Big_05="가" Then            
   		  Msg_05="영ㆍ유아용(만 3세이하의 어린이용을 말한다.) 제품류"
     elseif Cosmetic_Class_Big_05="나" Then     
         Msg_05="목욕용 제품류"
     elseif  Cosmetic_Class_Big_05="다" Then  
	   	   Msg_05="인체 세정용 제품류(화장비누 제외)"
	  elseif  Cosmetic_Class_Big_05="라" Then
	   	   Msg_05="눈 화장용 제품류"
    elseif  Cosmetic_Class_Big_05="마" Then
	   	   Msg_05="방향용 제품류"
	  elseif  Cosmetic_Class_Big_05="바" Then
	   	   Msg_05="두발 염색용 제품류"   
	  elseif  Cosmetic_Class_Big_05="사" Then
	   	   Msg_05="색조 화장용 제품"
	  elseif  Cosmetic_Class_Big_05="아" Then
	   	   Msg_05="두발용 제품류"   
	  elseif  Cosmetic_Class_Big_05="자" Then
	   	   Msg_05="손발톱용 제품류"
	  elseif  Cosmetic_Class_Big_05="차" Then
	   	   Msg_05="면도용 제품류"   
	  elseif  Cosmetic_Class_Big_05="카" Then
	   	   Msg_05="기초화장용 제품류"
	  elseif  Cosmetic_Class_Big_05="타" Then
	   	   Msg_05="체취 방지용 제품류" 
	 elseif  Cosmetic_Class_Big_05="파" Then
	   	   Msg_05="체모 제거용 제품류"
	   	   
	 else
	   	  Msg_05=""
     end If	   	   
	
	
	IF 	P_Class_05="가1" Then 
      MsgP_05="1) 영ㆍ유아용 샴푸, 린스"
	  
	    elseif P_Class_05="가2" Then 
	    MsgP_05="2) 영ㆍ유아용 로션, 크림"
	    
	    elseif P_Class_05="가3" Then 
	    MsgP_05="3) 영ㆍ유아용 오일"
	    
	    elseif P_Class_05="가4" Then 
	    MsgP_05="4) 영ㆍ유아용 인체 세정용 제품"
	    
	    elseif P_Class_05="가5" Then 
	    MsgP_05="5) 영ㆍ유아용 목욕용 제품"
	    
	    elseif P_Class_05="나1" Then 
	    MsgP_05="1) 목욕용 오일ㆍ정제ㆍ캡슐"
	    
	    elseif P_Class_05="나2" Then 
	    MsgP_05="2) 목욕용 소금류"
	    
	    elseif P_Class_05="나3" Then 
	    MsgP_05="3) 버블 배스"
	    
	    elseif P_Class_05="나4" Then 
	    MsgP_05="4) 그 밖의 목욕용 제품류"
	    
	    elseif P_Class_05="다1" Then 
	    MsgP_05="1) 폼 클렌저"
	    
	    elseif P_Class_05="다2" Then 
	    MsgP_05="2) 바디 클렌저"
	    
	    elseif P_Class_05="다3" Then 
	    MsgP_05="3) 액상비누"
	    
	    elseif P_Class_05="다4" Then 
	    MsgP_05="4) 외음부 세정제"
	    
	    elseif P_Class_05="다5" Then 
	    MsgP_05="5) 그 밖의 인체 세정용 제품류"
	    
	    elseif P_Class_05="라1" Then 
	    MsgP_05="1) 아이브로 펜슬"
	    
	    elseif P_Class_05="라2" Then 
	    MsgP_05="2) 아이 라이너"
	    
	    elseif P_Class_05="라3" Then 
	    MsgP_05="3) 아이 섀도"
	    
	    elseif P_Class_05="라4" Then 
	    MsgP_05="4) 마스카라"
	    
	    elseif P_Class_05="라5" Then 
	    MsgP_05="5) 아이 메이크업 리무버"
	    
	    elseif P_Class_05="라6" Then 
	    MsgP_05="6) 그 밖의 눈화장용 제품류"
	    
	    elseif P_Class_05="마1" Then 
	    MsgP_05="1) 향수"
	    
	    elseif P_Class_05="마2" Then 
	    MsgP_05="2) 분말향"
	    
	    elseif P_Class_05="마3" Then 
	    MsgP_05="3) 향낭"
	    
	    elseif P_Class_05="마4" Then 
	    MsgP_05="4) 코롱"
	    
	    elseif P_Class_05="마5" Then 
	    MsgP_05="5) 그 밖의 방향용 제품류"
	    
	    elseif P_Class_05="바1" Then 
	    MsgP_05="1) 헤어 틴트"
	    
	    elseif P_Class_05="바2" Then 
	    MsgP_05="2) 헤어 칼라스프레이"
	    
	    elseif P_Class_05="바3" Then 
	    MsgP_05="3) 그 밖의 염모용 제품류"
	    
	    elseif P_Class_05="사1" Then 
	    MsgP_05="1) 볼연지"
	    
	    elseif P_Class_05="사2" Then 
	    MsgP_05="2) 페이스 파우더, 페이스 케익"
	    
	    elseif P_Class_05="사3" Then 
	    MsgP_05="3) 리퀴드, 크림, 케익 파운데이션"
	    
	    elseif P_Class_05="사4" Then 
	    MsgP_05="4) 메이크업 베이스"
	    
	    elseif P_Class_05="사5" Then 
	    MsgP_05="5) 메이크업 픽서티브"
	    
	    elseif P_Class_05="사6" Then 
	    MsgP_05="6) 립스틱, 립라이너"
	    
	    elseif P_Class_05="사7" Then 
	    MsgP_05="7) 립글로스, 립밤"
	    
	    elseif P_Class_05="사8" Then 
	    MsgP_05="8) 바디페인팅, 분장용 제품"
	    
	    elseif P_Class_05="사9" Then 
	    MsgP_05="9) 그 밖의 메이크업 제품류"
	    
	    elseif P_Class_05="아1" Then 
	    MsgP_05="1) 헤어 컨디셔너"
	    
	    elseif P_Class_05="아2" Then 
	    MsgP_05="2) 헤어 토닉"
	    
	    elseif P_Class_05="아3" Then 
	    MsgP_05="3) 헤어 그루밍에이드"
	    
	    elseif P_Class_05="아4" Then 
	    MsgP_05="4) 헤어 크림, 로션"
	    
	    elseif P_Class_05="아5" Then 
	    MsgP_05="5) 헤어 오일"
	    
	    elseif P_Class_05="아6" Then 
	    MsgP_05="6) 포마드"
	    
	    elseif P_Class_05="아7" Then 
	    MsgP_05="7) 헤어 스프레이ㆍ무스ㆍ왁스ㆍ젤"
	    
	    elseif P_Class_05="아8" Then 
	    MsgP_05="8) 샴푸, 린스"
	    
	    elseif P_Class_05="아9" Then 
	    MsgP_05="9) 퍼머넌트 웨이브"
	    
	    elseif P_Class_05="아10" Then 
	    MsgP_05="10) 헤어 스트레이트너"
	    
	    elseif P_Class_05="아11" Then 
	    MsgP_05="11) 그 밖의 두발용 제품류"
	    
	    elseif P_Class_05="자1" Then 
	    MsgP_05="1) 베이스코트, 언더코트"
	    
	    elseif P_Class_05="자2" Then 
	    MsgP_05="2) 네일폴리시, 네일에나멜"
	    
	    elseif P_Class_05="자3" Then 
	    MsgP_05="3) 탑코트"
	    
	    elseif P_Class_05="자4" Then 
	    MsgP_05="4) 네일 크림ㆍ로션ㆍ에센스"
	    
	    elseif P_Class_05="자5" Then 
	    MsgP_05="5) 네일폴리시ㆍ네일에나멜 리무버"
	    
	    elseif P_Class_05="자6" Then 
	    MsgP_05="6) 그 밖의 손발톱용 제품류"
	    
	    elseif P_Class_05="차1" Then 
	    MsgP_05="1) 애프터셰이브 로션"
	    
	    elseif P_Class_05="차2" Then 
	    MsgP_05="2) 남성용 탤컴"
	    
	    elseif P_Class_05="차3" Then 
	    MsgP_05="3) 프리셰이브 로션"
	    
	    elseif P_Class_05="차4" Then 
	    MsgP_05="4) 셰이빙 크림"
	    
	    elseif P_Class_05="차5" Then 
	    MsgP_05="5) 셰이빙 폼"
	    
	    elseif P_Class_05="차6" Then 
	    MsgP_05="6) 그 밖의 면도용 제품류"
	    
	    elseif P_Class_05="카1" Then 
	    MsgP_05="1) 수렴ㆍ유연ㆍ영양화장수"
	    
	    elseif P_Class_05="카2" Then 
	    MsgP_05="2) 마사지 크림"
	    
	    elseif P_Class_05="카3" Then 
	    MsgP_05="3) 에센스, 오일"
	    
	    elseif P_Class_05="카4" Then 
	    MsgP_05="4) 파우더"
	    
	    elseif P_Class_05="카5" Then 
	    MsgP_05="5) 바디 제품"
	    
	    elseif P_Class_05="카6" Then 
	    MsgP_05="6) 팩, 마스크"
	    
	    elseif P_Class_05="카7" Then 
	    MsgP_05="7) 눈 주위 제품"
	    
	    elseif P_Class_05="카8" Then 
	    MsgP_05="8) 로션, 크림"
	    
	    elseif P_Class_05="카9" Then 
	    MsgP_05="9) 손ㆍ발의 피부연화 제품"
	    
	    elseif P_Class_05="카10" Then 
	    MsgP_05="10) 클렌징워터ㆍ클렌징오일ㆍ클렌징로션ㆍ클렌징크림 등 메이크업 리무버"
	    
	    elseif P_Class_05="카11" Then 
	    MsgP_05="11) 그 밖의 기초화장용 제품류"
	    
	    elseif P_Class_05="타1" Then 
	    MsgP_05="1) 데오도런트"
	    
	    elseif P_Class_05="타2" Then 
	    MsgP_05="2) 그 밖의 체취 방지용 제품류"
	    
	    else                  
	   	  MsgP_05=""
      end If
	 
	 
	 
  IF      Functional_05="F1" Then 
          MsgF_05="미백"
  elseif  Functional_05="F2" Then
	   	    MsgF_05="주름개선"
  elseif  Functional_05="F3" Then
	   	    MsgF_05="자외선 차단"
	elseif  Functional_05="F4" Then
	   	    MsgF_05="미백, 주름개선"
	elseif  Functional_05="F5" Then
	   	    MsgF_05="미백, 자외선 차단"
  elseif  Functional_05="F6" Then
	   	    MsgF_05="주름개선, 자외선 차단"
  elseif  Functional_05="F7" Then
	   	    MsgF_05="미백, 주름개선, 자외선 차단"
  elseif  Functional_05="F8" Then
	   	    MsgF_05="염모"
  elseif  Functional_05="F9" Then
	   	    MsgF_05="제모"
  elseif  Functional_05="F10" Then
	   	    MsgF_05="탈모 완화"
  elseif  Functional_05="F11" Then
	   	    MsgF_05="여드름성 피부 완화"
  elseif  Functional_05="F12" Then
	   	    MsgF_05="아토피성 피부 보습"
  elseif  Functional_05="F13" Then
	   	    MsgF_05="튼살로 인한 붉은 선 완화"
  elseif  Functional_05="F14" Then
	   	    MsgF_05="기타 복합유형"
  elseif  Functional_05="일반" Then
	   	    MsgF_05="일반"
  else                  
	   	  MsgF_05=""
     end If	 
     
       
     
     
     
     
     
     
     Cosmetic_Class_Big_06 = LEFT(P_Class_06,1)

'유형
  IF Cosmetic_Class_Big_06="가" Then            
   		  Msg_06="영ㆍ유아용(만 3세이하의 어린이용을 말한다.) 제품류"
     elseif Cosmetic_Class_Big_06="나" Then     
         Msg_06="목욕용 제품류"
     elseif  Cosmetic_Class_Big_06="다" Then  
	   	   Msg_06="인체 세정용 제품류(화장비누 제외)"
	  elseif  Cosmetic_Class_Big_06="라" Then
	   	   Msg_06="눈 화장용 제품류"
    elseif  Cosmetic_Class_Big_06="마" Then
	   	   Msg_06="방향용 제품류"
	  elseif  Cosmetic_Class_Big_06="바" Then
	   	   Msg_06="두발 염색용 제품류"   
	  elseif  Cosmetic_Class_Big_06="사" Then
	   	   Msg_06="색조 화장용 제품"
	  elseif  Cosmetic_Class_Big_06="아" Then
	   	   Msg_06="두발용 제품류"   
	  elseif  Cosmetic_Class_Big_06="자" Then
	   	   Msg_06="손발톱용 제품류"
	  elseif  Cosmetic_Class_Big_06="차" Then
	   	   Msg_06="면도용 제품류"   
	  elseif  Cosmetic_Class_Big_06="카" Then
	   	   Msg_06="기초화장용 제품류"
	  elseif  Cosmetic_Class_Big_06="타" Then
	   	   Msg_06="체취 방지용 제품류" 
	 elseif  Cosmetic_Class_Big_06="파" Then
	   	   Msg_06="체모 제거용 제품류"
	   	   
	 else
	   	  Msg_06=""
     end If	   	   
	
	
	IF 	P_Class_06="가1" Then 
      MsgP_06="1) 영ㆍ유아용 샴푸, 린스"
	  
	    elseif P_Class_06="가2" Then 
	    MsgP_06="2) 영ㆍ유아용 로션, 크림"
	    
	    elseif P_Class_06="가3" Then 
	    MsgP_06="3) 영ㆍ유아용 오일"
	    
	    elseif P_Class_06="가4" Then 
	    MsgP_06="4) 영ㆍ유아용 인체 세정용 제품"
	    
	    elseif P_Class_06="가5" Then 
	    MsgP_06="5) 영ㆍ유아용 목욕용 제품"
	    
	    elseif P_Class_06="나1" Then 
	    MsgP_06="1) 목욕용 오일ㆍ정제ㆍ캡슐"
	    
	    elseif P_Class_06="나2" Then 
	    MsgP_06="2) 목욕용 소금류"
	    
	    elseif P_Class_06="나3" Then 
	    MsgP_06="3) 버블 배스"
	    
	    elseif P_Class_06="나4" Then 
	    MsgP_06="4) 그 밖의 목욕용 제품류"
	    
	    elseif P_Class_06="다1" Then 
	    MsgP_06="1) 폼 클렌저"
	    
	    elseif P_Class_06="다2" Then 
	    MsgP_06="2) 바디 클렌저"
	    
	    elseif P_Class_06="다3" Then 
	    MsgP_06="3) 액상비누"
	    
	    elseif P_Class_06="다4" Then 
	    MsgP_06="4) 외음부 세정제"
	    
	    elseif P_Class_06="다5" Then 
	    MsgP_06="5) 그 밖의 인체 세정용 제품류"
	    
	    elseif P_Class_06="라1" Then 
	    MsgP_06="1) 아이브로 펜슬"
	    
	    elseif P_Class_06="라2" Then 
	    MsgP_06="2) 아이 라이너"
	    
	    elseif P_Class_06="라3" Then 
	    MsgP_06="3) 아이 섀도"
	    
	    elseif P_Class_06="라4" Then 
	    MsgP_06="4) 마스카라"
	    
	    elseif P_Class_06="라5" Then 
	    MsgP_06="5) 아이 메이크업 리무버"
	    
	    elseif P_Class_06="라6" Then 
	    MsgP_06="6) 그 밖의 눈화장용 제품류"
	    
	    elseif P_Class_06="마1" Then 
	    MsgP_06="1) 향수"
	    
	    elseif P_Class_06="마2" Then 
	    MsgP_06="2) 분말향"
	    
	    elseif P_Class_06="마3" Then 
	    MsgP_06="3) 향낭"
	    
	    elseif P_Class_06="마4" Then 
	    MsgP_06="4) 코롱"
	    
	    elseif P_Class_06="마5" Then 
	    MsgP_06="5) 그 밖의 방향용 제품류"
	    
	    elseif P_Class_06="바1" Then 
	    MsgP_06="1) 헤어 틴트"
	    
	    elseif P_Class_06="바2" Then 
	    MsgP_06="2) 헤어 칼라스프레이"
	    
	    elseif P_Class_06="바3" Then 
	    MsgP_06="3) 그 밖의 염모용 제품류"
	    
	    elseif P_Class_06="사1" Then 
	    MsgP_06="1) 볼연지"
	    
	    elseif P_Class_06="사2" Then 
	    MsgP_06="2) 페이스 파우더, 페이스 케익"
	    
	    elseif P_Class_06="사3" Then 
	    MsgP_06="3) 리퀴드, 크림, 케익 파운데이션"
	    
	    elseif P_Class_06="사4" Then 
	    MsgP_06="4) 메이크업 베이스"
	    
	    elseif P_Class_06="사5" Then 
	    MsgP_06="5) 메이크업 픽서티브"
	    
	    elseif P_Class_06="사6" Then 
	    MsgP_06="6) 립스틱, 립라이너"
	    
	    elseif P_Class_06="사7" Then 
	    MsgP_06="7) 립글로스, 립밤"
	    
	    elseif P_Class_06="사8" Then 
	    MsgP_06="8) 바디페인팅, 분장용 제품"
	    
	    elseif P_Class_06="사9" Then 
	    MsgP_06="9) 그 밖의 메이크업 제품류"
	    
	    elseif P_Class_06="아1" Then 
	    MsgP_06="1) 헤어 컨디셔너"
	    
	    elseif P_Class_06="아2" Then 
	    MsgP_06="2) 헤어 토닉"
	    
	    elseif P_Class_06="아3" Then 
	    MsgP_06="3) 헤어 그루밍에이드"
	    
	    elseif P_Class_06="아4" Then 
	    MsgP_06="4) 헤어 크림, 로션"
	    
	    elseif P_Class_06="아5" Then 
	    MsgP_06="5) 헤어 오일"
	    
	    elseif P_Class_06="아6" Then 
	    MsgP_06="6) 포마드"
	    
	    elseif P_Class_06="아7" Then 
	    MsgP_06="7) 헤어 스프레이ㆍ무스ㆍ왁스ㆍ젤"
	    
	    elseif P_Class_06="아8" Then 
	    MsgP_06="8) 샴푸, 린스"
	    
	    elseif P_Class_06="아9" Then 
	    MsgP_06="9) 퍼머넌트 웨이브"
	    
	    elseif P_Class_06="아10" Then 
	    MsgP_06="10) 헤어 스트레이트너"
	    
	    elseif P_Class_06="아11" Then 
	    MsgP_06="11) 그 밖의 두발용 제품류"
	    
	    elseif P_Class_06="자1" Then 
	    MsgP_06="1) 베이스코트, 언더코트"
	    
	    elseif P_Class_06="자2" Then 
	    MsgP_06="2) 네일폴리시, 네일에나멜"
	    
	    elseif P_Class_06="자3" Then 
	    MsgP_06="3) 탑코트"
	    
	    elseif P_Class_06="자4" Then 
	    MsgP_06="4) 네일 크림ㆍ로션ㆍ에센스"
	    
	    elseif P_Class_06="자5" Then 
	    MsgP_06="5) 네일폴리시ㆍ네일에나멜 리무버"
	    
	    elseif P_Class_06="자6" Then 
	    MsgP_06="6) 그 밖의 손발톱용 제품류"
	    
	    elseif P_Class_06="차1" Then 
	    MsgP_06="1) 애프터셰이브 로션"
	    
	    elseif P_Class_06="차2" Then 
	    MsgP_06="2) 남성용 탤컴"
	    
	    elseif P_Class_06="차3" Then 
	    MsgP_06="3) 프리셰이브 로션"
	    
	    elseif P_Class_06="차4" Then 
	    MsgP_06="4) 셰이빙 크림"
	    
	    elseif P_Class_06="차5" Then 
	    MsgP_06="5) 셰이빙 폼"
	    
	    elseif P_Class_06="차6" Then 
	    MsgP_06="6) 그 밖의 면도용 제품류"
	    
	    elseif P_Class_06="카1" Then 
	    MsgP_06="1) 수렴ㆍ유연ㆍ영양화장수"
	    
	    elseif P_Class_06="카2" Then 
	    MsgP_06="2) 마사지 크림"
	    
	    elseif P_Class_06="카3" Then 
	    MsgP_06="3) 에센스, 오일"
	    
	    elseif P_Class_06="카4" Then 
	    MsgP_06="4) 파우더"
	    
	    elseif P_Class_06="카5" Then 
	    MsgP_06="5) 바디 제품"
	    
	    elseif P_Class_06="카6" Then 
	    MsgP_06="6) 팩, 마스크"
	    
	    elseif P_Class_06="카7" Then 
	    MsgP_06="7) 눈 주위 제품"
	    
	    elseif P_Class_06="카8" Then 
	    MsgP_06="8) 로션, 크림"
	    
	    elseif P_Class_06="카9" Then 
	    MsgP_06="9) 손ㆍ발의 피부연화 제품"
	    
	    elseif P_Class_06="카10" Then 
	    MsgP_06="10) 클렌징워터ㆍ클렌징오일ㆍ클렌징로션ㆍ클렌징크림 등 메이크업 리무버"
	    
	    elseif P_Class_06="카11" Then 
	    MsgP_06="11) 그 밖의 기초화장용 제품류"
	    
	    elseif P_Class_06="타1" Then 
	    MsgP_06="1) 데오도런트"
	    
	    elseif P_Class_06="타2" Then 
	    MsgP_06="2) 그 밖의 체취 방지용 제품류"
	    
	    else                  
	   	  MsgP_06=""
      end If
	 
	 
	 
  IF      Functional_06="F1" Then 
          MsgF_06="미백"
  elseif  Functional_06="F2" Then
	   	    MsgF_06="주름개선"
  elseif  Functional_06="F3" Then
	   	    MsgF_06="자외선 차단"
	elseif  Functional_06="F4" Then
	   	    MsgF_06="미백, 주름개선"
	elseif  Functional_06="F5" Then
	   	    MsgF_06="미백, 자외선 차단"
  elseif  Functional_06="F6" Then
	   	    MsgF_06="주름개선, 자외선 차단"
  elseif  Functional_06="F7" Then
	   	    MsgF_06="미백, 주름개선, 자외선 차단"
  elseif  Functional_06="F8" Then
	   	    MsgF_06="염모"
  elseif  Functional_06="F9" Then
	   	    MsgF_06="제모"
  elseif  Functional_06="F10" Then
	   	    MsgF_06="탈모 완화"
  elseif  Functional_06="F11" Then
	   	    MsgF_06="여드름성 피부 완화"
  elseif  Functional_06="F12" Then
	   	    MsgF_06="아토피성 피부 보습"
  elseif  Functional_06="F13" Then
	   	    MsgF_06="튼살로 인한 붉은 선 완화"
  elseif  Functional_06="F14" Then
	   	    MsgF_06="기타 복합유형"
  elseif  Functional_06="일반" Then
	   	    MsgF_06="일반"
  else                  
	   	  MsgF_06=""
     end If	 
     

     
     
     
     
     
     
     
     Cosmetic_Class_Big_07 = LEFT(P_Class_07,1)

'유형
  IF Cosmetic_Class_Big_07="가" Then            
   		  Msg_07="영ㆍ유아용(만 3세이하의 어린이용을 말한다.) 제품류"
     elseif Cosmetic_Class_Big_07="나" Then     
         Msg_07="목욕용 제품류"
     elseif  Cosmetic_Class_Big_07="다" Then  
	   	   Msg_07="인체 세정용 제품류(화장비누 제외)"
	  elseif  Cosmetic_Class_Big_07="라" Then
	   	   Msg_07="눈 화장용 제품류"
    elseif  Cosmetic_Class_Big_07="마" Then
	   	   Msg_07="방향용 제품류"
	  elseif  Cosmetic_Class_Big_07="바" Then
	   	   Msg_07="두발 염색용 제품류"   
	  elseif  Cosmetic_Class_Big_07="사" Then
	   	   Msg_07="색조 화장용 제품"
	  elseif  Cosmetic_Class_Big_07="아" Then
	   	   Msg_07="두발용 제품류"   
	  elseif  Cosmetic_Class_Big_07="자" Then
	   	   Msg_07="손발톱용 제품류"
	  elseif  Cosmetic_Class_Big_07="차" Then
	   	   Msg_07="면도용 제품류"   
	  elseif  Cosmetic_Class_Big_07="카" Then
	   	   Msg_07="기초화장용 제품류"
	  elseif  Cosmetic_Class_Big_07="타" Then
	   	   Msg_07="체취 방지용 제품류" 
	 elseif  Cosmetic_Class_Big_07="파" Then
	   	   Msg_07="체모 제거용 제품류"
	   	   
	 else
	   	  Msg_07=""
     end If	   	   
	
	
	IF 	P_Class_07="가1" Then 
      MsgP_07="1) 영ㆍ유아용 샴푸, 린스"
	  
	    elseif P_Class_07="가2" Then 
	    MsgP_07="2) 영ㆍ유아용 로션, 크림"
	    
	    elseif P_Class_07="가3" Then 
	    MsgP_07="3) 영ㆍ유아용 오일"
	    
	    elseif P_Class_07="가4" Then 
	    MsgP_07="4) 영ㆍ유아용 인체 세정용 제품"
	    
	    elseif P_Class_07="가5" Then 
	    MsgP_07="5) 영ㆍ유아용 목욕용 제품"
	    
	    elseif P_Class_07="나1" Then 
	    MsgP_07="1) 목욕용 오일ㆍ정제ㆍ캡슐"
	    
	    elseif P_Class_07="나2" Then 
	    MsgP_07="2) 목욕용 소금류"
	    
	    elseif P_Class_07="나3" Then 
	    MsgP_07="3) 버블 배스"
	    
	    elseif P_Class_07="나4" Then 
	    MsgP_07="4) 그 밖의 목욕용 제품류"
	    
	    elseif P_Class_07="다1" Then 
	    MsgP_07="1) 폼 클렌저"
	    
	    elseif P_Class_07="다2" Then 
	    MsgP_07="2) 바디 클렌저"
	    
	    elseif P_Class_07="다3" Then 
	    MsgP_07="3) 액상비누"
	    
	    elseif P_Class_07="다4" Then 
	    MsgP_07="4) 외음부 세정제"
	    
	    elseif P_Class_07="다5" Then 
	    MsgP_07="5) 그 밖의 인체 세정용 제품류"
	    
	    elseif P_Class_07="라1" Then 
	    MsgP_07="1) 아이브로 펜슬"
	    
	    elseif P_Class_07="라2" Then 
	    MsgP_07="2) 아이 라이너"
	    
	    elseif P_Class_07="라3" Then 
	    MsgP_07="3) 아이 섀도"
	    
	    elseif P_Class_07="라4" Then 
	    MsgP_07="4) 마스카라"
	    
	    elseif P_Class_07="라5" Then 
	    MsgP_07="5) 아이 메이크업 리무버"
	    
	    elseif P_Class_07="라6" Then 
	    MsgP_07="6) 그 밖의 눈화장용 제품류"
	    
	    elseif P_Class_07="마1" Then 
	    MsgP_07="1) 향수"
	    
	    elseif P_Class_07="마2" Then 
	    MsgP_07="2) 분말향"
	    
	    elseif P_Class_07="마3" Then 
	    MsgP_07="3) 향낭"
	    
	    elseif P_Class_07="마4" Then 
	    MsgP_07="4) 코롱"
	    
	    elseif P_Class_07="마5" Then 
	    MsgP_07="5) 그 밖의 방향용 제품류"
	    
	    elseif P_Class_07="바1" Then 
	    MsgP_07="1) 헤어 틴트"
	    
	    elseif P_Class_07="바2" Then 
	    MsgP_07="2) 헤어 칼라스프레이"
	    
	    elseif P_Class_07="바3" Then 
	    MsgP_07="3) 그 밖의 염모용 제품류"
	    
	    elseif P_Class_07="사1" Then 
	    MsgP_07="1) 볼연지"
	    
	    elseif P_Class_07="사2" Then 
	    MsgP_07="2) 페이스 파우더, 페이스 케익"
	    
	    elseif P_Class_07="사3" Then 
	    MsgP_07="3) 리퀴드, 크림, 케익 파운데이션"
	    
	    elseif P_Class_07="사4" Then 
	    MsgP_07="4) 메이크업 베이스"
	    
	    elseif P_Class_07="사5" Then 
	    MsgP_07="5) 메이크업 픽서티브"
	    
	    elseif P_Class_07="사6" Then 
	    MsgP_07="6) 립스틱, 립라이너"
	    
	    elseif P_Class_07="사7" Then 
	    MsgP_07="7) 립글로스, 립밤"
	    
	    elseif P_Class_07="사8" Then 
	    MsgP_07="8) 바디페인팅, 분장용 제품"
	    
	    elseif P_Class_07="사9" Then 
	    MsgP_07="9) 그 밖의 메이크업 제품류"
	    
	    elseif P_Class_07="아1" Then 
	    MsgP_07="1) 헤어 컨디셔너"
	    
	    elseif P_Class_07="아2" Then 
	    MsgP_07="2) 헤어 토닉"
	    
	    elseif P_Class_07="아3" Then 
	    MsgP_07="3) 헤어 그루밍에이드"
	    
	    elseif P_Class_07="아4" Then 
	    MsgP_07="4) 헤어 크림, 로션"
	    
	    elseif P_Class_07="아5" Then 
	    MsgP_07="5) 헤어 오일"
	    
	    elseif P_Class_07="아6" Then 
	    MsgP_07="6) 포마드"
	    
	    elseif P_Class_07="아7" Then 
	    MsgP_07="7) 헤어 스프레이ㆍ무스ㆍ왁스ㆍ젤"
	    
	    elseif P_Class_07="아8" Then 
	    MsgP_07="8) 샴푸, 린스"
	    
	    elseif P_Class_07="아9" Then 
	    MsgP_07="9) 퍼머넌트 웨이브"
	    
	    elseif P_Class_07="아10" Then 
	    MsgP_07="10) 헤어 스트레이트너"
	    
	    elseif P_Class_07="아11" Then 
	    MsgP_07="11) 그 밖의 두발용 제품류"
	    
	    elseif P_Class_07="자1" Then 
	    MsgP_07="1) 베이스코트, 언더코트"
	    
	    elseif P_Class_07="자2" Then 
	    MsgP_07="2) 네일폴리시, 네일에나멜"
	    
	    elseif P_Class_07="자3" Then 
	    MsgP_07="3) 탑코트"
	    
	    elseif P_Class_07="자4" Then 
	    MsgP_07="4) 네일 크림ㆍ로션ㆍ에센스"
	    
	    elseif P_Class_07="자5" Then 
	    MsgP_07="5) 네일폴리시ㆍ네일에나멜 리무버"
	    
	    elseif P_Class_07="자6" Then 
	    MsgP_07="6) 그 밖의 손발톱용 제품류"
	    
	    elseif P_Class_07="차1" Then 
	    MsgP_07="1) 애프터셰이브 로션"
	    
	    elseif P_Class_07="차2" Then 
	    MsgP_07="2) 남성용 탤컴"
	    
	    elseif P_Class_07="차3" Then 
	    MsgP_07="3) 프리셰이브 로션"
	    
	    elseif P_Class_07="차4" Then 
	    MsgP_07="4) 셰이빙 크림"
	    
	    elseif P_Class_07="차5" Then 
	    MsgP_07="5) 셰이빙 폼"
	    
	    elseif P_Class_07="차6" Then 
	    MsgP_07="6) 그 밖의 면도용 제품류"
	    
	    elseif P_Class_07="카1" Then 
	    MsgP_07="1) 수렴ㆍ유연ㆍ영양화장수"
	    
	    elseif P_Class_07="카2" Then 
	    MsgP_07="2) 마사지 크림"
	    
	    elseif P_Class_07="카3" Then 
	    MsgP_07="3) 에센스, 오일"
	    
	    elseif P_Class_07="카4" Then 
	    MsgP_07="4) 파우더"
	    
	    elseif P_Class_07="카5" Then 
	    MsgP_07="5) 바디 제품"
	    
	    elseif P_Class_07="카6" Then 
	    MsgP_07="6) 팩, 마스크"
	    
	    elseif P_Class_07="카7" Then 
	    MsgP_07="7) 눈 주위 제품"
	    
	    elseif P_Class_07="카8" Then 
	    MsgP_07="8) 로션, 크림"
	    
	    elseif P_Class_07="카9" Then 
	    MsgP_07="9) 손ㆍ발의 피부연화 제품"
	    
	    elseif P_Class_07="카10" Then 
	    MsgP_07="10) 클렌징워터ㆍ클렌징오일ㆍ클렌징로션ㆍ클렌징크림 등 메이크업 리무버"
	    
	    elseif P_Class_07="카11" Then 
	    MsgP_07="11) 그 밖의 기초화장용 제품류"
	    
	    elseif P_Class_07="타1" Then 
	    MsgP_07="1) 데오도런트"
	    
	    elseif P_Class_07="타2" Then 
	    MsgP_07="2) 그 밖의 체취 방지용 제품류"
	    
	    else                  
	   	  MsgP_07=""
      end If
	 
	 
	 
  IF      Functional_07="F1" Then 
          MsgF_07="미백"
  elseif  Functional_07="F2" Then
	   	    MsgF_07="주름개선"
  elseif  Functional_07="F3" Then
	   	    MsgF_07="자외선 차단"
	elseif  Functional_07="F4" Then
	   	    MsgF_07="미백, 주름개선"
	elseif  Functional_07="F5" Then
	   	    MsgF_07="미백, 자외선 차단"
  elseif  Functional_07="F6" Then
	   	    MsgF_07="주름개선, 자외선 차단"
  elseif  Functional_07="F7" Then
	   	    MsgF_07="미백, 주름개선, 자외선 차단"
  elseif  Functional_07="F8" Then
	   	    MsgF_07="염모"
  elseif  Functional_07="F9" Then
	   	    MsgF_07="제모"
  elseif  Functional_07="F10" Then
	   	    MsgF_07="탈모 완화"
  elseif  Functional_07="F11" Then
	   	    MsgF_07="여드름성 피부 완화"
  elseif  Functional_07="F12" Then
	   	    MsgF_07="아토피성 피부 보습"
  elseif  Functional_07="F13" Then
	   	    MsgF_07="튼살로 인한 붉은 선 완화"
  elseif  Functional_07="F14" Then
	   	    MsgF_07="기타 복합유형"
  elseif  Functional_07="일반" Then
	   	    MsgF_07="일반"
  else                  
	   	  MsgF_07=""
     end If	 
     
     
     
     
     
     
     
     
     
     Cosmetic_Class_Big_08 = LEFT(P_Class_08,1)

'유형
  IF Cosmetic_Class_Big_08="가" Then            
   		  Msg_08="영ㆍ유아용(만 3세이하의 어린이용을 말한다.) 제품류"
     elseif Cosmetic_Class_Big_08="나" Then     
         Msg_08="목욕용 제품류"
     elseif  Cosmetic_Class_Big_08="다" Then  
	   	   Msg_08="인체 세정용 제품류(화장비누 제외)"
	  elseif  Cosmetic_Class_Big_08="라" Then
	   	   Msg_08="눈 화장용 제품류"
    elseif  Cosmetic_Class_Big_08="마" Then
	   	   Msg_08="방향용 제품류"
	  elseif  Cosmetic_Class_Big_08="바" Then
	   	   Msg_08="두발 염색용 제품류"   
	  elseif  Cosmetic_Class_Big_08="사" Then
	   	   Msg_08="색조 화장용 제품"
	  elseif  Cosmetic_Class_Big_08="아" Then
	   	   Msg_08="두발용 제품류"   
	  elseif  Cosmetic_Class_Big_08="자" Then
	   	   Msg_08="손발톱용 제품류"
	  elseif  Cosmetic_Class_Big_08="차" Then
	   	   Msg_08="면도용 제품류"   
	  elseif  Cosmetic_Class_Big_08="카" Then
	   	   Msg_08="기초화장용 제품류"
	  elseif  Cosmetic_Class_Big_08="타" Then
	   	   Msg_08="체취 방지용 제품류" 
	 elseif  Cosmetic_Class_Big_08="파" Then
	   	   Msg_08="체모 제거용 제품류"
	   	   
	 else
	   	  Msg_08=""
     end If	   	   
	
	
	IF 	P_Class_08="가1" Then 
      MsgP_08="1) 영ㆍ유아용 샴푸, 린스"
	  
	    elseif P_Class_08="가2" Then 
	    MsgP_08="2) 영ㆍ유아용 로션, 크림"
	    
	    elseif P_Class_08="가3" Then 
	    MsgP_08="3) 영ㆍ유아용 오일"
	    
	    elseif P_Class_08="가4" Then 
	    MsgP_08="4) 영ㆍ유아용 인체 세정용 제품"
	    
	    elseif P_Class_08="가5" Then 
	    MsgP_08="5) 영ㆍ유아용 목욕용 제품"
	    
	    elseif P_Class_08="나1" Then 
	    MsgP_08="1) 목욕용 오일ㆍ정제ㆍ캡슐"
	    
	    elseif P_Class_08="나2" Then 
	    MsgP_08="2) 목욕용 소금류"
	    
	    elseif P_Class_08="나3" Then 
	    MsgP_08="3) 버블 배스"
	    
	    elseif P_Class_08="나4" Then 
	    MsgP_08="4) 그 밖의 목욕용 제품류"
	    
	    elseif P_Class_08="다1" Then 
	    MsgP_08="1) 폼 클렌저"
	    
	    elseif P_Class_08="다2" Then 
	    MsgP_08="2) 바디 클렌저"
	    
	    elseif P_Class_08="다3" Then 
	    MsgP_08="3) 액상비누"
	    
	    elseif P_Class_08="다4" Then 
	    MsgP_08="4) 외음부 세정제"
	    
	    elseif P_Class_08="다5" Then 
	    MsgP_08="5) 그 밖의 인체 세정용 제품류"
	    
	    elseif P_Class_08="라1" Then 
	    MsgP_08="1) 아이브로 펜슬"
	    
	    elseif P_Class_08="라2" Then 
	    MsgP_08="2) 아이 라이너"
	    
	    elseif P_Class_08="라3" Then 
	    MsgP_08="3) 아이 섀도"
	    
	    elseif P_Class_08="라4" Then 
	    MsgP_08="4) 마스카라"
	    
	    elseif P_Class_08="라5" Then 
	    MsgP_08="5) 아이 메이크업 리무버"
	    
	    elseif P_Class_08="라6" Then 
	    MsgP_08="6) 그 밖의 눈화장용 제품류"
	    
	    elseif P_Class_08="마1" Then 
	    MsgP_08="1) 향수"
	    
	    elseif P_Class_08="마2" Then 
	    MsgP_08="2) 분말향"
	    
	    elseif P_Class_08="마3" Then 
	    MsgP_08="3) 향낭"
	    
	    elseif P_Class_08="마4" Then 
	    MsgP_08="4) 코롱"
	    
	    elseif P_Class_08="마5" Then 
	    MsgP_08="5) 그 밖의 방향용 제품류"
	    
	    elseif P_Class_08="바1" Then 
	    MsgP_08="1) 헤어 틴트"
	    
	    elseif P_Class_08="바2" Then 
	    MsgP_08="2) 헤어 칼라스프레이"
	    
	    elseif P_Class_08="바3" Then 
	    MsgP_08="3) 그 밖의 염모용 제품류"
	    
	    elseif P_Class_08="사1" Then 
	    MsgP_08="1) 볼연지"
	    
	    elseif P_Class_08="사2" Then 
	    MsgP_08="2) 페이스 파우더, 페이스 케익"
	    
	    elseif P_Class_08="사3" Then 
	    MsgP_08="3) 리퀴드, 크림, 케익 파운데이션"
	    
	    elseif P_Class_08="사4" Then 
	    MsgP_08="4) 메이크업 베이스"
	    
	    elseif P_Class_08="사5" Then 
	    MsgP_08="5) 메이크업 픽서티브"
	    
	    elseif P_Class_08="사6" Then 
	    MsgP_08="6) 립스틱, 립라이너"
	    
	    elseif P_Class_08="사7" Then 
	    MsgP_08="7) 립글로스, 립밤"
	    
	    elseif P_Class_08="사8" Then 
	    MsgP_08="8) 바디페인팅, 분장용 제품"
	    
	    elseif P_Class_08="사9" Then 
	    MsgP_08="9) 그 밖의 메이크업 제품류"
	    
	    elseif P_Class_08="아1" Then 
	    MsgP_08="1) 헤어 컨디셔너"
	    
	    elseif P_Class_08="아2" Then 
	    MsgP_08="2) 헤어 토닉"
	    
	    elseif P_Class_08="아3" Then 
	    MsgP_08="3) 헤어 그루밍에이드"
	    
	    elseif P_Class_08="아4" Then 
	    MsgP_08="4) 헤어 크림, 로션"
	    
	    elseif P_Class_08="아5" Then 
	    MsgP_08="5) 헤어 오일"
	    
	    elseif P_Class_08="아6" Then 
	    MsgP_08="6) 포마드"
	    
	    elseif P_Class_08="아7" Then 
	    MsgP_08="7) 헤어 스프레이ㆍ무스ㆍ왁스ㆍ젤"
	    
	    elseif P_Class_08="아8" Then 
	    MsgP_08="8) 샴푸, 린스"
	    
	    elseif P_Class_08="아9" Then 
	    MsgP_08="9) 퍼머넌트 웨이브"
	    
	    elseif P_Class_08="아10" Then 
	    MsgP_08="10) 헤어 스트레이트너"
	    
	    elseif P_Class_08="아11" Then 
	    MsgP_08="11) 그 밖의 두발용 제품류"
	    
	    elseif P_Class_08="자1" Then 
	    MsgP_08="1) 베이스코트, 언더코트"
	    
	    elseif P_Class_08="자2" Then 
	    MsgP_08="2) 네일폴리시, 네일에나멜"
	    
	    elseif P_Class_08="자3" Then 
	    MsgP_08="3) 탑코트"
	    
	    elseif P_Class_08="자4" Then 
	    MsgP_08="4) 네일 크림ㆍ로션ㆍ에센스"
	    
	    elseif P_Class_08="자5" Then 
	    MsgP_08="5) 네일폴리시ㆍ네일에나멜 리무버"
	    
	    elseif P_Class_08="자6" Then 
	    MsgP_08="6) 그 밖의 손발톱용 제품류"
	    
	    elseif P_Class_08="차1" Then 
	    MsgP_08="1) 애프터셰이브 로션"
	    
	    elseif P_Class_08="차2" Then 
	    MsgP_08="2) 남성용 탤컴"
	    
	    elseif P_Class_08="차3" Then 
	    MsgP_08="3) 프리셰이브 로션"
	    
	    elseif P_Class_08="차4" Then 
	    MsgP_08="4) 셰이빙 크림"
	    
	    elseif P_Class_08="차5" Then 
	    MsgP_08="5) 셰이빙 폼"
	    
	    elseif P_Class_08="차6" Then 
	    MsgP_08="6) 그 밖의 면도용 제품류"
	    
	    elseif P_Class_08="카1" Then 
	    MsgP_08="1) 수렴ㆍ유연ㆍ영양화장수"
	    
	    elseif P_Class_08="카2" Then 
	    MsgP_08="2) 마사지 크림"
	    
	    elseif P_Class_08="카3" Then 
	    MsgP_08="3) 에센스, 오일"
	    
	    elseif P_Class_08="카4" Then 
	    MsgP_08="4) 파우더"
	    
	    elseif P_Class_08="카5" Then 
	    MsgP_08="5) 바디 제품"
	    
	    elseif P_Class_08="카6" Then 
	    MsgP_08="6) 팩, 마스크"
	    
	    elseif P_Class_08="카7" Then 
	    MsgP_08="7) 눈 주위 제품"
	    
	    elseif P_Class_08="카8" Then 
	    MsgP_08="8) 로션, 크림"
	    
	    elseif P_Class_08="카9" Then 
	    MsgP_08="9) 손ㆍ발의 피부연화 제품"
	    
	    elseif P_Class_08="카10" Then 
	    MsgP_08="10) 클렌징워터ㆍ클렌징오일ㆍ클렌징로션ㆍ클렌징크림 등 메이크업 리무버"
	    
	    elseif P_Class_08="카11" Then 
	    MsgP_08="11) 그 밖의 기초화장용 제품류"
	    
	    elseif P_Class_08="타1" Then 
	    MsgP_08="1) 데오도런트"
	    
	    elseif P_Class_08="타2" Then 
	    MsgP_08="2) 그 밖의 체취 방지용 제품류"
	    
	    else                  
	   	  MsgP_08=""
      end If
	 
	 
	 
  IF      Functional_08="F1" Then 
          MsgF_08="미백"
  elseif  Functional_08="F2" Then
	   	    MsgF_08="주름개선"
  elseif  Functional_08="F3" Then
	   	    MsgF_08="자외선 차단"
	elseif  Functional_08="F4" Then
	   	    MsgF_08="미백, 주름개선"
	elseif  Functional_08="F5" Then
	   	    MsgF_08="미백, 자외선 차단"
  elseif  Functional_08="F6" Then
	   	    MsgF_08="주름개선, 자외선 차단"
  elseif  Functional_08="F7" Then
	   	    MsgF_08="미백, 주름개선, 자외선 차단"
  elseif  Functional_08="F8" Then
	   	    MsgF_08="염모"
  elseif  Functional_08="F9" Then
	   	    MsgF_08="제모"
  elseif  Functional_08="F10" Then
	   	    MsgF_08="탈모 완화"
  elseif  Functional_08="F11" Then
	   	    MsgF_08="여드름성 피부 완화"
  elseif  Functional_08="F12" Then
	   	    MsgF_08="아토피성 피부 보습"
  elseif  Functional_08="F13" Then
	   	    MsgF_08="튼살로 인한 붉은 선 완화"
  elseif  Functional_08="F14" Then
	   	    MsgF_08="기타 복합유형"
  elseif  Functional_08="일반" Then
	   	    MsgF_08="일반"
  else                  
	   	  MsgF_08=""
     end If	   

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

rs.close
Set rs=nothing

%>
<html>
<head>
<title>완제품 Db(기타)</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="basic_Finsh_Goods_Other.css" type="text/css">
</head>
<body bgcolor="#D7F1FA">
<center>
<script language="Javascript">
<!--
function Send() {

	var vA = document.form.Delivery_Amount.value;
	var vB = document.form.Lot_number_01.value;
	var vC = document.form.Expiration_Date_01.value;
	var vD = document.form.Good_class.value;
	var vE = document.form.Judge_Result.value;
	var vF = document.form.Supplier.value;
	var vG = document.form.Warehouse.value;
	var vH = document.form.Manage_No.value;
	
	var vP = document.form.Pass.value;
	

	
	
	if (vA == "") {
		alert("납품수량을 입력하세요.\n");
		document.form.Delivery_Amount.focus();
		return false;
		}
	if (isNaN(form.Delivery_Amount.value))
   { 
   alert("납품량 값이 유효한 값(숫자)이 아닙니다!\n\n\숫자라도 콤마가 있으면 안됩니다"); 
   document.form.Delivery_Amount.focus()
  return false;
   }
	
	if (vB == "") {
		alert("로트번호1은 필수 입니다.\n");
		document.form.Lot_number_01.focus();
		return false;
		}
		
	if (vC == "") {
		alert("사용기한 1은 필수 입니다.\n");
		document.form.Expiration_Date_01.focus();
		return false;
		}
		
				
	if (vD == "") {
		alert("제품종류를 선택하세요.\n");
		document.form.Good_class.focus();
		return false;
		}
		
		
	
	if (vE == "") {
		alert("판정결과를 선택하세요.\n");
		document.form.Judge_Result.focus();
		return false;
		}
		
		
		
  if (vF == "") {
		alert("제조사를 입력하세요.\n");
		document.form.Supplier.focus();
		return false;
		}

	if (vG == "") {
		alert("입고처를 입력하세요.\n");
		document.form.Warehouse.focus();
		return false;
		}
		
		if (vG == "") {
		alert("관리번호를 입력하세요.\n");
		document.form.Manage_No.focus();
		return false;
		}
		
		
	if (vP == "") {
		alert("암호를 입력하세요.\n");
		document.form.Pass.focus();
		return false;
		}

return true;
} // end function
//  -->
</script>

<%
to_year = year(date())

start_year = Cint(to_year)-1
end_year = Cint(to_year)+1

sdate=date()
tag="1"
%>
 <form name="form" method="post" action="write_ok.asp" onSubmit="return Send()" enctype="multipart/form-data">
    
     <form name="form" method="post" action="write_ok.asp" onSubmit="return Send()" enctype="multipart/form-data">
  <table cellspacing=0 cellpadding=0 border="0"  width="1024" cellpadding="0" cellspacing="1" style="word-wrap:break-word;word-break:break-all;">
    <tr> 
      <td width="512"  bgcolor="#D7F1FA" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <b>▶ 선택 완제품(기타) 세부 정보 </b></td>
      <td width="512"  bgcolor="#D7F1FA" style="text-align:right; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
    <a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0"></a></td>
    </tr>
</table>
  <table cellspacing=0 cellpadding=0 border="1"  width="1024" cellpadding="0" cellspacing="1" style="word-wrap:break-word;word-break:break-all;">
      <tr>
         <th width="120">
        완제품 코드</span></th>
          <td width="221" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=Product_Code%></td>
       <th width="120">
       등 록 자</span></td>
           <td width="221" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=Registor%></td>
        <th width="120">
        상품 구성</td>
       <td width="222" style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <% if Product_Name_KFDA_02 <> ""  then %>
     <font color=red>복합</font>
     
     <% else%>
     <font color=blue>단일</font>
     <% end if %></td>
       </tr>
         <tr> 
       <th>
      완제품명(더존)</td>
        <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
       <%=Product_Name_DZ%>&nbsp;</td>
       </tr>
       <tr>  
        <th>
       완제품명(최종)</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_Final%>&nbsp;</td>
      </tr>
      
          <tr>  
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크코드 1</td>
       <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Bulk_Code_01%></td>
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크제품명 1</td>
       <td colspan=3  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_KFDA_01%></td>
     </tr>
      
      <tr> 
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>유&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;형 1</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Class_01%>&nbsp;&nbsp;&nbsp;<%=Msg_01%>&nbsp;&nbsp;&nbsp;<%=MsgP_01%></td>
      
      </tr>
      
      <tr>  
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;량 1</td>
      <td style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Capacity_01%>&nbsp;<%=P_Capacity_Unit_01%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>사용기한 1</td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Period_of_Usage_01%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>제조업자 1</td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <%=Manufacturer_01%></td>
      </tr>
       <% if Product_Name_KFDA_02<>"" then %>
       
        <tr>
        <td colspan=6"><br style="line-height:1pt;"> </td>
        </tr>
              <tr>  
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크코드 2</td>
       <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Bulk_Code_02%></td>
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크제품명 2</td>
       <td colspan=3  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_KFDA_02%></td>
     </tr>
      
      <tr> 
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>유&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;형 2</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Class_02%>&nbsp;&nbsp;&nbsp;<%=Msg_02%>&nbsp;&nbsp;&nbsp;<%=MsgP_02%></td>
     
      </tr>
      
      <tr>  
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;량 2</td>
      <td style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Capacity_02%>&nbsp;<%=P_Capacity_Unit_02%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>사용기한 2</td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Period_of_Usage_02%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>제조업자 2</td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <%=Manufacturer_02%></td>
      </tr>
       <% else %>
      <% end if %>
      
       <% if Product_Name_KFDA_03<>"" then %>
       <tr>
        <td colspan=6"><br style="line-height:1pt;"> </td>
        </tr>
             <tr>  
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크코드 3</td>
       <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Bulk_Code_03%></td>
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크제품명 3</td>
       <td colspan=3  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_KFDA_03%></td>
     </tr>
      
      <tr> 
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>유&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;형 3</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Class_03%>&nbsp;&nbsp;&nbsp;<%=Msg_03%>&nbsp;&nbsp;&nbsp;<%=MsgP_03%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      
      </tr>
 
      
      <tr>  
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;량 3</td>
      <td style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Capacity_03%>&nbsp;<%=P_Capacity_Unit_03%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>사용기한 3</td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Period_of_Usage_03%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>제조업자 3</td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <%=Manufacturer_03%></td>
      </tr>
       <% else %>
      <% end if %>
      
       <% if Product_Name_KFDA_04<>"" then %>
       <tr>
        <td colspan=6"><br style="line-height:1pt;"> </td>
        </tr>
           <tr>  
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크코드 4</td>
       <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Bulk_Code_04%></td>
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크제품명 4</td>
       <td colspan=3  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_KFDA_04%></td>
     </tr>
      
      <tr> 
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>유&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;형 4</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Class_04%>&nbsp;&nbsp;&nbsp;<%=Msg_04%>&nbsp;&nbsp;&nbsp;<%=MsgP_04%></td>
     
      </tr>
      
      <tr>  
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;량 4</td>
      <td style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Capacity_04%>&nbsp;<%=P_Capacity_Unit_04%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>사용기한 4</td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Period_of_Usage_04%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>제조업자 4</td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <%=Manufacturer_04%></td>
      </tr>
       <% else %>
      <% end if %>
      
         <% if Product_Name_KFDA_05<>"" then %>
       <tr>
        <td colspan=6"><br style="line-height:1pt;"> </td>
        </tr>
             <tr>  
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크코드 5</td>
       <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Bulk_Code_05%></td>
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크제품명 5</td>
       <td colspan=3  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_KFDA_05%></td>
     </tr>
      
      <tr> 
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>유&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;형 5</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Class_05%>&nbsp;&nbsp;&nbsp;<%=Msg_05%>&nbsp;&nbsp;&nbsp;<%=MsgP_05%></td>
    
      </tr>
      
      <tr>  
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;량 5</td>
      <td style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Capacity_05%>&nbsp;<%=P_Capacity_Unit_05%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>사용기한 5</td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Period_of_Usage_05%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>제조업자 5</td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <%=Manufacturer_05%></td>
      </tr>
       <% else %>
      <% end if %>
      
       <% if Product_Name_KFDA_06<>"" then %>
       <tr>
        <td colspan=6"><br style="line-height:1pt;"> </td>
        </tr>
             <tr>  
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크코드 6</td>
       <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Bulk_Code_06%></td>
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크제품명 6</td>
       <td colspan=3  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_KFDA_06%></td>
     </tr>
      
      <tr> 
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>유&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;형 6</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Class_06%>&nbsp;&nbsp;&nbsp;<%=Msg_06%>&nbsp;&nbsp;&nbsp;<%=MsgP_06%></td>
      
      </tr>
      
      <tr>  
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;량 6</td>
      <td style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Capacity_06%>&nbsp;<%=P_Capacity_Unit_06%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>사용기한 6</td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Period_of_Usage_06%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>제조업자 6</td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <%=Manufacturer_06%></td>
      </tr>
       <% else %>
      <% end if %>
        <% if Product_Name_KFDA_07<>"" then %>
       <tr>
        <td colspan=6"><br style="line-height:1pt;"> </td>
        </tr>
            <tr>  
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크코드 7</td>
       <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Bulk_Code_07%></td>
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크제품명 7</td>
       <td colspan=3  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_KFDA_07%></td>
     </tr>
      
      <tr> 
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>유&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;형 7</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Class_07%>&nbsp;&nbsp;&nbsp;<%=Msg_07%>&nbsp;&nbsp;&nbsp;<%=MsgP_07%></td>
     
      </tr>
      
      <tr>  
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;량 7</td>
      <td style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Capacity_07%>&nbsp;<%=P_Capacity_Unit_07%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>사용기한 7</td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Period_of_Usage_07%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>제조업자 7</td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <%=Manufacturer_07%></td>
      </tr>
       <% else %>
      <% end if %>
         <% if Product_Name_KFDA_08<>"" then %>
       <tr>
        <td colspan=6"><br style="line-height:1pt;"> </td>
        </tr>
             <tr>  
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크코드 8</td>
       <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Bulk_Code_08%></td>
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>벌크제품명 8</td>
       <td colspan=3  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Product_Name_KFDA_08%></td>
     </tr>
      
      <tr> 
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>유&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;형 8</td>
       <td colspan=5  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Class_08%>&nbsp;&nbsp;&nbsp;<%=Msg_08%>&nbsp;&nbsp;&nbsp;<%=MsgP_08%></td>
     
      </tr>
      
      <tr>  
       <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>용&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;량 8</td>
      <td style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
        <%=P_Capacity_08%>&nbsp;<%=P_Capacity_Unit_08%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>사용기한 8</td>
       <td  style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
         <%=Period_of_Usage_08%></td>
      <td bgcolor="#008000" style="text-align:center; text-indent:0; margin:0; padding-top:8px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <font color=white><b>제조업자 8</td>
      <td   style="text-align:left; text-indent:0; margin:0; padding-top:6px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
      <%=Manufacturer_08%></td>
      </tr>
       <% else %>
      <% end if %>
      
       <tr>
        <th>
        완제품 비고</td>
          <td  colspan=5 style="text-align:left; text-indent:0; margin:0; padding-top:5px; padding-right:0px; padding-bottom:5px; padding-left:5px;">
          <%=Remarks_Finsh_Product%><br>&nbsp;</td>       
   </tr> 
    </table>
    <input type=hidden name=Product_Code value="<%=Product_Code%>">
<input type=hidden name=Registor value="<%=Registor%>">
<input type=hidden name=Product_Name_DZ value="<%=Product_Name_DZ%>">
<input type=hidden name=Product_Name_Final value="<%=Product_Name_Final%>">


<input type=hidden name=Bulk_Code_01 value="<%=Bulk_Code_01%>">
<input type=hidden name=Product_Name_KFDA_01 value="<%=Product_Name_KFDA_01%>">
<input type=hidden name=P_Class_01 value="<%=P_Class_01%>">
<input type=hidden name=P_Capacity_01 value="<%=P_Capacity_01%>">
<input type=hidden name=P_Capacity_Unit_01 value="<%=P_Capacity_Unit_01%>">
<input type=hidden name=Period_of_Usage_01 value="<%=Period_of_Usage_01%>">
<input type=hidden name=Manufacturer_01 value="<%=Manufacturer_01%>">
<input type=hidden name=Functional_01 value="<%=Functional_01%>">

<input type=hidden name=Bulk_Code_02 value="<%=Bulk_Code_02%>">
<input type=hidden name=Product_Name_KFDA_02 value="<%=Product_Name_KFDA_02%>">
<input type=hidden name=P_Class_02 value="<%=P_Class_02%>">
<input type=hidden name=P_Capacity_02 value="<%=P_Capacity_02%>">
<input type=hidden name=P_Capacity_Unit_02 value="<%=P_Capacity_Unit_02%>">
<input type=hidden name=Period_of_Usage_02 value="<%=Period_of_Usage_02%>">
<input type=hidden name=Manufacturer_02 value="<%=Manufacturer_02%>">
<input type=hidden name=Functional_02 value="<%=Functional_02%>">

<input type=hidden name=Bulk_Code_03 value="<%=Bulk_Code_03%>">
<input type=hidden name=Product_Name_KFDA_03 value="<%=Product_Name_KFDA_03%>">
<input type=hidden name=P_Class_03 value="<%=P_Class_03%>">
<input type=hidden name=P_Capacity_03 value="<%=P_Capacity_03%>">
<input type=hidden name=P_Capacity_Unit_03 value="<%=P_Capacity_Unit_03%>">
<input type=hidden name=Period_of_Usage_03 value="<%=Period_of_Usage_03%>">
<input type=hidden name=Manufacturer_03 value="<%=Manufacturer_03%>">
<input type=hidden name=Functional_03 value="<%=Functional_03%>">

<input type=hidden name=Bulk_Code_04 value="<%=Bulk_Code_04%>">
<input type=hidden name=Product_Name_KFDA_04 value="<%=Product_Name_KFDA_04%>">
<input type=hidden name=P_Class_04 value="<%=P_Class_04%>">
<input type=hidden name=P_Capacity_04 value="<%=P_Capacity_04%>">
<input type=hidden name=P_Capacity_Unit_04 value="<%=P_Capacity_Unit_04%>">
<input type=hidden name=Period_of_Usage_04 value="<%=Period_of_Usage_04%>">
<input type=hidden name=Manufacturer_04 value="<%=Manufacturer_04%>">
<input type=hidden name=Functional_04 value="<%=Functional_04%>">

<input type=hidden name=Bulk_Code_05 value="<%=Bulk_Code_05%>">
<input type=hidden name=Product_Name_KFDA_05 value="<%=Product_Name_KFDA_05%>">
<input type=hidden name=P_Class_05 value="<%=P_Class_05%>">
<input type=hidden name=P_Capacity_05 value="<%=P_Capacity_05%>">
<input type=hidden name=P_Capacity_Unit_05 value="<%=P_Capacity_Unit_05%>">
<input type=hidden name=Period_of_Usage_05 value="<%=Period_of_Usage_05%>">
<input type=hidden name=Manufacturer_05 value="<%=Manufacturer_05%>">
<input type=hidden name=Functional_05 value="<%=Functional_05%>">

<input type=hidden name=Bulk_Code_06 value="<%=Bulk_Code_06%>">
<input type=hidden name=Product_Name_KFDA_06 value="<%=Product_Name_KFDA_06%>">
<input type=hidden name=P_Class_06 value="<%=P_Class_06%>">
<input type=hidden name=P_Capacity_06 value="<%=P_Capacity_06%>">
<input type=hidden name=P_Capacity_Unit_06 value="<%=P_Capacity_Unit_06%>">
<input type=hidden name=Period_of_Usage_06 value="<%=Period_of_Usage_06%>">
<input type=hidden name=Manufacturer_06 value="<%=Manufacturer_06%>">
<input type=hidden name=Functional_06 value="<%=Functional_06%>">

<input type=hidden name=Bulk_Code_07 value="<%=Bulk_Code_07%>">
<input type=hidden name=Product_Name_KFDA_07 value="<%=Product_Name_KFDA_07%>">
<input type=hidden name=P_Class_07 value="<%=P_Class_07%>">
<input type=hidden name=P_Capacity_07 value="<%=P_Capacity_07%>">
<input type=hidden name=P_Capacity_Unit_07 value="<%=P_Capacity_Unit_07%>">
<input type=hidden name=Period_of_Usage_07 value="<%=Period_of_Usage_07%>">
<input type=hidden name=Manufacturer_07 value="<%=Manufacturer_07%>">
<input type=hidden name=Functional_07 value="<%=Functional_07%>">

<input type=hidden name=Bulk_Code_08 value="<%=Bulk_Code_08%>">
<input type=hidden name=Product_Name_KFDA_08 value="<%=Product_Name_KFDA_08%>">
<input type=hidden name=P_Class_08 value="<%=P_Class_08%>">
<input type=hidden name=P_Capacity_08 value="<%=P_Capacity_08%>">
<input type=hidden name=P_Capacity_Unit_08 value="<%=P_Capacity_Unit_08%>">
<input type=hidden name=Period_of_Usage_08 value="<%=Period_of_Usage_08%>">
<input type=hidden name=Manufacturer_08 value="<%=Manufacturer_08%>">
<input type=hidden name=Functional_08 value="<%=Functional_08%>">
<input type=hidden name=Remarks_Finsh_Product value="<%=Remarks_Finsh_Product%>">




    <table cellspacing=0 cellpadding=0 border="0"  width="1024" cellpadding="0" cellspacing="1" style="word-wrap:break-word;word-break:break-all;">
    <tr> 
      <td width="312"  bgcolor="#D7F1FA" style="text-align:left; text-indent:0; margin:0; padding-top:15px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
      <b>▶ 완제품 납품 정보입력 </b></td>
       <td width="400"  bgcolor="#D7F1FA" style="text-align:left; text-indent:0; margin:0; padding-top:15px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
                 <script>
var winObject = null;

function popupWindow1(){
var settings ='toolbar=0,directories=0,status=no,menubar=0,scrollbars=auto,resizable=no,height=800,width=1400,left=50,top=50';
winObject = window.open("Lot_Search.asp", "test9", settings);
}

</script>
         &nbsp;<input type=button value="적합 로트 리스트(20/01~22/03) 검색" onClick="popupWindow1();">
         </td>
         
      <td width="312"  bgcolor="#D7F1FA" style="text-align:right; text-indent:0; margin:0; padding-top:15px; padding-right:0px; padding-bottom:5px; padding-left:0px;">
    <a href="javascript:history.go(-1)"><img src="images/back.gif"  border="0"></a></td>
    </tr>
</table>


     <table cellspacing=0 cellpadding=0 border="1"  width="1024" cellpadding="0" cellspacing="1" style="word-wrap:break-word;word-break:break-all;">
 
    <tr>
        <td width="120"  bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
                <b>납품 수량</b></td>
         <td width="221" style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
                 <input type="text" name="Delivery_Amount" size="10" maxlength="12"></td>
	
	   <td rowspan=8 width="120"  bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>Lot No</b></td>
        <td rowspan=8  width="221" style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
         <div style="line-height:250%;">
                                        1. <input type="text" name="Lot_number_01" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>2. <input type="text" name="Lot_number_02" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>3. <input type="text" name="Lot_number_03" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>4. <input type="text" name="Lot_number_04" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>5. <input type="text" name="Lot_number_05" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>6. <input type="text" name="Lot_number_06" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>7. <input type="text" name="Lot_number_07" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>8. <input type="text" name="Lot_number_08" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      </div></td>
      
	 <td rowspan=8 width="120" bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>사용기한</b></font></td>
	  <td rowspan=8 width="222" style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
      <div style="line-height:250%;">
          1. <input type="text" name="Expiration_Date_0 1" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>2. <input type="text" name="Expiration_Date_02" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>3. <input type="text" name="Expiration_Date_03" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>4. <input type="text" name="Expiration_Date_04" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>5. <input type="text" name="Expiration_Date_05" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>6. <input type="text" name="Expiration_Date_06" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>7. <input type="text" name="Expiration_Date_07" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      <br>8. <input type="text" name="Expiration_Date_08" maxlength="30" style="width:80%;height:100;border:1;overflow:visible;text-overflow:ellipsis;">
      </div></td>
	</tr>
	
	  <tr>
        
	 <td  bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
          <b>수량 Lot 구분</b></td>
        <td   style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
                   <select name="Lot_No_Divide">
                                     <option value="구분됨">구분됨</option>
                                     <option value="구분안됨">구분안됨</option></select></span></td>
	
	</tr>
	 <tr>
        
	 <td  bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
          <b>제품 종류</b></td>
        <td   style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
                   <select name="Good_class"><option></option>
                                     <option value="화장품">화장품</option>
                                     <option value="화장품이외">화장품이외</option>
                                     </select></span></td>
	
	</tr>
	 <tr>
        
	 <td  bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
          <b>관리품 판정</b></td>
        <td   style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
                   <select name="Judge_Result"><option></option>
                                     <option value="적합">적합</option>
                                     <option value="부적합">부적합</option>
                                     <option value="보류">보류</option>
                                     <option value="기입고">기입고</option></select></td>
	
	</tr>
	  <tr>
	   <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>제조사</b></font></td>
	  <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
      <input type="text" name="Supplier"  maxlength="40" value="<%=Manufacturer_01%>" style="width:90%;height:100%;solid:1px;overflow:visible;text-overflow:ellipsis;"></td>
		</tr>
		<tr>
	   <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>입고처</b></font></td>
	  <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
      <input type="text" name="Warehouse"  maxlength="40"  value="엘앤피물류" style="width:90%;height:100%;solid:1px;overflow:visible;text-overflow:ellipsis;"></td>
		</tr>
		<tr>
	   <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>관리번호</b></font></td>
	  <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
      <input type="text" name="Manage_No"  maxlength="20" style="width:90%;height:100%;solid:1px;overflow:visible;text-overflow:ellipsis;"></td>
		</tr>
		<tr>
	   <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>성적서 입수</b></font></td>
	  <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
       <select name="COA_Obtain"><option></option>
                                     <option value="입수">입수</option>
                                     <option value="미입수">미입수</option>
                                     <option value="기입수">기입수</option></select></td>
		</tr>
		   <%
to_year = year(date())

start_year = Cint(to_year)-2
end_year = Cint(to_year)+1

WH_Date=date()
tag="1"
%>


	<tr>
	   <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>입고일</b></font></td>
	  <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
       <select name="Warehouse_Year" id="Warehouse_Year" style="width:70">
							  <% for k=start_year to end_year %>
                <option value="<%=k%>"  <%if k=year(WH_Date) then%> selected <% end if %>><%=k%>년</option>
								<% next%>							
            </select>
              <select name="Warehouse_Month" id="Warehouse_Month"  style="width:55">
							  <% for k=1 to 12 %>
                <option value="<%=k%>"  <%if k=month(WH_Date) then%> selected <% end if %>><%=k%>월</option>
								<% next%>
              </select>
              <select name="Warehouse_Day" id="Warehouse_Day"  style="width:55">
							  <% for k=1 to 31 %>
                <option value="<%=k%>"  <%if k=day(WH_Date) then%> selected <% end if %>><%=k%>일</option>
								<% next%> </select></td>
  
        <td rowspan=2 bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
            <b>비&nbsp;&nbsp;&nbsp;&nbsp;고</b></td>
              <td rowspan=2 colspan=3 style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
     <textarea  name="Remarks" style="width:95%;height:100%;solid:1px;overflow:visible;text-overflow:ellipsis;" rows="4"></textarea></td>
    
		</tr>
		
       <tr> 
       <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>암&nbsp;&nbsp;&nbsp;&nbsp;호</b></td>
       <td style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
        <input type="password" name="Pass" size="5" maxlength="5">&nbsp;<b>※수정/삭제시 필요</b></td>
       </tr>
      <tr> 
       <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>첨부 파일1</b></td>
       <td colspan=5 style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
         <input type="file" name="sfile1" style="width:98%;height:100%;solid:1px;overflow:visible;text-overflow:ellipsis;"></td>
		</tr>
		 <tr> 
       <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>첨부 파일2</b></td>
       <td colspan=5 style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
         <input type="file" name="sfile2" style="width:98%;height:100%;solid:1px;overflow:visible;text-overflow:ellipsis;"></td>
		</tr>
		 <tr> 
       <td bgcolor="#CCFFCC" style="text-align:center; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:0px;">
         <b>첨부 파일3</b></td>
       <td colspan=5 style="text-align:left; text-indent:0; margin:0; padding-top:3px; padding-right:0px; padding-bottom:1px; padding-left:5px;">
         <input type="file" name="sfile3" style="width:98%;height:100%;solid:1px;overflow:visible;text-overflow:ellipsis;"></td>
		</tr>
		</table>
	 
	
  
      
       <table cellspacing=0 cellpadding=0 border="0"  width="1024" cellpadding="0" cellspacing="1" style="word-wrap:break-word;word-break:break-all;">
    <tr height=60> 
            <td  align="center" bgcolor="#D7F1FA"> 
          <input type="image" img src="images/regist.gif"  border="0">&nbsp;&nbsp;&nbsp;
          <img src="images/cancel.gif" border="0" OnClick="javascript:document.form.reset()" style="cursor:hand"></td>
            </tr>
					
          </form>
        </table> 
        <br><br style="line-height:50pt;"> 
</body>
</html>
<% end if %>   
