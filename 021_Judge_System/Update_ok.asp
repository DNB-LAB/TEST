<%@Language="VBScript" CODEPAGE="65001" %>
<% 
Response.CharSet="utf-8" 
Session.codepage="65001" 
Response.codepage="65001" 
Response.ContentType="text/html;charset=utf-8" 
%>

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


<%
'수정할 글번호를 전송받음(URL전송)
Sid= Request.QueryString("sid")
Var3=Var2  & "&sid=" & Sid

'첨부가 있는 전송이므로 업로드 컴포넌트 객체를 생성하여 전송을 처리한다
Set Upload = Server.CreateObject("TABSUpload4.Upload") 

'최대업로드용량
Upload.MaxBytesToAbort = Maxsize * 1024 * 1024 'Maxsize 는 inc.asp에서 정의

Spath = Server.MapPath(".") & "\Upload_temp" '임시로 저장할 폴더
Upload.Start Spath

'사용자의 입력값을 전송받는다
Delivery_Amount = Upload.Form("Delivery_Amount")
Delivery_Amount = trim(Replace(Delivery_Amount,"'","''"))

Lot_number_01 = Upload.Form("Lot_number_01")
Lot_number_01 = trim(Replace(Lot_number_01,"'","''"))

Lot_number_02 = Upload.Form("Lot_number_02")
Lot_number_02 = trim(Replace(Lot_number_02,"'","''"))

Lot_number_03 = Upload.Form("Lot_number_03")
Lot_number_03 = trim(Replace(Lot_number_03,"'","''"))

Lot_number_04 = Upload.Form("Lot_number_04")
Lot_number_04 = trim(Replace(Lot_number_04,"'","''"))

Lot_number_05 = Upload.Form("Lot_number_05")
Lot_number_05 = trim(Replace(Lot_number_05,"'","''"))

Lot_number_06 = Upload.Form("Lot_number_06")
Lot_number_06 = trim(Replace(Lot_number_06,"'","''"))

Lot_number_07 = Upload.Form("Lot_number_07")
Lot_number_07 = trim(Replace(Lot_number_07,"'","''"))

Lot_number_08 = Upload.Form("Lot_number_08")
Lot_number_08 = trim(Replace(Lot_number_08,"'","''"))

Expiration_Date_01 = Upload.Form("Expiration_Date_01")
Expiration_Date_01 = trim(Replace(Expiration_Date_01,"'","''"))

Expiration_Date_02 = Upload.Form("Expiration_Date_02")
Expiration_Date_02 = trim(Replace(Expiration_Date_02,"'","''"))

Expiration_Date_03 = Upload.Form("Expiration_Date_03")
Expiration_Date_03 = trim(Replace(Expiration_Date_03,"'","''"))

Expiration_Date_04 = Upload.Form("Expiration_Date_04")
Expiration_Date_04 = trim(Replace(Expiration_Date_04,"'","''"))

Expiration_Date_05 = Upload.Form("Expiration_Date_05")
Expiration_Date_05 = trim(Replace(Expiration_Date_05,"'","''"))

Expiration_Date_06 = Upload.Form("Expiration_Date_06")
Expiration_Date_06 = trim(Replace(Expiration_Date_06,"'","''"))

Expiration_Date_07 = Upload.Form("Expiration_Date_07")
Expiration_Date_07 = trim(Replace(Expiration_Date_07,"'","''"))

Expiration_Date_08 = Upload.Form("Expiration_Date_08")
Expiration_Date_08 = trim(Replace(Expiration_Date_08,"'","''"))

Lot_No_Divide  = Upload.Form("Lot_No_Divide")
Lot_No_Divide = trim(Replace(Lot_No_Divide,"'","''"))

Good_class  = Upload.Form("Good_class")
Good_class = trim(Replace(Good_class,"'","''"))

Judge_Result  = Upload.Form("Judge_Result")
Judge_Result = trim(Replace(Judge_Result,"'","''"))

Supplier  = Upload.Form("Supplier")
Supplier = trim(Replace(Supplier,"'","''"))

Warehouse  = Upload.Form("Warehouse")
Warehouse = trim(Replace(Warehouse,"'","''"))

Manage_No  = Upload.Form("Manage_No")
Manage_No = trim(Replace(Manage_No,"'","''"))

COA_Obtain  = Upload.Form("COA_Obtain")
COA_Obtain = trim(Replace(COA_Obtain,"'","''"))

Warehouse_Year  = Upload.Form("Warehouse_Year")
Warehouse_Year = trim(Replace(Warehouse_Year,"'","''"))


Warehouse_Month  = Upload.Form("Warehouse_Month")
Warehouse_Month = trim(Replace(Warehouse_Month,"'","''"))

Warehouse_Day  = Upload.Form("Warehouse_Day")
Warehouse_Day = trim(Replace(Warehouse_Day,"'","''"))


Pass = Upload.Form("pass")
Pass = trim(Replace(Pass, "'","''"))

Remarks  = Upload.Form("Remarks")
Remarks = trim(Replace(Remarks,"'","''"))

UTime = now() '글 수정한 시각을 구한다

Warehouse_Date = Warehouse_Year & "-" & Warehouse_Month & "-" & Warehouse_Day  '등록자가 선택한 입고일


%>
<%
if not isdate(Warehouse_Date) then
%>
<body bgcolor="#D7F1FA">
<script laguage="javascript">
<!--
  alert("선택한 입고일 없는 날짜입니다.  \n\n다시 확인하시기 바랍니다.");
  history.go(-1);
// -->
</script>


<%
reponse.end
end if 
%>

<%

Warehouse_Date = CDATE(Warehouse_Date)  'Warehouse_Date를 날짜형으로 변환시켜 준다.  나증에 날짜별로 불러올 수 있도록 하기 위하여...

Set DB =Server.Createobject("ADODB.Connection")
DB.open ConnString

'기존  비밀번호가 맞는지확인하고 파일의 이름을 가져옴
SQL = " SELECT Pass, sFile1, sFile2 ,sFile3  FROM AL_021_Judge_System"
SQL = SQL & " WHERE sid = " & sid
Set RS = DB.Execute(SQL)

OldPass=Rs("Pass")
Oldfile1=RS("sFile1")
Oldfile2=RS("sFile2")
Oldfile3=RS("sFile3")

RS.close
set RS=nothing

'db에 저장된 패스워드와 사용자가 입력한 패스워드를 비교함
IF (Pass = OldPass) or (Pass=adminpass) then  'adminpass는 관리자용==> inc.asp에 정의********************************

  '일치할 경우 수정한다
'/////////




'//첨부파일1이 새로 업되면 저장한다

If Upload.Form("sFile1").FileSize <> 0 Then
'기존 첨부파일을 삭제한다
If Oldfile1<>"" then

Set FS=Server.CreateObject("Scripting.FileSystemObject")
OldfilePath1=Spath & "\" & Oldfile1
On Error Resume Next
FS.DeleteFile(OldfilePath1)
set FS=nothing
End if

'새로 업된 첨부파일1을  저장한다
Upload.Form("sfile1").Save("D:\000_LNP_Db\021_Judge_System\Upload_01") '// 중복시 자동 시리얼이 붙음
 sFile1 = Upload.Form("sfile1").SaveName '// 저장된 전체경로와 파일이름 
 sFile1 = Mid(sFile1,instrrev(sFile1,"\")+1)'// 저장된 파일이름 

sFile1path = Spath & "\" & sFile1
on error resume next
Upload.Form("sFile1").SaveAs(sFile1path)
end if

'//첨부파일2이 새로 업되면 저장한다
If Upload.Form("sFile2").FileSize <> 0 Then
'기존 첨부파일2를 삭제한다
If Oldfile2<>"" then
Set FS=Server.CreateObject("Scripting.FileSystemObject")
OldfilePath2=Spath & "\" & Oldfile2
On Error Resume Next
FS.DeleteFile(OldfilePath1)
set FS=nothing
End if

'새로 업된 첨부파일2를  저장한다
Upload.Form("sFile2").Save("D:\000_LNP_Db\021_Judge_System\Upload_02") '// 중복시 자동 시리얼이 붙음
 sFile2 = Upload.Form("sFile2").SaveName '// 저장된 전체경로와 파일이름 
 sFile2 = Mid(sFile2,instrrev(sFile2,"\")+1)'// 저장된 파일이름 

sFile2path = Spath & "\" & sFile2
on error resume next
Upload.Form("sFile2").SaveAs(sFile2path)
end if




'//첨부파일3이 새로 업되면 저장한다
If Upload.Form("sFile3").FileSize <> 0 Then
'기존 첨부파일3를 삭제한다
If Oldfile3<>"" then
Set FS=Server.CreateObject("Scripting.FileSystemObject")
OldfilePath3=Spath & "\" & Oldfile3
On Error Resume Next
FS.DeleteFile(OldfilePath1)
set FS=nothing
End if

'새로 업된 첨부파일3를  저장한다
Upload.Form("sFile3").Save("D:\000_LNP_Db\021_Judge_System\Upload_03") '// 중복시 자동 시리얼이 붙음
 sFile3 = Upload.Form("sFile3").SaveName '// 저장된 전체경로와 파일이름 
 sFile3 = Mid(sFile3,instrrev(sFile3,"\")+1)'// 저장된 파일이름 

sFile3path = Spath & "\" & sFile3
on error resume next
Upload.Form("sFile3").SaveAs(sFile3path)
end if




'테이블에 저장한다
SQL = " UPDATE AL_021_Judge_System SET"
  SQL = SQL & " Delivery_Amount='" & Delivery_Amount & "'"
  SQL = SQL & ",Lot_number_01='" & Lot_number_01 & "'"
  SQL = SQL & ",Lot_number_02='" & Lot_number_02 & "'"
  SQL = SQL & ",Lot_number_03='" & Lot_number_03 & "'"
  SQL = SQL & ",Lot_number_04='" & Lot_number_04 & "'"
  SQL = SQL & ",Lot_number_05='" & Lot_number_05 & "'"
  SQL = SQL & ",Lot_number_06='" & Lot_number_06 & "'"
  SQL = SQL & ",Lot_number_07='" & Lot_number_07 & "'"
  SQL = SQL & ",Lot_number_08='" & Lot_number_08 & "'"
  
  SQL = SQL & ",Expiration_Date_01='" & Expiration_Date_01 & "'"
  SQL = SQL & ",Expiration_Date_02='" & Expiration_Date_02 & "'"
  SQL = SQL & ",Expiration_Date_03='" & Expiration_Date_03 & "'"
  SQL = SQL & ",Expiration_Date_04='" & Expiration_Date_04 & "'"
  SQL = SQL & ",Expiration_Date_05='" & Expiration_Date_05 & "'"
  SQL = SQL & ",Expiration_Date_06='" & Expiration_Date_06 & "'"
  SQL = SQL & ",Expiration_Date_07='" & Expiration_Date_07 & "'"
  SQL = SQL & ",Expiration_Date_08='" & Expiration_Date_08 & "'"
  
  SQL = SQL & ",Lot_No_Divide='" & Lot_No_Divide & "'"
  SQL = SQL & ",Good_class='" & Good_class & "'"
  SQL = SQL & ",Judge_Result='" & Judge_Result & "'"
  SQL = SQL & ",Supplier='" & Supplier & "'"
  SQL = SQL & ",Warehouse='" & Warehouse & "'"
  SQL = SQL & ",Manage_No='" & Manage_No & "'"
  SQL = SQL & ",COA_Obtain='" & COA_Obtain & "'"
  
  SQL = SQL & ",Warehouse_Year='" & Warehouse_Year & "'"
  SQL = SQL & ",Warehouse_Month='" & Warehouse_Month & "'"
  SQL = SQL & ",Warehouse_Day='" & Warehouse_Day & "'"
  SQL = SQL & ",Warehouse_Date='" & Warehouse_Date & "'"
  
  SQL = SQL & ",Remarks='" & Remarks & "'"

  SQL = SQL & ",Utime='" & Utime & "'"
  
  
if sFile1<>"" then '// 파일이름은 업된 경우만 수정한다
SQL = SQL & " ,sFile1='" & sFile1 & "'"
end if
if sFile2<>"" then
SQL = SQL & " ,sFile2='" & sFile2 & "'"
end if
if sFile3<>"" then
SQL = SQL & " ,sFile3='" & sFile3 & "'"
end if
SQL = SQL & " WHERE sid = " & sid

DB.Execute(SQL)

DB.close
set DB=nothing


URL = "view.asp?sid=" &sid
Response.Redirect URL

Else
	
%>
<body bgcolor="#D7F1FA">
<script language="javascript">
alert("암호가 틀립니다. -_- !!");
history.back();
</script>

 <% end if %>
 <% end if %>
