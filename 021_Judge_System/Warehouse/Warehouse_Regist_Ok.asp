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
<body >
<body leftmargin="0" topmargin="0" bgcolor="#D7F1FA">
<script language="javascript">
		alert("로그인이 필요합니다. 로그인 하세요! \n\n\혹은 로그인 됐더라도 오래되어 종료되었습니다.. \n\n\재 로그인이 필요합니다.  login please !!!");
	window.open('../../../Log_in.asp','end','width=310,height=190,top=270, left=350');
</script>

<% else %>




<%

Original_sid = Request.QueryString("Original_sid")

Set Upload = Server.CreateObject("TABSUpload4.Upload") 

'최대업로드용량
Upload.MaxBytesToAbort = Maxsize * 1024 * 1024 'Maxsize 는 inc.asp에서 정의

'업로드된 파일을 저장할 서버의 폴더
Spath = Server.MapPath(".") & "\Upload_temp" '임시로 저장할 폴더
Upload.Start Spath


'사용자의 입력값을 전송받는다

Original_sid = Upload.Form("Original_sid")
Original_sid = trim(Replace(Original_sid,"'","''"))

W_Judge_Method = Upload.Form("W_Judge_Method")
W_Judge_Method = trim(Replace(W_Judge_Method,"'","''"))


W_Judge_Result = Upload.Form("W_Judge_Result")
W_Judge_Result = trim(Replace(W_Judge_Result,"'","''"))

W_Remarks = Upload.Form("W_Remarks")
W_Remarks = trim(Replace(W_Remarks,"'","''"))

Pass = Upload.Form("Pass")
Pass = trim(Replace(Pass,"'","''"))


W_Stime = now()
W_Sdate = Date() '글쓴 시각을 구한다
W_Utime = now() '글쓴 시각을 구한다

W_Registor = session("mname")

%>

<!-- #include file="inc_Warwhouse.asp" -->

<%

'디비연결
Set DB =Server.CreateObject("ADODB.Connection")
DB.open ConnString

'새글의 글번호를 얻기 위해 가장 큰 글번호를 찾는다

SQL = "SELECT MAX(sid) FROM " & STable
Set RS = DB.Execute(SQL)

IF IsNull(RS(0)) Then
	Nsid = 1
Else
	Nsid = RS(0) +1 '새글의 글번호로 한다
End IF

RS.Close
Set RS=nothing


If Upload.Form("sFile1").FileSize <> 0 Then
 Upload.Form("sfile1").Save("D:\000_LNP_Db\021_Judge_System\Warehouse\Upload_01") '// 중복시 자동 시리얼이 붙음
 sFile1 = Upload.Form("sfile1").SaveName '// 저장된 전체경로와 파일이름 
 sFile1 = Mid(sFile1,instrrev(sFile1,"\")+1)'// 저장된 파일이름 

sFile1path = Spath & "\" & sFile1
on error resume next
Upload.Form("sFile1").SaveAs(sFile1path) '//저장한 파일이름
END IF



If Upload.Form("sFile2").FileSize <> 0 Then
 Upload.Form("sfile2").Save("D:\000_LNP_Db\021_Judge_System\Warehouse\Upload_02") '// 중복시 자동 시리얼이 붙음
 sFile2 = Upload.Form("sfile2").SaveName '// 저장된 전체경로와 파일이름 
 sFile2 = Mid(sFile2,instrrev(sFile2,"\")+1)'// 저장된 파일이름 

sFile2path = Spath & "\" & sFile2
on error resume next
Upload.Form("sFile2").SaveAs(sFile2path)
END IF

If Upload.Form("sFile3").FileSize <> 0 Then
Upload.Form("sfile3").Save("D:\000_LNP_Db\021_Judge_System\Warehouse\Upload_03") '// 중복시 자동 시리얼이 붙음
 sFile3 = Upload.Form("sfile3").SaveName '// 저장된 전체경로와 파일이름 
 sFile3 = Mid(sFile3,instrrev(sFile3,"\")+1)'// 저장된 파일이름 


sFile3path = Spath & "\" & sFile3
on error resume next
Upload.Form("sFile3").SaveAs(sFile3path)
END IF


Set Upload = Nothing


Original_sid = Request.QueryString("Original_sid")



'새글의 값을 데이타베이스에 저장한다
SQL = "INSERT INTO " & STable & " VALUES ("
SQL = SQL & Nsid
SQL = SQL & "," & Nsid & " , 1, 0"
SQL = SQL & ",'" & Original_sid & "'"
SQL = SQL & ",'" & W_Registor & "'"
SQL = SQL & ",'" & W_Judge_Method & "'"
SQL = SQL & ",'" & W_Judge_Result & "'"
SQL = SQL & ",'" & W_Remarks & "'"
SQL = SQL & ",'" & Pass & "'"


SQL = SQL & ",1"
SQL = SQL & ",'" & W_Stime & "'"
SQL = SQL & ",'" & W_Sdate & "'"
SQL = SQL & ",'" & W_Utime & "'"
SQL = SQL & ",'" & sFile1 & "'"
SQL = SQL & ",'" & sFile2 & "'"
SQL = SQL & ",'" & sFile3 & "')"
DB.Execute SQL


DB.Close
Set DB = Nothing

URL = "../view.asp?sid=" & Original_sid
Response.Redirect URL
%>


 <% end if %>


