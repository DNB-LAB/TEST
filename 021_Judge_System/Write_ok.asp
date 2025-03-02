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
'첨부가 있는 전송이므로 업로드 컴포넌트 객체를 생성하여 전송을 처리한다
Set Upload = Server.CreateObject("TABSUpload4.Upload") 



'업로드된 파일을 저장할 서버의 폴더
Spath = Server.MapPath(".") & "\Upload_temp" '임시로 저장할 폴더
Upload.Start Spath


'사용자의 입력값을 전송받는다
Product_Code=Upload.Form("Product_Code")
Product_Code = trim(Replace(Product_Code,"'","''"))

Registor = Upload.Form("Registor")
Registor = trim(Replace(Registor,"'","''"))

Product_Name_DZ = Upload.Form("Product_Name_DZ")
Product_Name_DZ = trim(Replace(Product_Name_DZ,"'","''"))

Product_Name_Final = Upload.Form("Product_Name_Final")
Product_Name_Final = trim(Replace(Product_Name_Final,"'","''"))

Bulk_Code_01 = Upload.Form("Bulk_Code_01")
Bulk_Code_01 = trim(Replace(Bulk_Code_01,"'","''"))

Product_Name_KFDA_01 = Upload.Form("Product_Name_KFDA_01")
Product_Name_KFDA_01 = trim(Replace(Product_Name_KFDA_01,"'","''"))

P_Class_01 = Upload.Form("P_Class_01")
P_Class_01 = trim(Replace(P_Class_01,"'","''"))

P_Capacity_01 = Upload.Form("P_Capacity_01")
P_Capacity_01 = trim(Replace(P_Capacity_01,"'","''"))

P_Capacity_Unit_01 = Upload.Form("P_Capacity_Unit_01")
P_Capacity_Unit_01 = trim(Replace(P_Capacity_Unit_01,"'","''"))

Period_of_Usage_01 = Upload.Form("Period_of_Usage_01")
Period_of_Usage_01 = trim(Replace(Period_of_Usage_01,"'","''"))

Manufacturer_01 = Upload.Form("Manufacturer_01")
Manufacturer_01 = trim(Replace(Manufacturer_01,"'","''"))

Functional_01 = Upload.Form("Functional_01")
Functional_01 = trim(Replace(Functional_01,"'","''"))

Bulk_Code_02 = Upload.Form("Bulk_Code_02")
Bulk_Code_02 = trim(Replace(Bulk_Code_02,"'","''"))

Product_Name_KFDA_02 = Upload.Form("Product_Name_KFDA_02")
Product_Name_KFDA_02 = trim(Replace(Product_Name_KFDA_02,"'","''"))

P_Class_02 = Upload.Form("P_Class_02")
P_Class_02 = trim(Replace(P_Class_02,"'","''"))

P_Capacity_02 = Upload.Form("P_Capacity_02")
P_Capacity_02 = trim(Replace(P_Capacity_02,"'","''"))

P_Capacity_Unit_02 = Upload.Form("P_Capacity_Unit_02")
P_Capacity_Unit_02 = trim(Replace(P_Capacity_Unit_02,"'","''"))

Period_of_Usage_02 = Upload.Form("Period_of_Usage_02")
Period_of_Usage_02 = trim(Replace(Period_of_Usage_02,"'","''"))

Manufacturer_02 = Upload.Form("Manufacturer_02")
Manufacturer_02 = trim(Replace(Manufacturer_02,"'","''"))

Functional_02 = Upload.Form("Functional_02")
Functional_02 = trim(Replace(Functional_02,"'","''"))

Bulk_Code_03 = Upload.Form("Bulk_Code_03")
Bulk_Code_03 = trim(Replace(Bulk_Code_03,"'","''"))

Product_Name_KFDA_03 = Upload.Form("Product_Name_KFDA_03")
Product_Name_KFDA_03 = trim(Replace(Product_Name_KFDA_03,"'","''"))

P_Class_03 = Upload.Form("P_Class_03")
P_Class_03 = trim(Replace(P_Class_03,"'","''"))

P_Capacity_03 = Upload.Form("P_Capacity_03")
P_Capacity_03 = trim(Replace(P_Capacity_03,"'","''"))

P_Capacity_Unit_03 = Upload.Form("P_Capacity_Unit_03")
P_Capacity_Unit_03 = trim(Replace(P_Capacity_Unit_03,"'","''"))

Period_of_Usage_03 = Upload.Form("Period_of_Usage_03")
Period_of_Usage_03 = trim(Replace(Period_of_Usage_03,"'","''"))

Manufacturer_03 = Upload.Form("Manufacturer_03")
Manufacturer_03 = trim(Replace(Manufacturer_03,"'","''"))

Functional_03 = Upload.Form("Functional_03")
Functional_03 = trim(Replace(Functional_03,"'","''"))

Bulk_Code_04 = Upload.Form("Bulk_Code_04")
Bulk_Code_04 = trim(Replace(Bulk_Code_04,"'","''"))

Product_Name_KFDA_04 = Upload.Form("Product_Name_KFDA_04")
Product_Name_KFDA_04 = trim(Replace(Product_Name_KFDA_04,"'","''"))

P_Class_04 = Upload.Form("P_Class_04")
P_Class_04 = trim(Replace(P_Class_04,"'","''"))

P_Capacity_04 = Upload.Form("P_Capacity_04")
P_Capacity_04 = trim(Replace(P_Capacity_04,"'","''"))

P_Capacity_Unit_04 = Upload.Form("P_Capacity_Unit_04")
P_Capacity_Unit_04 = trim(Replace(P_Capacity_Unit_04,"'","''"))

Period_of_Usage_04 = Upload.Form("Period_of_Usage_04")
Period_of_Usage_04 = trim(Replace(Period_of_Usage_04,"'","''"))

Manufacturer_04 = Upload.Form("Manufacturer_04")
Manufacturer_04 = trim(Replace(Manufacturer_04,"'","''"))

Functional_04 = Upload.Form("Functional_04")
Functional_04 = trim(Replace(Functional_04,"'","''"))


Bulk_Code_05 = Upload.Form("Bulk_Code_05")
Bulk_Code_05 = trim(Replace(Bulk_Code_05,"'","''"))

Product_Name_KFDA_05 = Upload.Form("Product_Name_KFDA_05")
Product_Name_KFDA_05 = trim(Replace(Product_Name_KFDA_05,"'","''"))

P_Class_05 = Upload.Form("P_Class_05")
P_Class_05 = trim(Replace(P_Class_05,"'","''"))

P_Capacity_05 = Upload.Form("P_Capacity_05")
P_Capacity_05 = trim(Replace(P_Capacity_05,"'","''"))

P_Capacity_Unit_05 = Upload.Form("P_Capacity_Unit_05")
P_Capacity_Unit_05 = trim(Replace(P_Capacity_Unit_05,"'","''"))

Period_of_Usage_05 = Upload.Form("Period_of_Usage_05")
Period_of_Usage_05 = trim(Replace(Period_of_Usage_05,"'","''"))

Manufacturer_05 = Upload.Form("Manufacturer_05")
Manufacturer_05 = trim(Replace(Manufacturer_05,"'","''"))

Functional_05 = Upload.Form("Functional_05")
Functional_05 = trim(Replace(Functional_05,"'","''"))

Bulk_Code_06 = Upload.Form("Bulk_Code_06")
Bulk_Code_06 = trim(Replace(Bulk_Code_06,"'","''"))

Product_Name_KFDA_06 = Upload.Form("Product_Name_KFDA_06")
Product_Name_KFDA_06 = trim(Replace(Product_Name_KFDA_06,"'","''"))

P_Class_06 = Upload.Form("P_Class_06")
P_Class_06 = trim(Replace(P_Class_06,"'","''"))

P_Capacity_06 = Upload.Form("P_Capacity_06")
P_Capacity_06 = trim(Replace(P_Capacity_06,"'","''"))

P_Capacity_Unit_06 = Upload.Form("P_Capacity_Unit_06")
P_Capacity_Unit_06 = trim(Replace(P_Capacity_Unit_06,"'","''"))

Period_of_Usage_06 = Upload.Form("Period_of_Usage_06")
Period_of_Usage_06 = trim(Replace(Period_of_Usage_06,"'","''"))

Manufacturer_06 = Upload.Form("Manufacturer_06")
Manufacturer_06 = trim(Replace(Manufacturer_06,"'","''"))

Functional_06 = Upload.Form("Functional_06")
Functional_06 = trim(Replace(Functional_06,"'","''"))

Bulk_Code_07 = Upload.Form("Bulk_Code_07")
Bulk_Code_07 = trim(Replace(Bulk_Code_07,"'","''"))

Product_Name_KFDA_07 = Upload.Form("Product_Name_KFDA_07")
Product_Name_KFDA_07 = trim(Replace(Product_Name_KFDA_07,"'","''"))

P_Class_07 = Upload.Form("P_Class_07")
P_Class_07 = trim(Replace(P_Class_07,"'","''"))

P_Capacity_07 = Upload.Form("P_Capacity_07")
P_Capacity_07 = trim(Replace(P_Capacity_07,"'","''"))

P_Capacity_Unit_07 = Upload.Form("P_Capacity_Unit_07")
P_Capacity_Unit_07 = trim(Replace(P_Capacity_Unit_07,"'","''"))

Period_of_Usage_07 = Upload.Form("Period_of_Usage_07")
Period_of_Usage_07 = trim(Replace(Period_of_Usage_07,"'","''"))

Manufacturer_07 = Upload.Form("Manufacturer_07")
Manufacturer_07 = trim(Replace(Manufacturer_07,"'","''"))

Functional_07 = Upload.Form("Functional_07")
Functional_07 = trim(Replace(Functional_07,"'","''"))

Bulk_Code_08 = Upload.Form("Bulk_Code_08")
Bulk_Code_08 = trim(Replace(Bulk_Code_08,"'","''"))

Product_Name_KFDA_08 = Upload.Form("Product_Name_KFDA_08")
Product_Name_KFDA_08 = trim(Replace(Product_Name_KFDA_08,"'","''"))

P_Class_08 = Upload.Form("P_Class_08")
P_Class_08 = trim(Replace(P_Class_08,"'","''"))

P_Capacity_08 = Upload.Form("P_Capacity_08")
P_Capacity_08 = trim(Replace(P_Capacity_08,"'","''"))

P_Capacity_Unit_08 = Upload.Form("P_Capacity_Unit_08")
P_Capacity_Unit_08 = trim(Replace(P_Capacity_Unit_08,"'","''"))

Period_of_Usage_08 = Upload.Form("Period_of_Usage_08")
Period_of_Usage_08 = trim(Replace(Period_of_Usage_08,"'","''"))

Manufacturer_08 = Upload.Form("Manufacturer_08")
Manufacturer_08 = trim(Replace(Manufacturer_08,"'","''"))

Functional_08 = Upload.Form("Functional_08")
Functional_08 = trim(Replace(Functional_08,"'","''"))


Remarks_Finsh_Product = Upload.Form("Remarks_Finsh_Product")
Remarks_Finsh_Product = trim(Replace(Remarks_Finsh_Product,"'","''"))

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



STime = Now() '글쓴 시각을 구한다

Sdate = Date() '의뢰일자를 구하기 위해서 함수

UTime = now() '최초 작성시 글쓴시간을 구하고 나중에 수정함에 따라서 업데이트 파일에서 고쳐줄 것


Registor_name = session("mname")

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



'디비연결
Set DB =Server.CreateObject("ADODB.Connection")
DB.open ConnString

'새글의 글번호를 얻기 위해 가장 큰 글번호를 찾는다
SQL = "SELECT MAX(sid) FROM  " & STable
Set RS = DB.Execute(SQL)
IF IsNull(RS(0)) Then
	Nsid = 1
Else
	Nsid = RS(0) +1 '새글의 글번호로 한다
End IF


RS.Close
Set RS = Nothing


'3개의 파일을 각각 업되면 저장하고 파일이름 구한다

If Upload.Form("sFile1").FileSize <> 0 Then
 Upload.Form("sfile1").Save("D:\000_LNP_Db\021_Judge_System\Upload_01") '// 중복시 자동 시리얼이 붙음
 sFile1 = Upload.Form("sfile1").SaveName '// 저장된 전체경로와 파일이름 
 sFile1 = Mid(sFile1,instrrev(sFile1,"\")+1)'// 저장된 파일이름 

sFile1path = Spath & "\" & sFile1
on error resume next
Upload.Form("sFile1").SaveAs(sFile1path) '//저장한 파일이름
END IF

If Upload.Form("sFile2").FileSize <> 0 Then
 Upload.Form("sfile2").Save("D:\000_LNP_Db\021_Judge_System\Upload_02") '// 중복시 자동 시리얼이 붙음
 sFile2 = Upload.Form("sfile2").SaveName '// 저장된 전체경로와 파일이름 
 sFile2 = Mid(sFile2,instrrev(sFile2,"\")+1)'// 저장된 파일이름 

sFile2path = Spath & "\" & sFile2
on error resume next
Upload.Form("sFile2").SaveAs(sFile2path)
END IF


If Upload.Form("sFile3").FileSize <> 0 Then
 Upload.Form("sfile3").Save("D:\000_LNP_Db\021_Judge_System\Upload_03") '// 중복시 자동 시리얼이 붙음
 sFile3 = Upload.Form("sfile3").SaveName '// 저장된 전체경로와 파일이름 
 sFile3 = Mid(sFile3,instrrev(sFile3,"\")+1)'// 저장된 파일이름 

sFile3path = Spath & "\" & sFile3
on error resume next
Upload.Form("sFile3").SaveAs(sFile3path)
END IF



Set Upload = Nothing


'새글의 값을 데이타베이스에 저장한다
SQL = "INSERT INTO " & STable & " VALUES ("
SQL = SQL & Nsid
SQL = SQL & "," & Nsid & " , 1, 0"
SQL = SQL & ",'" & Product_Code & "'"
SQL = SQL & ",'" & Registor & "'"
SQL = SQL & ",'" & Product_Name_DZ & "'"
SQL = SQL & ",'" & Product_Name_Final & "'"
SQL = SQL & ",'" & Bulk_Code_01 & "'"
SQL = SQL & ",'" & Product_Name_KFDA_01 & "'"
SQL = SQL & ",'" & P_Class_01 & "'"
SQL = SQL & ",'" & P_Capacity_01 & "'"
SQL = SQL & ",'" & P_Capacity_Unit_01 & "'"
SQL = SQL & ",'" & Period_of_Usage_01 & "'"
SQL = SQL & ",'" & Manufacturer_01 & "'"
SQL = SQL & ",'" & Functional_01 & "'"

SQL = SQL & ",'" & Bulk_Code_02 & "'"
SQL = SQL & ",'" & Product_Name_KFDA_02 & "'"
SQL = SQL & ",'" & P_Class_02 & "'"
SQL = SQL & ",'" & P_Capacity_02 & "'"
SQL = SQL & ",'" & P_Capacity_Unit_02 & "'"
SQL = SQL & ",'" & Period_of_Usage_02 & "'"
SQL = SQL & ",'" & Manufacturer_02 & "'"
SQL = SQL & ",'" & Functional_02 & "'"


SQL = SQL & ",'" & Bulk_Code_03 & "'"
SQL = SQL & ",'" & Product_Name_KFDA_03 & "'"
SQL = SQL & ",'" & P_Class_03 & "'"
SQL = SQL & ",'" & P_Capacity_03 & "'"
SQL = SQL & ",'" & P_Capacity_Unit_03 & "'"
SQL = SQL & ",'" & Period_of_Usage_03 & "'"
SQL = SQL & ",'" & Manufacturer_03 & "'"
SQL = SQL & ",'" & Functional_03 & "'"


SQL = SQL & ",'" & Bulk_Code_04 & "'"
SQL = SQL & ",'" & Product_Name_KFDA_04 & "'"
SQL = SQL & ",'" & P_Class_04 & "'"
SQL = SQL & ",'" & P_Capacity_04 & "'"
SQL = SQL & ",'" & P_Capacity_Unit_04 & "'"
SQL = SQL & ",'" & Period_of_Usage_04 & "'"
SQL = SQL & ",'" & Manufacturer_04 & "'"
SQL = SQL & ",'" & Functional_04 & "'"


SQL = SQL & ",'" & Bulk_Code_05 & "'"
SQL = SQL & ",'" & Product_Name_KFDA_05 & "'"
SQL = SQL & ",'" & P_Class_05 & "'"
SQL = SQL & ",'" & P_Capacity_05 & "'"
SQL = SQL & ",'" & P_Capacity_Unit_05 & "'"
SQL = SQL & ",'" & Period_of_Usage_05 & "'"
SQL = SQL & ",'" & Manufacturer_05 & "'"
SQL = SQL & ",'" & Functional_05 & "'"


SQL = SQL & ",'" & Bulk_Code_06 & "'"
SQL = SQL & ",'" & Product_Name_KFDA_06 & "'"
SQL = SQL & ",'" & P_Class_06 & "'"
SQL = SQL & ",'" & P_Capacity_06 & "'"
SQL = SQL & ",'" & P_Capacity_Unit_06 & "'"
SQL = SQL & ",'" & Period_of_Usage_06 & "'"
SQL = SQL & ",'" & Manufacturer_06 & "'"
SQL = SQL & ",'" & Functional_06 & "'"


SQL = SQL & ",'" & Bulk_Code_07 & "'"
SQL = SQL & ",'" & Product_Name_KFDA_07 & "'"
SQL = SQL & ",'" & P_Class_07 & "'"
SQL = SQL & ",'" & P_Capacity_07 & "'"
SQL = SQL & ",'" & P_Capacity_Unit_07 & "'"
SQL = SQL & ",'" & Period_of_Usage_07 & "'"
SQL = SQL & ",'" & Manufacturer_07 & "'"
SQL = SQL & ",'" & Functional_07 & "'"


SQL = SQL & ",'" & Bulk_Code_08 & "'"
SQL = SQL & ",'" & Product_Name_KFDA_08 & "'"
SQL = SQL & ",'" & P_Class_08 & "'"
SQL = SQL & ",'" & P_Capacity_08 & "'"
SQL = SQL & ",'" & P_Capacity_Unit_08 & "'"
SQL = SQL & ",'" & Period_of_Usage_08 & "'"
SQL = SQL & ",'" & Manufacturer_08 & "'"
SQL = SQL & ",'" & Functional_08 & "'"


SQL = SQL & ",'" & Remarks_Finsh_Product & "'"
SQL = SQL & ",'" & Delivery_Amount & "'"

SQL = SQL & ",'" & Lot_number_01 & "'"
SQL = SQL & ",'" & Lot_number_02 & "'"
SQL = SQL & ",'" & Lot_number_03 & "'"
SQL = SQL & ",'" & Lot_number_04 & "'"
SQL = SQL & ",'" & Lot_number_05 & "'"
SQL = SQL & ",'" & Lot_number_06 & "'"
SQL = SQL & ",'" & Lot_number_07 & "'"
SQL = SQL & ",'" & Lot_number_08 & "'"
SQL = SQL & ",'" & Expiration_Date_01 & "'"
SQL = SQL & ",'" & Expiration_Date_02 & "'"
SQL = SQL & ",'" & Expiration_Date_03 & "'"
SQL = SQL & ",'" & Expiration_Date_04 & "'"
SQL = SQL & ",'" & Expiration_Date_05 & "'"
SQL = SQL & ",'" & Expiration_Date_06 & "'"
SQL = SQL & ",'" & Expiration_Date_07 & "'"
SQL = SQL & ",'" & Expiration_Date_08 & "'"

SQL = SQL & ",'" & Lot_No_Divide & "'"
SQL = SQL & ",'" & Good_class & "'"
SQL = SQL & ",'" & Judge_Result & "'"
SQL = SQL & ",'" & Supplier & "'"
SQL = SQL & ",'" & Warehouse & "'"
SQL = SQL & ",'" & Manage_No & "'"
SQL = SQL & ",'" & COA_Obtain & "'"

SQL = SQL & ",'" & Warehouse_Year & "'"
SQL = SQL & ",'" & Warehouse_Month & "'"
SQL = SQL & ",'" & Warehouse_Day & "'"
SQL = SQL & ",'" & Warehouse_Date & "'"


SQL = SQL & ",'" & Pass & "'"
SQL = SQL & ",'" & Remarks & "'"
SQL = SQL & ",'" & Registor_name & "'"
SQL = SQL & ",1"
SQL = SQL & ",'" & STime & "'"
SQL = SQL & ",'" & Sdate & "'"
SQL = SQL & ",'" & uTime & "'"
SQL = SQL & ",'" & sFile1 & "'"
SQL = SQL & ",'" & sFile2 & "'"
SQL = SQL & ",'" & sFile3 & "')"

DB.Execute SQL

DB.Close
SET DB=Nothing

URL = "list.asp?" & Var2
Response.Redirect URl
%>


 <% end if %>