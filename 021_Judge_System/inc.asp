<%
'INCLUDE ȭ��

'��� ȭ�Ͽ� ����ȭ�Ϸ� ���Ǹ�, �� ���������� �������� ���� ��������(ConnString),���̺�(STable),
'����������(IPage)�� ������ ����.


'===================================================================================================

        ConnString = "Provider=SQLOLEDB;Data Source=(Local);Initial Catalog=LNP_COSMETICS;User ID=Sa;Password=1656;" 

'===================================================================================================


STable=Request.QueryString("Table")  '�Խ��� ��ũȭ�Ͽ��� URL�� ���۵� �� 

IF Stable="" then '���̺� ���� ���� ��� ����Ʈ ����
Stable="AL_021_Judge_System"
end if

SMode=Request.QueryString("Mode")    '��Ϻ���ȭ�Ͽ��� ������/���������� ���⿡�� ���۵� �� 
If (SMode = "") Then
    SMode = "qa"
End If

Var1 = "Table=" & STable

'�˻��� �ʵ�� �ܾ ���۹޴´�
Field = Request.QueryString("Field")
Str = Request.QueryString("Str")

Var1 = Var1 & "&Field=" & Field
Var1= Var1 & "&Str=" & Str


SPage = Request.QueryString("Page") '������ ������ �ʿ��� ��� ���Ͽ��� ���۵� ��
If (SPage = "") Then
    SPage = "1"
End If

IPage = CInt(SPage)              'SPage������ ���� INT������ ��ȯ

Var2 = Var1 & "&Page=" & SPage


'==========  ���� ���� ���  =================

Pagesize = 10 '����� ���ڵ尳��
Groupsize= 15 '����� ����������



adminpass="2241" '�����ڿ� ����������ȣ



Maxsize = 20 '�ִ���� ���ε� �뷮
'============================================


%>
