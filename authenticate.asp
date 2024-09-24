<!--#include file="../connection/constr.asp"-->
<!--#include file="GeneralUpdate.asp"-->

<%
session("flag")=false
Session("DMStatus")= "Dialer"
Session("repeatflag")=""

    session("TeST")=""

if left(Request.Form("APHONE"),2)="00" then
	Session("aphone")=right(trim(Request.form("APHONE")),len(trim(Request.form("APHONE"))-2))
elseif left(Request.Form("APHONE"),1)="0" then
	Session("aphone")=right(trim(Request.form("APHONE")),len(trim(Request.form("APHONE"))-1))
else
	Session("aphone")=trim(Request.form("APHONE"))
end if	
session("CAMPAIGNID") = "AZ275_USTC_0924"

'' --------This code is to display the ENDPAGE message to the user without updating saledate field
EndP=getanyfieldvalue("AZ275_USTC_0924Pageinfo","Aphone",Session("aphone"),"Quesno","","")	
DISP1=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"DISP","","")
Dat=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"Saledate","","")
Doneby1=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"AgentId","","")

'last page code -Shivani

  EndPg1=""
    EndPg=getanyfieldvalue("AZ275_USTC_0924Pageinfo","Aphone",Session("aphone"),"quesno","","")&""	
    
    if EndPg="," or EndPg="" then        
        Session("newpage")= "S1.asp"   
    else       
            a=split(EndPg,"/")
            b=ubound(a)
             For i=0 to b				
                 Session("newpage")=a(i)
			next
    end if
'last page code end -Shivani
If  UCASE(DISP1)<>"HS" then

 %>
 
    <link href="../Style/css/bootstrap.css" rel="stylesheet" />
    <link href="../Style/css/custome.css" rel="stylesheet" />
  

    <!-- for fancybox -->

<section class="container">
           

<div class="completeProgress">
    <span class="progress-icon glyphicon glyphicon-ok" style="border-color: #f7981c; color: #f7981c;"></span>
</div>
<style>
    .logo {
        display: none;
    }
    body {
        background:none;
        background-color: #f5f5f5;
    }
</style>
<hr style="width: 25%;" />
<%
If  UCASE(DISP1)="Q" then

     	response.write "<h4 style='font-size: 30px; text-align: center;'>Sorry! this Survey is already Screen-out on <font color=green>" & day(Dat) & " " & MonthName(month(Dat)) & " " & Year(Dat) & "  </font></h4></center>"
        session("disposition")=""
        session("Aphone")=""
        response.end
end if
If  UCASE(DISP1)="S_OQ" then
     	
        response.write "<h4 style='font-size: 30px; text-align: center;'>Sorry! this Survey was already completed.But it is in OverQuota.</h4></center>"
        session("disposition")=""
        session("Aphone")=""
        response.End
end if
If  UCASE(DISP1)="S_QC" then
     	response.write "<h4 style='font-size: 30px; text-align: center;'>Sorry! this Survey was already completed.But Rejected  on <font color=green>" & day(Dat) & " " & MonthName(month(Dat)) & " " & Year(Dat) & "  </font></h4></center>"
        session("disposition")=""
        session("Aphone")=""
        response.End
end if
If  UCASE(DISP1)="S" then 
    	response.write "<h4 style='font-size: 30px; text-align: center;'>Sorry! this Survey was already completed on <font color=green>" & day(Dat) & " " & MonthName(month(Dat)) & " " & Year(Dat) & "  </font></h4></center>"
        session("disposition")=""
        session("Aphone")=""
        response.End
end if
If  Instr(1,Ucase(DISP1),"S_")>= 1 then
    response.write "<h4 style='font-size: 30px; text-align: center;'>Sorry! this Survey was already completed.But Rejected  on <font color=green>" & day(Dat) & " " & MonthName(month(Dat)) & " " & Year(Dat) & "  </font></h4></center>"
     session("disposition")=""
     session("Aphone")=""
	response.end
end if	
    %>
                   
    
    </section>
<%
   
   End if  

	'' Calculate Duration. 	
    attempt1=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"attempt1","","")	
    pdate=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"saledate","","")	
    
		if attempt1 <>"" then
		attempt2=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"attempt2","","")	
			if attempt2 <> "" then
				attempt3=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"attempt3","","")	
				if attempt3 <> "" then  
					call UpdateAZ275_USTC_0924("attempt4",now(),"")
					call UpdateAZ275_USTC_0924("attempt3e",pdate,"")
				else
					call UpdateAZ275_USTC_0924("attempt3",now(),"")
					call UpdateAZ275_USTC_0924("attempt2e",pdate,"")
				end if
			else
					call UpdateAZ275_USTC_0924("attempt2",now(),"")
					call UpdateAZ275_USTC_0924("attempt1e",pdate,"")
			end if
		else
			call UpdateAZ275_USTC_0924("attempt1",now(),"")
		end if

StartDateTime=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"StartDateTime","","")	
if StartDateTime = "" or StartDateTime = " " or isnull(StartDateTime)=true then
	call UpdateAZ275_USTC_0924("StartDateTime",now(),"")
'else
'	SaleDate=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"SaleDate","","")	
'	Duration=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"Duration","","")
'		
'	if Duration="" or len(Duration) =  0 or isnull(Duration) then Duration=0
'	response.Write DateDiff("s",StartDateTime,SaleDate) 
'	
'	secnd1=DateDiff("s",StartDateTime,SaleDate) 
'	if secndl=0 or secndl="" or secndl= " " then secndl=0 
'	response.write  Duration & "-" & secndl
'	Duration=cdbl(Duration) + cdbl(secnd1)
'	response.Write Duration
'	call UpdateAZ275_USTC_0924("Duration",Duration,"")
'	call UpdateAZ275_USTC_0924("StartDateTime",now(),"")
end if
'' Calculate Duration.		


'--Random Unique Identifier--
 dim con, rs,rsCNT,rsCallback,rsChk
		set con= Server.Createobject("Adodb.Connection")
		set rs = server.Createobject("Adodb.Recordset")
        set rsCNT = server.Createobject("Adodb.Recordset")
        set rsCallback = server.Createobject("Adodb.Recordset")
        set rsChk = server.Createobject("Adodb.Recordset")
	   con.Open   connect2MR2009 

rsCallback.open "Select Uni_Id from AZ275_USTC_0924 where aphone='"& Session("aphone") &"' ",con
if isNull(rsCallback.Fields("Uni_Id")) or rsCallback.Fields("Uni_Id")="" then

rs.open "SELECT SUBSTRING(CONVERT(varchar(40), NEWID()),0,19) as Uni_Id",con
            rsCNT.open "select count(*) as 'cnt' from AZ275_USTC_0924",con
	        Uni_Id=rs.Fields("Uni_Id")&"-"&rsCNT.Fields("cnt")
            
            rsChk.open "SELECT * from AZ275_USTC_0924 where Uni_Id='" & Uni_Id &"'",con

           
			if not rsChk.EOF then
                Uni_Id=rs.Fields("Uni_Id")&"-"&rsCNT.Fields("cnt")&"-"&rsCNT.Fields("cnt")
                call UpdateAZ275_USTC_0924("Uni_Id",Uni_Id,"")	
            else
                call UpdateAZ275_USTC_0924("Uni_Id",Uni_Id,"")	
            end if
end if


'--Random Unique Identifier--

'--Ip Tracker--
set cn=server.createobject("adodb.connection")
Set cmd = Server.CreateObject("Adodb.command")
cn.open connect2MR2009
cmd.ActiveConnection=cn
	
dim tno,sqlcnt
tno=Session("Aphone")
ip =  request.ServerVariables("remote_addr")
browser = Request.ServerVariables("HTTP_USER_AGENT")
SSQL  = "insert into catiiptracker(campname,aphone,ip,browser,surveydate) values ('" & session("CAMPAIGNID")  & "','" & tno & "','" & ip & "','" & browser & "','" & now() & "')" 
cmd.CommandText=ssql
cmd.Execute
'--Ip Tracker--




'' --------This code is to display the ENDPAGE message to the user without updating saledate field - 9-may-06	

'' --------This code is to display the ENDPAGE message to the user without updating saledate field - 9-may-06	

'' Followig Query to update 11 updation statement into 1 update query.
UpdateQuery="ACOMPANY='" & Replace(Request.form("ACOMPANY"),"'","`") &"', AFNAME ='" & Replace(Request.form("AFNAME"),"'","`") &"' , ALNAME ='" & Replace(Request.form("ALNAME"),"'","`") &"'  ,EMAIL='" & Request.form("AEMAIL") &"', FRONTID='" & Request.form("AGENTID") &"', AGENTID='" & Request.form("AGENTID") &"', SHIFT='" & Request.form("SHIFT") &"', COUNTRY='" & Request.form("COUNTRY") &"', CountryCode='" & Request.form("CountryCode") &"', CountryforReport='" & Request.form("CountryforReport") &"',sessioncampname='" & session("CAMPAIGNID") &"'" 
'' Followig Query to update 11 updation statement into 1 update query.
call UpdateAZ275_USTC_0924(UpdateQuery,"","T")

UpdateQuery="JOBTITLE='"& Replace(Request.form("JOBTITLE"),"'","`") &"', Industry='"& Replace(Request.form("Industry"),"'","`") &"',ccomment='"& Replace(Request.form("comment"),"'","`") &"',respid='"& Request.form("respid") &"',phone_type='"& Request.form("phone_type") &"',panelmember='"& Request.form("panelmember") &"', Device_Type ='" & Request.form("Device_Type") &"'"  
'' Followig Query to update 11 updation statement into 1 update query.
call UpdateAZ275_USTC_0924(UpdateQuery,"","T")

    UpdateQuery="W_Revenue='" & Request.form("W_Revenue") &"', W_Revenue_Range ='" & Request.form("W_Revenue_Range") &"',ip ='" & ip & "', L_Revenue ='" & Request.form("L_Revenue") &"' ,L_Revenue_Range='" & Request.form("L_Revenue_Range") &"',Data_Group='"& Request.form("Data_Group") &"',SampleId='" & Session("SampleId") &"',test='"&  Session("itest") &"',partnermemberDate='" & now() & "'" 
'' Followig Query to update 11 updation statement into 1 update query.
call UpdateAZ275_USTC_0924(UpdateQuery,"","T")

LeadId=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"LeadId","","")&""
if LeadId="" or LeadId=" " then
        call UpdateAZ275_USTC_0924("LeadId",Session("LeadId"),"")
else
        exst=InStr(1, LeadId, Session("LeadId"))
        if exst<1 then
            call UpdateAZ275_USTC_0924("LeadId",LeadId&"|"&Session("LeadId"),"")
        end if
end if


''--------- email checkbox update start---------
if Request.Form("chkAEMAIL")="0" then
    call UpdateAZ275_USTC_0924("EMAIL","REFUSED","")
else
    call UpdateAZ275_USTC_0924("EMAIL",Request.form("AEMAIL"),"")
end if
''--------- email checkbox update end-----------

Phone=Session("aphone")

    EndP=getanyfieldvalue("AZ275_USTC_0924","Aphone",Session("aphone"),"count(*)","Disp='Q'","")	
	if EndP ="1" or Endp >="1" then
		response.Redirect "Endpage.asp"
	end if

	Set cn_auth = Server.CreateObject("ADODB.Connection")
	Set rs_auth = Server.CreateObject("ADODB.Recordset")	
	cn_auth.Open connect2MR2009
	rs_auth.open "select * from AZ275_USTC_0924PAGEINFO where aphone='" & Session("aphone") & "'" ,cn_auth
	if not rs_auth.eof then
		Navipage=rs_auth("quesno")
		if instr(1,lcase(rs_auth("quesno")),"authenticate.asp")>0 then
			navi= "S1.asp"
		end if
	else
			'Response.redirect "Main.asp"
			navi= "S1.asp"
	end if
	
	if Navipage="" or isnull(Navipage) then
'		Response.redirect "Main.asp"
			navi= "S1.asp"
	else
		Navi=replace(Ucase(Navipage),"/PROJECTS/","")
		Navi=replace(Ucase(Navi),"PROJECTS/","")
		Navi=replace(Ucase(Navi),"/AZ275_USTC_0924_O/","")
		Navi=replace(Ucase(Navi),"AZ275_USTC_0924_O/","")
		Navi=replace(Ucase(Navi),"/AZ275_USTC_0924/","")
		Navi=replace(Ucase(Navi),"AZ275_USTC_0924/","")
		Navi=replace(Ucase(Navi),"/BENGALI/","")
		Navi=replace(Ucase(Navi),"BENGALI/","")
		Navi=replace(Ucase(Navi),"BENGALI","")
		
		Navi=replace(Ucase(Navi),"/ENG/","")
		Navi=replace(Ucase(Navi),"ENG/","")
		Navi=replace(Ucase(Navi),"ENG","")
		Navi=replace(Ucase(Navi),"/HINDI/","")
		Navi=replace(Ucase(Navi),"HINDI/","")
		Navi=replace(Ucase(Navi),"HINDI","")

		Navi=replace(Ucase(Navi),"/GUJARATI/","")
		Navi=replace(Ucase(Navi),"GUJARATI/","")
		Navi=replace(Ucase(Navi),"GUJARATI","")
		Navi=replace(Ucase(Navi),"/KANNADA/","")
		Navi=replace(Ucase(Navi),"KANNADA/","")
		Navi=replace(Ucase(Navi),"KANNADA","")
		Navi=replace(Ucase(Navi),"/MALYALAM/","")
		Navi=replace(Ucase(Navi),"MALYALAM/","")
		Navi=replace(Ucase(Navi),"MALYALAM","")
		Navi=replace(Ucase(Navi),"/MARATHI/","")
		Navi=replace(Ucase(Navi),"MARATHI/","")
		Navi=replace(Ucase(Navi),"MARATHI","")
		Navi=replace(Ucase(Navi),"/ORIYA/","")
		Navi=replace(Ucase(Navi),"ORIYA/","")
		Navi=replace(Ucase(Navi),"ORIYA","")
		Navi=replace(Ucase(Navi),"/TAMIL/","")
		Navi=replace(Ucase(Navi),"TAMIL/","")
		Navi=replace(Ucase(Navi),"TAMIL","")
		Navi=replace(Ucase(Navi),"/TELUGU/","")
		Navi=replace(Ucase(Navi),"TELUGU/","")
		Navi=replace(Ucase(Navi),"TELUGU","")
        Navi=replace(Ucase(Navi),"/TELEGU/","")
		Navi=replace(Ucase(Navi),"TELEGU/","")
		Navi=replace(Ucase(Navi),"TELEGU","")
		Navi=replace(Ucase(Navi),"/PUNJABI/","")
		Navi=replace(Ucase(Navi),"PUNJABI/","")
		Navi=replace(Ucase(Navi),"PUNJABI","")
		Navi=replace(Ucase(Navi),"..\\","")
		session ("Navi")=Navi
	end if
	Session("Page")= Navi
%>
<html>
	<head>
		<title>New Page 2</title>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
		<meta name="ProgId" content="FrontPage.Editor.Document">
	</head>
	<body>
		<form name="frm" action="MainPage.asp" method="get" ID="Form1">
			<input type="hidden" name="navi" ID="Hidden1"> <input type="hidden" name="aphone" ID="Hidden2">
		</form>
		<script language="javascript">
					document.frm.navi.value="CatiPage.asp";
					document.frm.aphone.value="<%=session("aphone")%>";
					document.frm.submit();
		</script>
		
		
		<form name="frm1" ID="Form2">
			<script language="javascript">
	function ClickMe()
		{
			document.frm1.b1.disabled=true;
			document.frm1.action="InsertField.asp";
			document.frm1.submit();
		}
			</script>

		</form>
	</body>
</html>