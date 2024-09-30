<!--#include file="../connection/constr.asp"-->
<%
'response.Write "hello"
'response.End
if session("aphone")="" or session("aphone")=" " or Len(Session("aphone"))=0 then 
    response.Write "<center><h2 style='color:red;'>Your session is expired!!! </h2></center>"
else

vars=split(request.QueryString("vars"),",")


sqlstr="update " & Session("CAMPAIGNID") & " set "

for each v in vars
      sqlstr = sqlstr & v & "='',"  
Next
sqlstr=mid(sqlstr,1,Len(sqlstr)-1)
sqlstr = sqlstr & "  ,Disp='HS' where aphone='" & Session("aphone") & "'"

response.Write sqlstr


	Set cnlast = Server.CreateObject("ADODB.Connection")
	Set rs1 = Server.CreateObject("ADODB.Recordset")
    Set rs2 = Server.CreateObject("ADODB.Recordset")
	if cnlast.State <> 1 then
		cnlast.Open connect2MR2009		
	end if
	
	if rs1.State="1" then rs1.Close()
	rs1.Open "Select QuesList, quesno from " & Session("CAMPAIGNID")& "pageinfo where aphone='"  & Session("aphone") & "'", cnlast
	response.Write rs1.source

	

	if Len(rs1.Fields(0).value)>0 then 
        str = split(rs1.Fields(0),",")
    else
        response.Write str
        'response.end
        'response.Redirect(rs1.Fields(1).value)
    end if
  	
   if Instr(1,request.ServerVariables("HTTP_REFERER"),"?")=0 then
                path=request.ServerVariables("HTTP_REFERER")
                campfolder=split(path,"/")
                lang = campfolder(Ubound(campfolder)-1)
                foldernm=campfolder(Ubound(campfolder)-2)
   else
                Qpath=split(request.ServerVariables("HTTP_REFERER"),"?")
                path=Qpath(1)
                campfolder=split(path,"/")              
                lang = campfolder(Ubound(campfolder)-1)
                foldernm=campfolder(Ubound(campfolder)-2)    
   end if
   
  'response.Write  Ubound(str)&"AA" 
  'response.end
   if ubound(str) > 0 then
        fullpath="/projects/" & foldernm & "/" & lang & "/" & str(Ubound(str)-1) & ".asp" 
   else
   'response.Write "AAAAAAAAAAAAAAA"
   'response.end
        response.Redirect(rs1.Fields(1).value)
   end if
   
   response.Write fullpath  & "Full path"
   'response.end
     
   fullpath=URLDecode(fullpath)

        if rs1.State="1" then rs1.Close()
        rs1.Open "update " & Session("CAMPAIGNID")& "pageinfo set QuesList=(Select replace(replace(QuesList,'"& str(Ubound(str)-1) & ",',''),',,',',') from " & Session("CAMPAIGNID") & "Pageinfo where aphone='" & Session("aphone") & "') where aphone='" & Session("aphone") & "'", cnlast     
  
if request.QueryString("vars")<>"" and Len(request.QueryString("vars"))>0 and request.QueryString("vars")<>" " then 
        if rs1.State="1" then rs1.Close()
        rs1.Open sqlstr, cnlast
end if
   
    
         if str(Ubound(str)-1)<>"main" and str(Ubound(str)-1)<>"authenticate.asp" and str(Ubound(str)-1)<>"Catipage" then 
         if rs1.State="1" then rs1.Close()
         response.Write "update " & Session("CAMPAIGNID")& "pageinfo set quesno='" & replace(Ucase(fullpath),Ucase(".asp.asp"),UCase(".asp")) & "' where aphone='" & Session("aphone") & "'"
         'response.end

    'For last page -Shivani
     lastno=str(Ubound(str)-1) & ".asp" 
    if rs2.State="1" then rs2.Close()
'End For last page -Shivani


         ' Edit by dinesh on 07-Oct-2013
         if (Ucase (right( replace(Ucase(fullpath),Ucase(".asp.asp"),UCase(".asp")), 5) )="/.ASP") then
            '
          else  
            rs1.Open "update " & Session("CAMPAIGNID")& "pageinfo set quesno='" & replace(Ucase(fullpath),Ucase(".asp.asp"),UCase(".asp")) & "' where aphone='" & Session("aphone") & "'", cnlast  
    'For last page -Shivani
            rs2.open "update " & Session("CAMPAIGNID")& " set lastQ='"& Replace(Ucase(lastno),Ucase(".asp.asp"),UCase(".asp")) &"' where aphone='" & Session("aphone") & "'", cnlast       
            
    'End For last page -Shivani
         end if
      end if

      if Instr(1,Ucase(request.ServerVariables("HTTP_REFERER")),"/PROJECTS/")=0 then     
        fullpath=replace(Ucase(fullpath),Ucase("/PROJECTS/"),UCase("/"))   
        fullpath=replace(Ucase(fullpath),Ucase(".asp.asp"),UCase(".asp"))        
    else     
        fullpath=replace(Ucase(fullpath),Ucase(".asp.asp"),UCase(".asp"))    
    end if
   response.redirect(fullpath)
end if
    
    
    
    Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    URLDecode = sOutput
End Function
%>