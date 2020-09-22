<!-- #include file = "setup.asp" -->
<% 
WriteHeader

%>
ASP is working!!!<BR><BR>
<% 

OpenCon sUserConnPath
ChooseTable "Users"
%>
<TABLE BORDER="1" CELLPADDING="6" CELLSPACING="0" BGCOLOR="#FFFFFF" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0">
<TR>
<TD ALIGN="Left"><Font COLOR="#000000" SIZE="2"><B>Username</B></FONT></TD>
<TD ALIGN="Left"><Font COLOR="#000000" SIZE="2"><B>Password</B></FONT></TD>
<TD ALIGN="Left"><Font COLOR="#000000" SIZE="2"><B>Access</B></FONT></TD>
</TR>
<%
if isrecordsetempty = false then
    do until rs.eof
    %>
<TR>
<TD ALIGN="Left"><Font COLOR="#000000" SIZE="2"><% Response.Write(rs.fields("Username").value) %></FONT></TD>
<TD ALIGN="Left"><Font COLOR="#000000" SIZE="2"><% Response.Write(rs.fields("Password").value) %></FONT></TD>
<TD ALIGN="Left"><Font COLOR="#000000" SIZE="2"><% Response.Write(rs.fields("AccessType").value) %></FONT></TD>
</TR>
    <%
    rs.movenext
    loop
end if
%>
</TABLE>
<%
CloseCon
WriteFooter
%>