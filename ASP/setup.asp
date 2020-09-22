<!-- #include file = "database.asp" -->
<%

sDataConnPath = Server.MapPath("/ASP/database/data.mdb")
sUserConnPath = Server.MapPath("/ASP/database/users.mdb")

Sub WriteMeta
%>
<HTML>
<HEAD><TITLE>Test</TITLE>
<META NAME="keywords" CONTENT="my page, about, me">
<META NAME="description" CONTENT="Great Page, Come See">
<META NAME="generator" CONTENT="HTML Webmaster">
<META NAME="author" CONTENT="Your Name">
<META NAME="copyright" CONTENT="Date, Author, Company">
<META NAME="resource-type" CONTENT="document">
<META NAME="revisit-after" CONTENT="30 days">
<META NAME="classification" CONTENT="Internet">
<META NAME="distribution" CONTENT="Global">
<META NAME="rating" CONTENT="General">
<META NAME="language" CONTENT="en-us">
<META NAME="Robots" CONTENT="ALL">
<meta http-equiv="content-style-type" content="text/css">
<link rel="STYLESHEET" type="text/css" href="main.css">
</HEAD>
<%
End Sub

Sub WriteHeader
WriteMeta
%>
<BODY BACKGROUND="background.gif" BGCOLOR="#D3D0C4" Text="#000000" LINK="#61422F" VLINK="#876A58" ALINK="#D03720">
<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:61422F}  A:visited {text-decoration: none; color:876A58}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:D03720; }-->
</style>
<style>
BODY {SCROLLBAR-FACE-COLOR: #988D73; SCROLLBAR-HIGHLIGHT-COLOR: #D3D0C4; SCROLLBAR-SHADOW-COLOR: #000000; SCROLLBAR-3DLIGHT-COLOR: #000000; SCROLLBAR-ARROW-COLOR:  #000000; SCROLLBAR-TRACK-COLOR: #D3D0C4; SCROLLBAR-DARKSHADOW-COLOR: #D3D0C4; }
</style>
<FONT COlOR="#000000" FACE="Arial" SIZE="2">
<%
End Sub

Sub WriteFooter
%>
</Font>
</BODY>
</HTML>
<%
End Sub

%>