<div align="center">

## \_\-\~Who is looking at your site? v1\.3\~\-\_


</div>

### Description

Allows you to see the user sessions and pages that each user is viewing on your website currently. Updated so now it works, but just hides the application variables. PLEASE VOTE FOR THIS SCRIPT! Kudos to Brian Battles, check out his scripts. Plus, you have to love statistics... Otherwise don't bother :) Its a pain but if you need to know stats (like me) its worth it.
 
### More Info
 
Must be able to use a global.asa. Almost impossible to use without. Be sure to set your session timeout to 10 minutes or so if at all possible. You will see a bunch of people that aren't there if you don't.

Good idea to restart your global.asa file as often as possible.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[JSBeckerist](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jsbeckerist.md)
**Level**          |Intermediate
**User Rating**    |3.9 (47 globes from 12 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jsbeckerist-who-is-looking-at-your-site-v1-3__4-7634/archive/master.zip)

### API Declarations

Go to Beckerist.com for a working model. Use this code freely... PLEASE VOTE FOR THIS CODE!


### Source Code

```
<!--This goes into a page which you include on ALL PAGES IN YOUR SITE... This is how you post things so that you can list them-->
<%APPLICATION.LOCK
sSessURL= Request.ServerVariables ("URL")
sQString= Request.ServerVariables ("QUERY_STRING")
if sQString <>"" then
sSessURL=sSessURL&"?"&sQString
end if
if session("name1")<>"" then
	Application("yoyo"&Session.SessionID) = Null
	sSessID=session("name1")
else
	sSessID = Session.SessionID
end if
Application("yoyo"&sSessID) = sSessURL
APPLICATION.UNLOCK
%>
<!--~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~New Page~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~-->
<!--This is the page that you want to output the users on, you can add more variables as it goes on-->
<h2 align="center">Yourdomain.com Stats - Who Is Online?</h2>
<BR>
<!--#include virtual="include_top_page.asp"--><center>
<%if request.form("name1")<>"" then
session("name1")=request.form("name1")
elseif session("name1")<>"" then%>
Hello <%=session("name1")%>!
<%else%>
<form method="Post" action="">
Give yourself a name:
<input type="text" name="name1"><input type="submit" value="Give me the name!"></form>
<%end if%>
<table>
<tr><td width="50%">Session ID</td><td>Looking where?</td></tr>
<%
 For Each strSV in Application.Contents
 		if left(strSV,4)="yoyo" then
if Application.Contents(strSV)<>"" then%>
 <tr><td><%=mid(strSV,5,999)%></td><td><%=Application.Contents(strSV)%></td></tr>
<%end if%><% end if
 Next%></table>New Page~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~-->
<!-- This next part goes in the session_onend sub in your global.asa file-->
Sub Session_OnEnd
	Application.Lock
		sSessID = Session.SessionID
		Application("yoyo"&sSessID) = ""
	Application.Remove "yoyo"&sSessID
	Application.Unlock
End Sub
```

