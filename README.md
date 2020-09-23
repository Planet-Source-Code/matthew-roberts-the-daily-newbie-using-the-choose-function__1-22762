<div align="center">

## The Daily Newbie \- Using the Choose\(\) Function


</div>

### Description

04/28/2001 - Describes the usage of the Choose() function; A really neat but seldom used command in VB.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-the-daily-newbie-using-the-choose-function__1-22762/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type"
content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">
<title>Daily Newbie - 04/28/2001</title>
</head>
<body bgcolor="#FFFFFF">
<p> </p>
<p class="MsoTitle"><img width="100%" height="3"
v:shapes="_x0000_s1027"></p>
<p align="center" class="MsoTitle"><font size="7"><strong>The
Daily Newbie</strong></font></p>
<p align="center" class="MsoTitle"><strong>&#8220;To Start Things
Off Right&#8221;</strong></p>
<p align="center" class="MsoTitle"><font size="1">Fourth
Edition                   
                                     
April 28,
2001                      
                                                  
Free</font></p>
<p align="center" class="MsoTitle"><img width="100%" height="3"
v:shapes="_x0000_s1027"></p>
<p align="center" class="MsoNormal" style="text-align:center"> </p>
<p align="center" class="MsoNormal" style="text-align:center"> </p>
<p class="MsoNormal"><font face="Arial"><strong>About this
feature:</strong></font></p>
<p class="MsoBodyText"><font size="2" face="Arial">
The initial plan for the Daily Newbie was to cover each function VB has to offer
in alphabetical order. I have now modified this plan slightly to skip over some of
the more advanced (or tedious) commands that I don't think the Newbie would benefit from.
Thanks again all who have written in support of this effort. It makes a difference.</font></p>
<p class="MsoNormal">Today's command is not widely known for some reason, but is faily useful.
I have been guilty of writing functions that do the exact same thing several times. I think you will
like this one.<font size="2" face="Arial"></font></p>
<p class="MsoNormal"><font size="2" face="Arial"></font></p>
<p class="MsoNormal" style="margin-left:135.0pt;text-indent:-135.0pt"><font size="2"
face="Arial"><strong>Today&#8217;s Keyword:</strong>
                             </font><font
size="4" face="Arial"> Choose()</font></p>
<p class="MsoNormal"
<font size="2"
face="Arial"><strong>Name Derived
From:        </strong>          </font></p>
<blockquote>
  <p class="MsoNormal"><font
  size="2" face="Arial"><strong>Choose</strong> (of
  course) &#8211; &#8220;(1) : to make a selection"
						 - <em><a href="http://www.webster.com/">Webster's online
  dictionary.</a></em></font></p>
  </blockquote>
 </blockquote>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Used for   </strong>                               
Making a choice between several possible options.</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>VB Help Description: </strong>            Selects and returns a value from a list ofarguments.
</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Plain
English:    </strong>                       Returns the option associated with the value passed it (I will just have to show you!)</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Syntax:       </strong>                              Choose(index, Choice1, Choice2, etc...)</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Usage:       </strong>                               strDecision = Choose(intChoice ,  "Just do it" ,  "Don't do it" ,  "It's your life" )
 </font></p>
<p class="MsoNormal"
style="margin-left:135.35pt;text-indent:-135.35pt"><font size="2"
face="Arial"><strong>Copy & Paste Code:</strong></font></p>
<br>
<br>
Today's code snippet will prompt for a month number and return a string
that corresponds to it.
<br>
<br>
<pre>
 Dim Choice
 Dim strMonth As String
  Do
  Choice = Val(InputBox("Enter a Number (1-12):"))
  If Choice + 0 = 0 Then Exit Do
  strMonth = Choose(Choice, "Jan", "Feb", "Mar", "Apr", _
    "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
  MsgBox strMonth
  Loop
</pre>
 <p class="MsoNormal"
 style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"> </p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Notes: </strong></p>
I really like the Choose function. It comes in handy for with evaluating which option button
was clicked, or anything else that returns an index number. Unfortunatly, like the Array() command
covered a couple of articles ago, the Choose() function requires a seperate hard coded value for
each possible choice. This isn't neccesarily a bad thing, but I am allergic to hard coding, so it
just rubs me wrong. I guess the chances of the order of the months changing is pretty slim...
<br><br>
<font size="2" face="Arial"><strong>Things to watch out for: </strong></font></p>
<li>Although the Choose() Statement only returns a single value, it still evaluates each one. In effect,
it acts like a compact series of If...Then statements. This can result in the sometimes baffling behavior
of displaying one message box with the correct value and many empty ones. For this reason, the results
of a Choose() statement should be returned to a variable before displaying it in a message box.
<br><br>
<li>If the Index value passed in is null, an error will result. Therefore, if you are using a variant
as your Index, you should add zero to it to initialize it as a number. This will make the default value zero, not null.
<br><br>
Tomorrow's Keyword:			Chr()
</font></p>
</body>
</html>

