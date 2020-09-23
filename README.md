<div align="center">

## Screensaver Tutorial \*UPDATED\* \(Again\)

<img src="PIC20024191922199982.jpg">
</div>

### Description

This article is for you who struggle to find out how to create a screensaver with vb. This article is an update to my last one. This one includes a working source code example. I have updated this yet again. This one ends on mouse movements. Please vote for me.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-04-21 09:17:24
**By**             |[Person82484536536578579365897](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/person82484536536578579365897.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Screensave741784212002\.zip](https://github.com/Planet-Source-Code/person82484536536578579365897-screensaver-tutorial-updated-again__1-33958/archive/master.zip)





### Source Code

```
<style>
<!--
 /* Font Definitions */
@font-face
	{font-family:"Comic Sans MS";
	panose-1:3 15 7 2 3 3 2 2 2 4;
	mso-font-charset:0;
	mso-generic-font-family:script;
	mso-font-pitch:variable;
	mso-font-signature:647 0 0 0 159 0;}
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";}
h1
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:16.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;
	mso-font-kerning:0pt;}
p.MsoTitle, li.MsoTitle, div.MsoTitle
	{margin:0in;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	font-size:26.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Comic Sans MS";
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";
	font-weight:bold;}
p.MsoBodyText, li.MsoBodyText, div.MsoBodyText
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:18.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Comic Sans MS";
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";
	font-weight:bold;}
p.MsoBodyText2, li.MsoBodyText2, div.MsoBodyText2
	{margin:0in;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	font-size:20.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Comic Sans MS";
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";
	font-weight:bold;}
p.MsoBodyText3, li.MsoBodyText3, div.MsoBodyText3
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:16.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";
	font-weight:bold;}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 9.0pt 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-CA style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoTitle>Screensaver Tutorial</p>
<p class=MsoBodyText>In this tutorial you will learn how to create a
screensaver with visual basic.<span style="mso-spacerun: yes">  </span>It is
fairly easy once you study it so now lets get into the project.</p>
<p class=MsoNormal><b><span style='font-size:18.0pt;mso-bidi-font-size:12.0pt;
font-family:"Comic Sans MS"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:18.0pt;mso-bidi-font-size:12.0pt;
font-family:"Comic Sans MS"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoBodyText2>We will create a screensaver today</p>
<p class=MsoNormal><b><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt;
font-family:"Comic Sans MS"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt;
mso-bidi-font-family:Arial'>To start off, start a new project in visual basic.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt;
mso-bidi-font-family:Arial'>Name the project “Blanker”.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt;
mso-bidi-font-family:Arial'>Add four forms and name one ‘frmMain’, one
‘frmSettings’, one ‘frmPassSetup’ and one ‘frmPassword’. Also add a new module.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt;
mso-bidi-font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt;
mso-bidi-font-family:Arial'>Click the menu project|Blanker Properties.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt;
mso-bidi-font-family:Arial'>In the ‘General’ tab, change the ‘Startup Object’
to Sub Main.<o:p></o:p></span></p>
<h1><span style='font-weight:normal'>Choose the ‘Make’ tab and in the title
text box, type “SCRNSAVER Blanker”<o:p></o:p></span></h1>
<p class=MsoNormal><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt'>Now
click ok and you’re ready to start coding!<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:16.0pt;mso-bidi-font-size:12.0pt'>In
the module you created, copy these API’s into it:<o:p></o:p></span></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Option
Explicit 'All variables must be declared<o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>'Constants<o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Public
Const SW_SHOWNORMAL = 1<o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Const APP_NAME = &quot;Blanker&quot;<o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>‘API’s<o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>‘This
Function Shows and hides the cursor.<o:p></o:p></span></b></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Declare
Function ShowCursor Lib &quot;user32&quot; (ByVal bShow<span
style="mso-spacerun: yes">  </span>As Long) As Long<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>‘This
function finds if another instance of the saver is running.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Declare
Function FindWindow Lib &quot;user32&quot; Alias
&quot;FindWindowA&quot;(ByVallpClassName As String, ByVal lpWindowName As
String) As Long<o:p></o:p></span></p>
<b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Arial;
mso-fareast-font-family:"Times New Roman";mso-bidi-font-family:"Times New Roman";
mso-ansi-language:EN-CA;mso-fareast-language:EN-US;mso-bidi-language:AR-SA'><br
clear=all style='page-break-before:always'>
</span></b>
<p class=MsoBodyText3><span style='font-weight:normal'>Now copy the main sub:<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Sub
Main ()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>‘This
sub is called when windows wants the screensaver to run a certain event.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>‘This
locates what windows wants and activates it.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Select
Case Mid(UCase$(Trim$(Command$)), 1, 2)<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Case
&quot;/C&quot; 'Configurations mode called<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmSettings.Show
1<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Case
&quot;&quot;, &quot;/S&quot; 'Screensaver mode<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>runScreensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Case
&quot;/A&quot; 'Password protect dialog<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmPassSetup.Show
1<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Case
&quot;/P&quot; 'Preview mode<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>‘The
preview mode is very advanced. It is when you see a clip of it on the little
‘monitor. Just leave the monitor screen blank.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Select<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now copy the other
functions that go into the module:<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub runScreensaver() 'Run the screen saver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>checkInstance
'Make sure no other instances are running<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>ShowCursor
False 'Disable cursor<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>'load
Screen Saver's main form<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Load
frmMain<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmMain.Show<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub checkInstance()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>'If
no previous instance is running, exit sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
Not App.PrevInstance Then Exit Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>'check
for another instance of screen saver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
FindWindow(vbNullString, APP_NAME) Then Exit Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>'Set
our caption so other instances can find<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>'us
in the previous line.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmMain.Caption
= APP_NAME<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Sub
exitScreensaver() 'Exit the screensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>ShowCursor
True ‘Show the cursor<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now for the Settings
form.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Set the form’s
properties as follows:<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid teal 1.5pt;mso-padding-alt:0in 5.4pt 0in 5.4pt'>
 <tr>
 <td width=295 valign=top style='width:221.4pt;border-top:solid teal 1.5pt;
 border-left:solid teal 1.5pt;border-bottom:solid aqua .75pt;border-right:
 none;background:black;mso-shading:white;mso-pattern:solid black;padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><i><span style='color:white'>Property<o:p></o:p></span></i></p>
 </td>
 <td width=295 valign=top style='width:221.4pt;border-top:solid teal 1.5pt;
 border-left:none;border-bottom:solid aqua .75pt;border-right:solid teal 1.5pt;
 background:black;mso-shading:white;mso-pattern:solid black;padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><i><span style='color:white'>Setting<o:p></o:p></span></i></p>
 </td>
 </tr>
 <tr>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 solid teal 1.5pt;border-bottom:solid aqua .75pt;border-right:none;mso-border-top-alt:
 solid aqua .75pt;background:navy;mso-shading:white;mso-pattern:solid navy;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><i><span style='color:white'>Border Style<o:p></o:p></span></i></p>
 </td>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 none;border-bottom:solid aqua .75pt;border-right:solid teal 1.5pt;mso-border-top-alt:
 solid aqua .75pt;background:teal;mso-shading:white;mso-pattern:solid teal;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white;font-weight:normal'>4 –
 FixedToolWindow<o:p></o:p></span></p>
 </td>
 </tr>
 <tr>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 solid teal 1.5pt;border-bottom:solid aqua .75pt;border-right:none;mso-border-top-alt:
 solid aqua .75pt;background:navy;mso-shading:white;mso-pattern:solid navy;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white'>Caption<o:p></o:p></span></p>
 </td>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 none;border-bottom:solid aqua .75pt;border-right:solid teal 1.5pt;mso-border-top-alt:
 solid aqua .75pt;background:teal;mso-shading:white;mso-pattern:solid teal;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white;font-weight:normal'>About<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoBodyText3><span style='font-weight:normal'>Now just add a label and
set it’s caption to “Created By: “ and your name or something.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>That form had no code.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now let’s move on to the
password setup form.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Set the caption to
“Password” and the border style to 4.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now add a label that
says “Password:” and add a text box beside it named txtPassword. Now add two
command buttons with one named “cmdOK” and one named “cmdCancel”. Set the
captions to ‘OK’ and ‘Cancel’.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now copy this code:<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub cmdCancel_Click()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Unload
Me<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub cmdOK_Click()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>SaveSetting
&quot;Blanker&quot;, &quot;Settings&quot;, &quot;Password&quot;,
txtPassword.Text<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Unload
Me<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now let’s get into the
other password form.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Set the form’s settings
as the same as the other password’s form’s settings.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now add the same
controls as the password setup form with the same properties.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now just enter this
code:<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub cmdOK_Click()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
txtPassword.Text = GetSetting(&quot;Blanker&quot;, &quot;Settings&quot;,
&quot;Password&quot;, &quot;&quot;) Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Unload
Me<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Else<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>txtPassword.Text
= &quot;&quot;<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Me.Hide<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>MsgBox
&quot;Wrong password! Try again!&quot;, vbOKOnly + vbCritical,
&quot;Password&quot;<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Unload
Me<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub cmdCancel_Click()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Unload
Me<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Now for the main form:<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-weight:normal'>Set the form’s
properties as so:<o:p></o:p></span></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid teal 1.5pt;mso-padding-alt:0in 5.4pt 0in 5.4pt'>
 <tr>
 <td width=295 valign=top style='width:221.4pt;border-top:solid teal 1.5pt;
 border-left:solid teal 1.5pt;border-bottom:solid aqua .75pt;border-right:
 none;background:black;mso-shading:white;mso-pattern:solid black;padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><i><span style='color:white'>Property<o:p></o:p></span></i></p>
 </td>
 <td width=295 valign=top style='width:221.4pt;border-top:solid teal 1.5pt;
 border-left:none;border-bottom:solid aqua .75pt;border-right:solid teal 1.5pt;
 background:black;mso-shading:white;mso-pattern:solid black;padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><i><span style='color:white'>Setting<o:p></o:p></span></i></p>
 </td>
 </tr>
 <tr>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 solid teal 1.5pt;border-bottom:solid aqua .75pt;border-right:none;mso-border-top-alt:
 solid aqua .75pt;background:navy;mso-shading:white;mso-pattern:solid navy;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><i><span style='color:white'>BackColor<o:p></o:p></span></i></p>
 </td>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 none;border-bottom:solid aqua .75pt;border-right:solid teal 1.5pt;mso-border-top-alt:
 solid aqua .75pt;background:teal;mso-shading:white;mso-pattern:solid teal;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white;font-weight:normal'>&amp;H00000000&amp;<o:p></o:p></span></p>
 </td>
 </tr>
 <tr>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 solid teal 1.5pt;border-bottom:solid aqua .75pt;border-right:none;mso-border-top-alt:
 solid aqua .75pt;background:navy;mso-shading:white;mso-pattern:solid navy;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white'>BorderStyle<o:p></o:p></span></p>
 </td>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 none;border-bottom:solid aqua .75pt;border-right:solid teal 1.5pt;mso-border-top-alt:
 solid aqua .75pt;background:teal;mso-shading:white;mso-pattern:solid teal;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white;font-weight:normal'>0 – None<o:p></o:p></span></p>
 </td>
 </tr>
 <tr>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 solid teal 1.5pt;border-bottom:solid aqua .75pt;border-right:none;mso-border-top-alt:
 solid aqua .75pt;background:navy;mso-shading:white;mso-pattern:solid navy;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white'>Caption<o:p></o:p></span></p>
 </td>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 none;border-bottom:solid aqua .75pt;border-right:solid teal 1.5pt;mso-border-top-alt:
 solid aqua .75pt;background:teal;mso-shading:white;mso-pattern:solid teal;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white;font-weight:normal'>Blanker<o:p></o:p></span></p>
 </td>
 </tr>
 <tr>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 solid teal 1.5pt;border-bottom:solid teal 1.5pt;border-right:none;mso-border-top-alt:
 solid aqua .75pt;background:navy;mso-shading:white;mso-pattern:solid navy;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white'>WindowState<o:p></o:p></span></p>
 </td>
 <td width=295 valign=top style='width:221.4pt;border-top:none;border-left:
 none;border-bottom:solid teal 1.5pt;border-right:solid teal 1.5pt;mso-border-top-alt:
 solid aqua .75pt;background:teal;mso-shading:white;mso-pattern:solid teal;
 padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoBodyText3><span style='color:white;font-weight:normal'>2 -
 Maximized<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoBodyText3><span style='font-weight:normal'>Now add a line control
anywhere on the form and set it’s BorderColor to &amp;H00FFFFFF&amp;. Now just
copy this code and your screensaver will be finished.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Public
bWhite As Boolean<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_Activate()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Line1.X1
= frmMain.Width \ 2<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Line1.Y1
= frmMain.Height \ 2<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Line1.X2
= frmMain.Width \ 2<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Line1.Y2
= frmMain.Height \ 2<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Timer1.Enabled
= True<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_Click()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
GetSetting(&quot;Blanker&quot;, &quot;Settings&quot;, &quot;Password&quot;,
&quot;&quot;) &lt;&gt; &quot;&quot; Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmPassword.Show
'If a password is set then show the password box<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Else<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver
'exit the screensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_DblClick()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
GetSetting(&quot;Blanker&quot;, &quot;Settings&quot;, &quot;Password&quot;,
&quot;&quot;) &lt;&gt; &quot;&quot; Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmPassword.Show
'If a password is set then show the password box<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Else<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver
'exit the screensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
GetSetting(&quot;Blanker&quot;, &quot;Settings&quot;, &quot;Password&quot;,
&quot;&quot;) &lt;&gt; &quot;&quot; Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmPassword.Show
'If a password is set then show the password box<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Else<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver
'exit the screensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_KeyPress(KeyAscii As Integer)<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
GetSetting(&quot;Blanker&quot;, &quot;Settings&quot;, &quot;Password&quot;,
&quot;&quot;) &lt;&gt; &quot;&quot; Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmPassword.Show
'If a password is set then show the password box<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Else<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver
'exit the screensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
GetSetting(&quot;Blanker&quot;, &quot;Settings&quot;, &quot;Password&quot;,
&quot;&quot;) &lt;&gt; &quot;&quot; Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmPassword.Show
'If a password is set then show the password box<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Else<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver
'exit the screensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_Load()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>ShowCursor
False 'Hide the cursor<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>bWhite
= True<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As
Single)<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
GetSetting(&quot;Blanker&quot;, &quot;Settings&quot;, &quot;Password&quot;,
&quot;&quot;) &lt;&gt; &quot;&quot; Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmPassword.Show
'If a password is set then show the password box<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Else<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver
'exit the screensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As
Single)<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Static
x0 As Integer<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Static
y0 As Integer<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">    </span>' Do nothing except in screen saver mode.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">    </span>If RunMode &lt;&gt; rmScreenSaver Then
Exit Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes"> </span><span style="mso-spacerun: yes">   </span>'
Unload on large mouse movements.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">    </span>If ((x0 = 0) And (y0 = 0)) Or _<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">        </span>((Abs(x0 - X) &lt; 5) And (Abs(y0 - Y)
&lt; 5)) _<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">        </span>Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">            </span>' It's a small movement.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">            </span>x0 = X<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">            </span>y0 = Y<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">            </span>Exit Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><span
style="mso-spacerun: yes">    </span>End If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
GetSetting(&quot;Blanker&quot;, &quot;Settings&quot;, &quot;Password&quot;,
&quot;&quot;) &lt;&gt; &quot;&quot; Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>frmPassword.Show
'If a password is set then show the password box<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Else<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>exitScreensaver
'exit the screensaver<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Private
Sub Timer1_Timer()<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Line1.BorderWidth
= Line1.BorderWidth + 3 'Increase the border width by 3<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
Line1.BorderWidth &gt;= 1000 Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Line1.BorderWidth
= 1 'If the line's border width gets bigger than the screen then set it back to
one.<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
bWhite = True Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Line1.BorderColor
= vbGreen<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>bWhite
= False<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>GoTo
This<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>If
bWhite = False Then<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Line1.BorderColor
= vbWhite<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>bWhite
= True<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
If<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>This:<o:p></o:p></span></p>
<p class=MsoBodyText3><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>End
Sub<o:p></o:p></span></p>
<p class=MsoBodyText3 align=center style='text-align:center'><span
style='font-weight:normal'>Don’t forget: Compile and then change extention
to.scr.<o:p></o:p></span></p>
<p class=MsoBodyText3 align=center style='text-align:center'><span
style='font-weight:normal'>Now my tutorial ends on making screen savers. If you
need further instructions because you did not understand, then check the source
code that came with this tutorial.</span><span style='font-size:11.0pt;
mso-bidi-font-size:12.0pt'><o:p></o:p></span></p>
<p class=MsoBodyText3 align=center style='text-align:center'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Edited By: Aaron Lindsay Now
The Age Of Eleven <o:p></o:p></span></p>
<p class=MsoBodyText3 align=center style='text-align:center'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Created By: Aaron Lindsay Of
The Age Of Ten<o:p></o:p></span></p>
<p class=MsoBodyText3 align=center style='text-align:center'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Friday, March 30, 2002<o:p></o:p></span></p>
<p class=MsoBodyText3 align=center style='text-align:center'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt'>Please Vote<o:p></o:p></span></p>
<p class=MsoBodyText3 align=center style='text-align:center'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText3>Note:<span style="mso-spacerun: yes">  </span>When
testing the screensaver from visual basic, don’t use the shortcut of F5 to
start it because it will instantly disappear. Click on the ‘&gt;’ part.</p>
<p class=MsoBodyText3>Note #2: This was created in VB6. There is no guarantee
that it will work on other versions (It won’t work on VB4 16 bit version &amp;
lower).</p>
<p class=MsoBodyText3>Note #3: If you’re making your own screensaver based on
this tutorial, change all the ‘Blanker’s to the name of your choice.</p>
</div>
</body>
```

