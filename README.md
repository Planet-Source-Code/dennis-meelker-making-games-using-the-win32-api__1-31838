<div align="center">

## Making games using the Win32 api


</div>

### Description

This is an complete tutorial on how to make a game using the Win32 api. The pictures wouldn't work so download the zip. Please leave a comment to tell me what you think of it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-02-15 21:16:48
**By**             |[Dennis Meelker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dennis-meelker.md)
**Level**          |Beginner
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Making\_gam554672152002\.zip](https://github.com/Planet-Source-Code/dennis-meelker-making-games-using-the-win32-api__1-31838/archive/master.zip)





### Source Code

```
<p class=MsoNormal align=center style='text-align:center'><b><span
style='font-size:24.0pt;mso-bidi-font-size:12.0pt'>Making games using the Win32
api.<o:p></o:p></span></b></p>
<p class=MsoNormal align=center style='text-align:center'><b><span
style='font-size:16.0pt;mso-bidi-font-size:12.0pt'>By Dennis Meelker<o:p></o:p></span></b></p>
<p class=MsoNormal align=center style='text-align:center'><b><span
style='font-size:16.0pt;mso-bidi-font-size:12.0pt'>Meelkertje@hotmail.com</span></b><br
clear=all style='mso-special-character:line-break;page-break-before:always'>
</p>
<p class=MsoNormal>In this tutorial I will show you how to make a game that
runs fast using the Win32 Api. I will try to explain everything as good as
possible, so if you don’t understand something read it over and over until you
get it, got it? If you really can’t understand it you can always e-mail me.Now
lets get started.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>So why should we use the Win32 API instead of using DirectX,
I personally think DirectX is way to hard to learn if you just want to make a
flat type game. I you want to make a 3D shooter with all effects like anti
alias and stuff, you will need to learn DirectX for sure, there is just no way
you can do this in VB using only API’s. But if you want to make a flat game
like Pacman or some kind of Platform I prefer using the API. But this is just
my opinion so if you want to use DirectX go ahead.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>First we will make a new project just create a standard exe,
now name the form frmMain or something like that. This will be the game form. Now,
the next two things are really important, set the AutoRedraw property to True
and set The Scalemode to Pixel.</p>
<p class=MsoNormal>When Autoredraw property is set to true the things that our
game drew on the form won’t just disappear when the form is refreshed. We set
the Scalemode to pixel because the API’s all need pixels as parameter and not
twips, if we hadn’t changed it we had to turn the twips into pixels all the
time witch would cause trouble for sure.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>Now that we have our form ready create a new module and call
it something like modInvaders orso. Now add these api Declaration’s to the
module:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Declare Function BitBlt Lib
 "gdi32" Alias "BitBlt" (ByVal hDestDC As Long, ByVal x As
 Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal
 hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
 As Long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SRCAND = &H8800C6<span
 style="mso-spacerun: yes">  </span>' (DWORD) dest = source AND dest<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SRCCOPY = &HCC0020 ' (DWORD) dest
 = source<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SRCPAINT = &HEE0086<span
 style="mso-spacerun: yes">        </span>' (DWORD) dest = source OR dest</span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><span style="mso-spacerun: yes"> </span></p>
<p class=MsoNormal>As you can see, we just added the Bitblt api function to our
project, the bitlbt function will be the core of our game, it is used to draw
all the graphics on the screen. I will explain all the parameters here:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>hDestDC<span style='mso-tab-count:1'>          </span>- This
is the DC of the form/control to draw to.</p>
<p class=MsoNormal>x<span style='mso-tab-count:2'>                      </span>-
This is the x position to draw to</p>
<p class=MsoNormal>y<span style='mso-tab-count:2'>                      </span>-
This is the y position to draw to</p>
<p class=MsoNormal>nWidth<span style='mso-tab-count:2'>             </span>-
This is the width of the picture or a part of a picture to copy</p>
<p class=MsoNormal>nheight<span style='mso-tab-count:2'>              </span>-
This is the height of the picture or a part of a picture to copy</p>
<p class=MsoNormal>hSrcDC<span style='mso-tab-count:1'>            </span>-
This is the DC of the form/control that containt the picture we want to copy</p>
<p class=MsoNormal>xSrc<span style='mso-tab-count:2'>                 </span>-
The source x coordinate</p>
<p class=MsoNormal>ySrc<span style='mso-tab-count:2'>                 </span>-
The source y coordinate</p>
<p class=MsoNormal>dwRop<span style='mso-tab-count:2'>             </span>-
This specifies how the graphic should be drawn</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>When you see this you will probably say, “Dennis? What is a
DC?”.</p>
<p class=MsoNormal>So I will explain that right now, DC stands for device
context, it’s just a place inside the memory of your PC where the picture of a
form or any anther control is stored. Now if you didn’t know what a DC was you
probably won’t know what the dwRop property is, well, as I said above it tells
the BitBlt functions how to draw, “Are there different way’s to draw??”, yes
there are, you noticed the three constants below the BitBlt function that we
added to our module those are three ways to draw, you can just set the dwRop to
SRCCOPY or any of them. The SRCAND and SRCPAINT are almost always used
together, I don’t know what they do exactly but I do know that when you use the
SRCAND first and the SRCPAINT after that you can get a transparent picture, you
probably don’t understand what I just said, never mind, I will explain this
later on. The SRCCOPY constant is nothing special, it just copies the part of a
picture you specified to the specified DC.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>If we are making a good game, we want the area surrounding
our characters to be transparent, if it would not be transparent, below are two
examples of a character, the first one has a transparent background, the second
one doesn’t.</p>
<table cellpadding=0 cellspacing=0 align=left>
 <tr>
 <td width=60 height=0></td>
 <td width=61></td>
 <td width=47></td>
 <td width=60></td>
 </tr>
 <tr>
 <td height=60></td>
 <td align=left valign=top><img width=61 height=60
 src="./Making%20games%20using%20the%20Win32%20api_files/image003.jpg"
 v:shapes="_x0000_s1027"></td>
 <td></td>
 <td align=left valign=top><img width=60 height=60
 src="./Making%20games%20using%20the%20Win32%20api_files/image004.jpg"
 v:shapes="_x0000_s1026"></td>
 </tr>
 </table>
 </span><![endif]><!--[if gte vml 1]></o:wrapblock><![endif]--><br
style='mso-ignore:vglayout' clear=ALL>
<![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<br style='mso-ignore:vglayout' clear=ALL>
<p class=MsoNormal><!--[if gte vml 1]><v:shape id="_x0000_s1028" type="#_x0000_t75"
 style='position:absolute;margin-left:0;margin-top:63.6pt;width:99pt;height:49.5pt;
 z-index:3;mso-position-horizontal:left'>
 <v:imagedata src="./Making%20games%20using%20the%20Win32%20api_files/image005.png"
 o:title=""/>
 <w:wrap type="square"/>
</v:shape><![if gte mso 9]><o:OLEObject Type="Embed" ProgID="PBrush"
 ShapeID="_x0000_s1028" DrawAspect="Content" ObjectID="_1075312935">
</o:OLEObject>
<![endif]><![endif]--><![if !vml]><img width=132 height=66
src="./Making%20games%20using%20the%20Win32%20api_files/image006.jpg"
align=left hspace=12 v:shapes="_x0000_s1028"><![endif]>I think you will now
understand why we want the background of our character transparent. To make the
background of our pictures transparent we have to create a mask. On the right
you can see a picture that is ready to be drawn with a transparent background.
The first picture has a black background, the black will be transparent, now
you will probably think: “But his arms have black stripes, wont he get
transparent arms?”, to solve that problem I created a “Mask” a mask is a
picture with everything that should be transparent white, and the rest black.
With these two images and the BitBlt function we are ready to draw the
character with a transparent background.</p>
<p class=MsoNormal><span style="mso-spacerun: yes"> </span></p>
<p class=MsoNormal>Before we go on, make sure you have a picture like the one
above,<span style="mso-spacerun: yes">  </span>you can also use the one above
if you want to.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>Now add a Command button and a picturebox to the form set
the picturebox’s autoredraw property to true, the scalemode to Pixel, the
borderstyle to zero and set the Autosize property to true. Now load your
picture in the picturebox by setting the picture property.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>Doubleclick on the button you just inserted and add the
following code:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span lang=EN-US style='font-size:10.0pt;mso-bidi-font-size:
 12.0pt;font-family:"Courier New";mso-ansi-language:EN-US'>BitBlt me.hDc, 0 ,
 0 , 20, 20, picture1.hDc, 0, 0, SRCCOPY<o:p></o:p></span></p>
 <p class=MsoNormal><span lang=EN-US style='font-size:10.0pt;mso-bidi-font-size:
 12.0pt;font-family:"Courier New";mso-ansi-language:EN-US'>Me.Refresh</span><span
 lang=EN-US style='mso-ansi-language:EN-US'><o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><span lang=EN-US style='mso-ansi-language:EN-US'>If your
picture has other sizes you must change them, but watch out, you have to use
the sizes of one part of the picture, the picture I showed you has a width of
40 pixels and a height of 20, it consists of two pictures that are put together
so the sizes of one picture are 20x20.<o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-US style='mso-ansi-language:EN-US'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-US style='mso-ansi-language:EN-US'>If you
start the program and you press the button you will see you character with a
black background. Now that you know how to use the SRCCOPY constant we will go
on with the other two. Add another Command button and set it’s code to:<o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-US style='mso-ansi-language:EN-US'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span lang=EN-US style='font-size:10.0pt;mso-bidi-font-size:
 12.0pt;font-family:"Courier New";mso-ansi-language:EN-US'>BitBlt me.hDc, 0 ,
 0 , 20, 20, picture1.hDc, 20, 0, SRCAND<o:p></o:p></span></p>
 <p class=MsoNormal><span lang=EN-US style='font-size:10.0pt;mso-bidi-font-size:
 12.0pt;font-family:"Courier New";mso-ansi-language:EN-US'>BitBlt me.hDc, 0 ,
 0 , 20, 20, picture1.hDc, 0, 0, SRCPAINT<o:p></o:p></span></p>
 <p class=MsoNormal><span lang=EN-US style='font-size:10.0pt;mso-bidi-font-size:
 12.0pt;font-family:"Courier New";mso-ansi-language:EN-US'>Me.Refresh</span><span
 lang=EN-US style='mso-ansi-language:EN-US'><o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><span lang=EN-US style='mso-ansi-language:EN-US'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-US style='mso-ansi-language:EN-US'>When you
press this button you will see your character with a transparent background!!<o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-US style='mso-ansi-language:EN-US'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>The first thing you should now when making a game is that
you should use as less timers as possible , timers are really bad things to use
in a game. Now I hear you thinking things like: “But how can I move the bullet
my spaceship just fired without a timer?”, the solution is:… Loops!. “Loops??”,
yes loops, because a game usual needs to time a lot using timers will only make
your game run slow, and that is the most worst thing to have, imagine you made
a great looking game with killer graphics, but it runs soooooo slow because you
used about two dozen timers. We wont have this kind of trouble, cause we are
using loops!!</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>Well, with the most games you will have a main loop, a main
loop is a loop that runs your game, when the loop stops.. the game stops. In
this loop we will check for keypresses and we will move our characters, we will
move bullets, draw the players and powerups and so on. A simple loop would look
like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Do<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>‘Game
 Stuff<span style="mso-spacerun: yes">  </span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>DoEvents<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Loop<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>But if we put all the stuff I mentioned above in here and
start our loop you will notice it goes way to fast, we will need to slow our
loop a bit down. Now I will change the loop like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'>Const
 TickDifference as long = 10<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'>Dim
 LastTick<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'>LastTick =
 GetTickCount()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'>Do<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><span
 style="mso-spacerun: yes">   </span>Curtick = GetTickCount()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><span
 style="mso-spacerun: yes">   </span>If<span style="mso-spacerun: yes"> 
 </span>Curtick – LastTick > TickDifference then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><span
 style="mso-spacerun: yes">      </span>‘Game Stuff<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><span
 style="mso-spacerun: yes">   </span>End if<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><span
 style="mso-spacerun: yes">   </span>DoEvents<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'>Loop<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>The first thing you will probably see is the GetTickCount
function, if you never worked with the API before you will probably don’t know
what it does. So I will tell you, the function GetTickCount returns the amount
of milliseconds that elapsed since windows has started. So if you look at the
rest of our loop you will see that we first get the current tick and store it
in the LastTick variable. Then we start the loop and we store the current tick
in the CurTick variable. Now comes the important part, we check if the
difference between CurTick and LastTick is ten, if it is the game stuff will be
executed. So if we make our loop this way, every ten milliseconds the game
stuff will be executed, this gives you game a speed of 100 Fps!! The
declaration of the GetTickCount is so:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'>Public
 Declare Function GetTickCount Lib "kernel32" () As Long<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>For the people that don’t know the DoEvents command, I will
explain it here. The DoEvents command is really important in our main loop, the
DoEvents command lets the pc do things like updating the screen, if we let it
out of our loop you would not see anything happen because the pc hasn’t any
time to redraw the screen, so it stays empty.</p>
<p class=MsoNormal>I will now tell you how to obtain keypresses. If you ever
used the keydown event with a game you probably noticed that if you hold a key
down, your character first goes forward one step, then it pauses and then it
goes on. It’s very simple…”We don’t want that!” so we won’t use any event, we
will use the GetKeyState API, it’s declaration is as followed:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Declare Function GetKeyState Lib
 "user32" (ByVal nVirtKey As Long) As Integer<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const KEY_DOWN As Integer = &H1000</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 mso-bidi-font-family:"Times New Roman"'><o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>The only parameter this functions has is the nVirtKey, this
is the key you want to check. I’ve included the KEY_DOWN constant witch is
needed to check for a keypress, if you want to check if the space bar is
pressed you simply use this code</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>If GetKeyState(vbKeySpace) and KEY_DOWN then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>‘Statements<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End If<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>When creating a game you should make a function called
something like: GetUserInput or something like that, it would also be handy if
you declared a Boolean variable for every key so you can use it inside your
whole game, the sub would look like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Function GetUserInput()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>UpPressed = GetKeyState(vbKeyUp) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>DownPressed = GetKeyState(vbKeyDown) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>LeftPressed = GetKeyState(vbKeyLeft) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>RightPressed = GetKeyState(vbKeyRight) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Function<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>If you call the newly created function inside our main loop
we can check for keypresses everywhere in our game. </p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>If you want the keys to customisable you could also declare
a long for every key, then you should also make something like a InitKeys sub
in witch the variables will be loaded with the right keycodes, you can then let
the user choose his own configuration, the code would look like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Sub InitKeys()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>UpKey =
 vbKeyUp<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>DownKey
 = vbKeyDown<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>LeftKey
 = vbKeyLeft<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>RightKey
 = vbKeyRight<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Sub<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Function GetUserInput()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>UpPressed = GetKeyState(UpKey) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>DownPressed = GetKeyState(DownKey) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>LeftPressed = GetKeyState(LeftKey) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>RightPressed = GetKeyState(RightKey) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Function<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>On to the next subject, Sound, without sound, a game is
usual boring, so you need sound. You can download sounds from various websites,
you can also record the yourself using a microphone, if I need a simple sound
like a beng, I just put my mirophone near my desk and punch on the table. If
you have your sound saved as a .wav file you can play it two ways, you can use
the mci control Microsoft made. And you can use the sndPlaySound api, we will
use the sndPlaySound api, because using the control only makes your game slower
and bigger because you will need to include a ocx of 150 kb. De declaration of
the sndPlaySound stands below:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Declare Function sndPlaySound Lib
 "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As
 String, ByVal uFlags As Long) As Long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SND_ASYNC = &H1<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SND_LOOP = &H8<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SND_NODEFAULT = &H2<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>As you can see the sndPlaySound function has two parameters,
ipszSoundName and uFlags. The ipszSoundName is the filename of the .wav file to
play, uFlags can be set to various settings. I explain the setting below:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>SND_ASYC – The file is played and the program continues, if
you don’t use this one the program waits until the sound is done.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>SND_LOOP – The sound will be looped, if you want to stop the
sound just call the sndPlaySound function again, with no file specified.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>SND_NODEFAULT – When this flag is not set, the system
default beep will sound if the given file can’t be found. When you set it there
just won’t be sound. It’s smart to always use this flag when creating a game.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>In most games you will only use the SND_ASYNC and the
SND_NODEFAULT flags, therefore I always create a function called PlaySound,
like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Function PlaySound(sFileName as string)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>SndPlaySound sFileName, SND_ASYNC + SND_NODEFAULT<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Function<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>To play a sound just call the PlaySound function. Because
the sound is usual in the map of the game I always create a function to add a \
to the app.path variable if necessary, I do so because else, if I always add a
\ the path can become something like c:\\, and that wont work. The new code
will then be.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Function FixPath(sPath as string) as string<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>If
 Right(sPath,1) = “\” then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">     
 </span>FixPath = sPath<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>Else<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">     
 </span>FixPath = sPath & “\”<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>End If<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Function<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Function PlaySound(sFileName as string)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>SndPlaySound FixPath(App.Path) & sFileName, SND_ASYNC +
 SND_NODEFAULT<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Function<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>This way you only need to specify the filename, the program
will then automatically add the directory in witch the game is installed.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>Okay, now that you now the basics ( if you don’t know the
basics, just read it again ) we will make a small game in witch you can walk
around a character. Create a new project, make a button that says “new game” on
the first form. Add a new form, set up the form as we did earlier. Make a
module with all the declarations and constants we talked about, they are below:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Declare Function BitBlt Lib "gdi32"
 Alias "BitBlt" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As
 Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long,
 ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Declare Function GetKeyState Lib
 "user32" Alias "GetKeyState" (ByVal nVirtKey As Long) As
 Integer<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Declare Function sndPlaySound Lib
 "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As
 String, ByVal uFlags As Long) As Long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Declare Function GetTickCount Lib
 "kernel32" Alias "GetTickCount" () As Long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SRCAND = &H8800C6<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SRCCOPY = &HCC0020<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SRCPAINT = &HEE0086<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const KEY_DOWN As Integer = &H1000<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SND_ASYNC = &H1<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Const SND_NODEFAULT = &H2<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>Also add some declaration to the module, just put them at
the bottom:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public UpKey as long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public DownKey as long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public LeftKey as long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public RightKey as long<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public UpPressed as boolean<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public DownPressed as boolean<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public LeftPressed as boolean<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public RightPressed as Boolean<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public PlayerX as integer<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public PlayerY as integer<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public TimeToEnd as boolean<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<span style='font-size:12.0pt;font-family:"Times New Roman";mso-fareast-font-family:
"Times New Roman";mso-ansi-language:EN-GB;mso-fareast-language:EN-US;
mso-bidi-language:AR-SA'><br clear=all style='mso-special-character:line-break;
page-break-before:always'>
</span>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>Now create a new sub in the module called MainLoop, you can
also copy it from below</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Sub MainLoop()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>Const
 TickDifference as long = 10<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>Dim
 LastTick<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>PlayerX
 = 0<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>PlayerY
 = 0<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>TimeToEnd = False<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>InitKeys<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>LastTick
 = GetTickCount()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>Do until
 TimeToEnd<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">      </span>If
 GetKeyState(vbKeyEsc) and KEY_DOWN Then TimeToEnd = True<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">     
 </span>Curtick = GetTickCount()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">     
 </span>If<span style="mso-spacerun: yes">  </span>Curtick – LastTick >
 TickDifference then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">        
 </span>GetUserInput<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span>If
 UpPressed Then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">           
 </span>PlayerY = PlayerY -4<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">        
 </span>End if<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span>If
 DownPressed Then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">           
 </span>PlayerY = PlayerY + 4<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">        
 </span>End if<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span>If
 LeftPressed Then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">           
 </span>PlayerX = PlayerX - 4<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">        
 </span>End if<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span>If
 RightPressed Then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">          
 </span>PlayerX = PlayerX + 4<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">        
 </span>End if<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">        
 </span>‘Check if the player is still in the screen<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span>If
 PlayerX < 0 Then PlayerX = 0<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span>If
 PlayerX > 100 Then PlayerX = 400<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span>If
 PlayerY < 0 Then PlayerY = 0<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span>If
 PlayerY > 100 Then PlayerY = 400<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">         </span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">        
 </span>‘Draw the player<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">        
 </span>Draw Player<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">      </span>End
 If<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">      </span>‘Let
 the pc do its stuff<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">     
 </span>DoEvents<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>Loop<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Sub</span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<span style='font-size:12.0pt;font-family:"Times New Roman";mso-fareast-font-family:
"Times New Roman";mso-ansi-language:EN-GB;mso-fareast-language:EN-US;
mso-bidi-language:AR-SA'><br clear=all style='mso-special-character:line-break;
page-break-before:always'>
</span>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>Add the two subs InitKeys and GetUserInput to the module:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Sub InitKeys()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>UpKey =
 vbKeyUp<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>DownKey
 = vbKeyDown<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>LeftKey
 = vbKeyLeft<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>RightKey
 = vbKeyRight<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Sub<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Function GetUserInput()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>UpPressed = GetKeyState(UpKey) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>DownPressed = GetKeyState(DownKey) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>LeftPressed = GetKeyState(LeftKey) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>RightPressed = GetKeyState(RightKey) And KEY_DOWN<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Function<o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal><!--[if gte vml 1]><v:shape id="_x0000_s1029" type="#_x0000_t75"
 style='position:absolute;margin-left:0;margin-top:63.05pt;width:1in;height:36pt;
 z-index:4;mso-position-horizontal:left'>
 <v:imagedata src="./Making%20games%20using%20the%20Win32%20api_files/image007.png"
 o:title=""/>
 <w:wrap type="square"/>
</v:shape><![if gte mso 9]><o:OLEObject Type="Embed" ProgID="PBrush"
 ShapeID="_x0000_s1029" DrawAspect="Content" ObjectID="_1075312937">
</o:OLEObject>
<![endif]><![endif]--><![if !vml]><img width=96 height=48
src="./Making%20games%20using%20the%20Win32%20api_files/image008.jpg"
align=left hspace=12 v:shapes="_x0000_s1029"><![endif]>Now, for the drawing
I’ve created a new sub, in the sub the BitBlt function is called twice, create
a Picturebox on the game-form, set the scalemode to true, borderstyle to zero,
autosize to true and autoredraw to true. Finally, set the visible property to
false. Now add your graphic. I used this one. Now copy the DrawPlayer sub into
the module.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>Public Sub DrawPlayer()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>FrmMain.Cls<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>BitBlt
 frmMain.hDc, PlayerX, PlayerY, 20, 20, Picture1.hDC, 20, 0, SRCAND<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">   </span>BitBlt
 frmMain.hDc, PlayerX, PlayerY, 20, 20, Picture1.hDC, 0, 0, SRCPAINT<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  
 </span>FrmMain.Refresh<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>End Sub</span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>The last thing you need to do is add this to the click event
of the command button on the first form:</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;
 border:none;mso-border-alt:solid windowtext .5pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=619 valign=top style='width:464.4pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>FrmMain.Show<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>DoEvents<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><![if !supportEmptyParas]> <![endif]><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>MainLoop</span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>That’s it, you have created a small game in witch you can
move around a smile, off course this isn’t a fun game, but this game includes
all the basics, now you can add things like a background, you can just set the
picture property of the form to any picture you like. You can turn this into an
RPG or into a pacman type game. You can add sound.</p>
<p class=MsoNormal><![if !supportEmptyParas]> <![endif]><o:p></o:p></p>
<p class=MsoNormal>For more practise with this way of making games, search for
“My Pacman” at http://www.Planet-Source-Code.com , you will find a pacman game
I made, it uses the same techniques as I explained on this tutorial. I hope you
find this document , useful. You can send all questions and other things to
Meelkertje@hotmail.com.</p>
</div>
```

