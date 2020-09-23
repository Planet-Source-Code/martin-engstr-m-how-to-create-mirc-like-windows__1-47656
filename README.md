<div align="center">

## How to create mIRC like windows

<img src="PIC20038121548509436.JPG">
</div>

### Description

This article will show you how to make unlimited windows out of 1, think of mIRC.. you think he made all 30 windows manual? no he made like this, a template window then created 30 more of it the same way as this tutorial will show you.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Martin Engström](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/martin-engstr-m.md)
**Level**          |Intermediate
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/martin-engstr-m-how-to-create-mirc-like-windows__1-47656/archive/master.zip)





### Source Code

==================================
<br>
 HOW TODO MIRC LIKE WINDOWS
<br>
   TUTORIAL BY:
<br>
   Xealot
<br>
==================================
<br>
<br>
Hi all, now im gonna explain to you how to create
<br>
unlimited ammount of windows out of 1 form ( a.k.a. template ).
<br>
Well, to start with this tutorial will require some programming
<br>
experience by you cause im not gonna explain everything i.e.
<br>
what a TYPE (inside a module) is or what variables is =/ ..
<br>
u have to find tutorials for that elsewhere...
<br>
<br>
Basicly what u need to know is how to create forms,
<br>
modules and how to type :P
<br>
But if you wanna understand the code, u might need to learn
<br>
what a TYPE is inside a module aswell.. and ** IMPORTANT **
what an ARRAY is!! ( this thing is basicly based on an array )
<br>
<br>
Lets start, theoreticly you index an form.. ( i hope you know how to index a control ).
<br>
Like you know you can make Text1(1).text .. thats an indexed
<br>
Text1.. and im calling the #1 .. it works the same as an array.
<br>
We are going to do the same to a form BUT .. as you may have noticed..
<br>
( if you take a look at the properties window ) theres nothing called
<br>
"Index" on form..
<br>
<br>
So this is a bit weird, create an module and inside type like:
<br>
<br>
public type Xealot
<br>
q as new form1
<br>
end type
<br>
public io(1 to 5) as xealot
<br>
<br>
Ok let me explain what this code does..
<br>
Say you have made a FORM called Form1.. ( this is the form we wish to make loads of ).
<br>
this piece of code will make an TYPE as we call "xealot" << it doesnt matter what you call it.
<br>
inside types, as you may know.. you can DIM variables.. except you dont type public/dim ..
<br>
we DIM an variable called q as a new Form1..
this means you create a new copy of Form1..
<br>
a bit strange.. basicly it creates a new window that looks like Form1.. also executes
<br>
same code as in Form 1.
<br>
<br>
public io(1 to 5) as xealot
<br>
<br>
that will make an public variable called io ... it has an array from 1 to 5..
<br>
the number 1 to 5 means we will create 5 windows of Form1.. ( change the number if you wish ).
<br>
now.. its *kinda* done.. make a Form2 and goto project properties and set Form2 as the startup form.
<br>
Inside form2 make an button (commmand) and its code would be:
<br>
<br>
for n = 1 to 5
<br>
io(n).q.show
<br>
next
<br>
<br>
DONE! this would make so when you press command1 << the button,.. 5 windows that looks like Form1
<br>
should appear :D ... have fun..
<br>
oh yes and .. a small trick like saving info as "This copy of Form1 is number 4 thats created" is simple..
<br>
You ever wondered what the propertie "tag" is in a form? its basicly
<br>
an variable for u to put nonsense in.. you can store stuff in there.
<br>
<br>
Lets experiment with it, shall we?
<br>
Ok lets edit the command1's code in Form2 to:
<br>
<br>
for n = 1 to 5
<br>
io(n).q.show
<br>
io(n).q.tag = trim$(str$(n))
<br>
next
<br>
<br>
then lets have fun with Form1, make an button..
<br>
inside the code (command1) type:
<br>
<br>
msgbox(me.tag)
<br>
<br>
This should look like, say you press the button on the 3'rd form thats created,
<br>
you should get a msgbox saying "3" :)
<br>
<br>
Thats not too hard was it?
<br>
ah well.. have fun!
<br>
<br>
[  [ Xealot ]   ]
<br>
[ littlebrainbug@hotmail.com ]
<br>
[ #Fraggedsoft @ Quakenet ]
<br>

