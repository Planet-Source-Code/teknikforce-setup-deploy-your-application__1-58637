<div align="center">

## Setup & Deploy Your Application


</div>

### Description

Okay, so you've finished your application and are all set to distribute it to your clients, but do you know how? Are you aware of the issues and pitfalls you'll need to face while you design the setup for your application? Do you know about the right tools? This article discusses the basics of setup design, and will talk about several important issues that you should take care of when you write your setups.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TeknikForce](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/teknikforce.md)
**Level**          |Intermediate
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access, VBA MS Excel
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/teknikforce-setup-deploy-your-application__1-58637/archive/master.zip)





### Source Code

```
<h1><span style='font-size:14.0pt;mso-bidi-font-size:12.0pt'>Setup &amp; Deploy
your application: Writing A Good Setup For Your Application<o:p></o:p></span></h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>So you’ve finally managed to complete your latest
application after putting in a lot of last minute effort, spending sleepless
nights and consuming a million cups of Coffee. Pleased with yourself you’re
ready to lean back and relax. The grunt work’s all done. Right? Wrong!<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p>A veteran programmer won’t consider an application
complete until he has shipped it to the client and waited for 15 days. The
successful deployment of an application is just as tough if not any harder than
the actual coding of the application. I should know, I’ve made packaged
software that tens of thousands of users use each day. I’ve experienced a
seemingly endless variety of setup problems and provided solutions to them,
always after a valiant struggle. This article talks about that. Making setups
for your applications that your customers can run to install your app.</p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>Let’s act as if we are absolute beginner and can’t tell a
mouse from the rodent that scurries about our kitchens.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>What is a setup?<o:p></o:p></span></b></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>The job of the setup is to<o:p></o:p></span></p>
<ol style='margin-top:0in' start=1 type=1>
 <li style='mso-list:l0 level1 lfo1;tab-stops:list .5in'><span
   style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>Copy
   all the files that your application needs to the target computer.<o:p></o:p></span></li>
 <li style='mso-list:l0 level1 lfo1;tab-stops:list .5in'><span
   style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>Give
   users access to some kind of icon so that he can click on it and run your
   app.<o:p></o:p></span></li>
 <li style='mso-list:l0 level1 lfo1;tab-stops:list .5in'><span
   style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>Provide
   users with a way to un-install your app (very important) if they later
   want to.<o:p></o:p></span></li>
 <li style='mso-list:l0 level1 lfo1;tab-stops:list .5in'><span
   style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>Perform
   all registry entries, file associations etc., which your application might
   need.<o:p></o:p></span></li>
</ol>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>Why do you need a setup?</h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>For the pure and simple reason that in today’s development
environment there’s a very low chance that you’ll ever make a useful app that
will work all by itself. Most applications need other files to function with,
like databases, runtime file, DLLs, etc. The setup is needed to make sure that
all the files/information that your application needs is in place and ready for
use. Also it is there to give the user an easy way to install your app on his
computer and make the software accessible to him when he needs it.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>Before we start</h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>This is a comprehensive tutorial, a big one. We are not
going to limit ourselves to any single setup program or issue, but instead try
to get a working grasp on the setup process itself. <o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>More about MSM files</h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>As applications grow bigger and use more and more
components, it’s often impossible to keep track of all the runtime files and
registry entries needed to successfully install them on a PC. Microsoft
developed the MSM file format to solve this problem. An MSM file is a
collection of all files and registry entries needed to install a component in a
single file. Consider it like a zip file with registry information. Most of the
leading professional setup tools like Installshield, etc., have full support
for imports of MSM files into their setup. Microsoft’s tool Visual Studio
Installer too supports MSM files. The runtime installations of almost all major
Microsoft technologies are now available in the merge module format that you
can just drop into your project to install.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>Setup fundamentals</h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>Let’s explore the basic philosophy behind setup in a little
more detail. The task of the setup is to perform all tasks and actions
necessary to ensure that your software runs correctly. For most software this
means copying files, performing registry entries and creating an icon that the
user can click to run your application. If something unexpected happens and
your setup fails, then you have to make sure that you undo all the changes that
you made to the user’s PC. Your setup should also create an un-installer to
remove all files and registry entries that you may have made on the computer.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>Making the setup</h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>To illustrate the different steps involved in designing a
setup, let’s make the setup for an imaginary application called Fooapp. We’ll
go through the different steps of making the setup for Fooapp and at the same
time we will also learn about a number of different problems that you as a
setup maker will need to overcome.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>Identify the Dependencies</h1>
<p>First we have to figure out what files to package with
Fooapp. Let’s suppose that we’re using a database to store the addresses of
FooCustomers through FooApp, and we also have a set of pictures that we need to
load in FooApp. Our application depends on these files to run (thus
‘dependencies’). Our setup should copy these files to the correct folder so
that our application can access them when it runs. </p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>We can call these files the first level dependencies:
Files that are unique to your software and are absolutely necessary. These
files are usually copied within the folder where you copy the mail executable
for the application (APPDIR). </p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>In these days of distributed development it’s hardly
possible that you will be able to make an app without using any components.
These components may include ActiveX controls, DLL files, .Net assemblies, etc.
These files form the second level dependencies for your setup project. These
files are necessary for your application but may be ‘shared’ by other
applications.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>There are many strategies to identify the dependencies for
your project. You can use a tool like Depends.exe that ships with Visual Studio
to identify dependencies on DLLs. If your application is made using Visual
Basic then you can use the ‘Package and deployment wizard’ to generate a list
of dependencies for you. Sometimes you will have to use your experience to
identify the technologies you have used and the files you need to run it. This
is especially true when you use a lot of third party components that may in
turn use other components themselves.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>A good way to make sure that you have all dependencies is
to keep installing your application on a ‘clean’ installation of Windows. This
way you will be able to identify each and every file that your application
needs to run.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p><b>Copying files<o:p></o:p></b></p>
<p>Once you’ve identified all dependencies you’re ready to
copy them to the target PC. You should know where to drop each file so that
your application can use them. Here’s a quick categorization of some basic
types.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p><span style='font-family:Wingdings;mso-ascii-font-family:
Arial;mso-hansi-font-family:Arial;mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>à</span></span>
Your application files – Your main executable file and all data files
associated with it should be copied to the application directory. Although it’s
quite possible to install to and use the data from any directory on the PC,
it’s recommended that you restrict your files to the application directory
only. These files may include your database files, pictures, etc.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p><span style='font-family:Wingdings;mso-ascii-font-family:
Arial;mso-hansi-font-family:Arial;mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>à</span></span>
C/C++ Style DLLs – The old c/c++ style DLLs have to be referenced by path. When
you reference such a DLL (like when calling a Windows API function) Windows
will first look in your application directory, and then in the Windows System
directory for that DLL. If it can’t find the DLL you will get an error. If
you’re using one of these DLLs in your project and it’s unique to it then you
should copy it to your application folder (the directory where you copy the
main executable). If you’re using a DLL that is used by many applications then
you can copy it to the Windows\System folder. This is just a matter of personal
taste.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p><span style='font-family:Wingdings;mso-ascii-font-family:
Arial;mso-hansi-font-family:Arial;mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>à</span></span>
ActiveX DLLs and OCX files – Before your application can use the ActiveX DLLs
and custom controls, information about them has to be entered in the Windows
Registry. This information includes the Globally Unique Identifier(GUID) for
the component, and the path where it resides. By convention ActiveX DLLs and
OCX files are always copied to the Windows System folder. Although it’s
possible to use an ActiveX file in any other directory too, it’s recommended
that you copy only to the Windows System folder.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p><b>Registering ActiveX controls<o:p></o:p></b></p>
<p>As I told you in the earlier section, any ActiveX file
that you use has to be entered in the Windows registry before your application
can use it. Microsoft provides a free tool to add this information to registry
and most setup programs too use it. The tool is ‘regsvr32.exe’, it can be found
in your Windows system directory. </p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>Most setup programs allow you to mark a file for
registration. There might be flags like ‘regserver’ (in InnoSetup) or you may
be able to select an attribute that will allow you to mark a file for
registration. Make sure that you test each file before you mark it for
registration in the Windows registry if you mark an invalid file your setup
will show an error, and sometimes it may even crash.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>To test whether a file is valid or not, run the following
command on it – ‘regsever32 filename.’ The file may have to be copied to the
Windows System directory before you can do it. </p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p><b>Shared Files<o:p></o:p></b></p>
<p>The concept of ‘Shared files’ is very important and I’ve
found that many first time setup makers who neglect this end up earning the
wrath of their users afterwards. </p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>In today’s development scenario almost every programmer
makes the use of third party activex components and DLLs. In a common situation
many applications installed on a computer will use the same components to run.
If during the un-installation process your application deletes these ‘shared
files’ or changes them in any way then those other applications will not run.
As a responsible setup maker it’s your job to ensure that your application
installs and uninstalls without wrecking any other application on the user’s
computer, so you have to take great care that you do not delete any shared file
from the user’s PC.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>How to identify whether a file is shared?</p>
<p>The thumb rule is: if it’s not unique to your app, it’s
shared. The shared files will include all activex DLLs and OCX controls not
developed by you exclusively for this application (like the tab control, or the
chart control), all DLLs that are a part of any standard library (like
MSVCRT.DLL, or MSVBVM60.DLL, etc). If your setup deletes any of these files
during un-installation, you should prepare to welcome a very angry customer at
your doorstep.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p><b>Overwriting Used Files<o:p></o:p></b></p>
<p>Often you will need to replace the older version of a DLL
with a new version. You can do this without a problem if your DLL file is not
in use by any other application, however if it’s being used by any other app and
you try to overwrite it, your setup will crash! Some system DLLs are
perpetually in use and can’t be overwritten at all when Windows is running. To
overwrite these files you need to reboot Windows and overwrite them before
Windows loads completely. Most setup making programs allow you to mark such
files for ‘reboot before overwriting.’ Take care that you identify and mark
these files correctly or your setup will never run to the end.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>The most common files that need re-start for overwriting
are – OLEAUT32.DLL and OLEPRO32.DLL. These files are used by virtually every
Windows based applications and can’t be over-written without re-starting. You
should identify the other files manually.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>To manually identify such files – load as many
applications as you can, and then try copying each file in your application to
its target location manually using Windows Explorer. If you can successfully
copy it then the file does not need rebooting. If you can’t, then better mark
this file for rebooting.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p><b>Creating icons for your application<o:p></o:p></b></p>
<p>After you’ve finished copying all the files needed to run
your program you should create the icons through which your users may run the
application you made. If you use a setup making program like Innosetup, or
Installshield, you will be able to specify an icon file and the file to which
the icon will point. It’s recommended that you create an icon for your main
executable, the help file and also for your company’s webpage.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p>You may also create an icon on the user’s desktop and on
the quick launch bar if you think your application will be used very often. Do
make sure however to ask the user before you create an icon on the desktop or
the quick launch bar. It’s bad manners not to.</p>
<p><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h1>Un-installation</h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>Un-installation is just as important as installation for any
setup. Most new setup making programs create an un-installer automatically that
the user can run from the ‘Add/Remove Programs’ section of the Windows control
panel. Make sure that you do not remove any vital file that can effect other applications
during un-installation (look up the section on shared files above.) However,
most of the time this will be transparent to you and will not require any work
from your side, so relax </span><span style='font-size:10.0pt;mso-bidi-font-size:
12.0pt;font-family:Wingdings;mso-ascii-font-family:Arial;mso-hansi-font-family:
Arial;mso-bidi-font-family:Arial;mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>J</span></span><span
style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>That’s it!</h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>Yep, that’s the setup making process in a gist, and all the
important precautions have been outlined for you. Your skills will grow from
setup to setup though, so get right to it and have fun!<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>Do not forget to test your setup a lot. The best thing to do
is to install on a clean installation of Windows, and then also on a very-used
computer. Your app should install and work on both,<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>Some free Setup making utilities</h1>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Wingdings;mso-ascii-font-family:Arial;mso-hansi-font-family:Arial;
mso-bidi-font-family:Arial;mso-char-type:symbol;mso-symbol-font-family:Wingdings'><span
style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>è</span></span><span
style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'> Innosetup
– My favorite setup maker, it’s free, fast and efficient. <br>
(Download: http://jrsoftware.org)<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Wingdings;mso-ascii-font-family:Arial;mso-hansi-font-family:Arial;
mso-bidi-font-family:Arial;mso-char-type:symbol;mso-symbol-font-family:Wingdings'><span
style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>è</span></span><span
style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'> Visual
Studio Installer 1.1 – Microsoft’s installer. Not bad. This is free too.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>(Download:
http://msdn.microsoft.com/vstudio/downloads/tools/vsi11/download1.aspx)<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>Innosetup has a huge community of developers who absolutely
love it because of what it can do. It’s one of the fastest installers on the
net, it’s absolutely free and it can do everything that a professional setup
application like Installshield can do (without the fancy bells and whistles
though.) However, there’s one important feature that Innosetup doesn’t have.
Support for Microsoft Merge Modules (.MSM files), that’s where the Microsoft
visual studio installer comes in.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>Innosetup also has scripting capabilities and is perfect for
power programmers who need to make their setups as flexible as they can. The
only feature that InnoSetup lacks is support for imaged during installation and
for MSM files. Using Innosetup requires some coding skills and you’ll need to
spend some time studying it.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>You should use the Microsoft’s Visual Studio Installer
wherever you can’t use Innosetup. This will mostly include situations in which
you can’t find an alternative to using MSM files (Installer Microsoft Speech
API 5 for example, the runtimes are only available in MSM format.) Be warned
though, Visual Studio setups are slow and klunky. They just can’t compare with
the blazing speed or ease of use of setups made using Innosetup.<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>By Cyril M Gupta<o:p></o:p></span></p>
<p><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Arial'>Cyril@cyrilgupta.com<o:p></o:p></span></p>
```

