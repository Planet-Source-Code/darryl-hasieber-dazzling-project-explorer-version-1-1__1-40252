This VB AddIn was written as for a couple of reasons:
1. It was something I wanted to do. I actually started out writting an application that would graphically display an applicaions structure. I wanted to see all the projects that made up an application, the components, controls, procedures, ariables etc inside the project and to then see were each of these were used. There are applications that do this but I just felt like having a go at performing the same thing myself.
2. I read something in the book Code Complete by Steve McConnell that sounded brilliant and so changed my application to a Project Explorer VB AddIn. Basically he suggested that a good project explorer should show items right down to the procedure level. I thought that this would be a good half way point for my original application.
3. I had never written a VB AddIn and so thought that this would be a usefull and pretty cool one to start with.

This version that you have is version 1.1, the second generation of this AddIn. The original is still available at Planet Source Code and my website http://www.dazzlingsoftware.com. I have left v1.0 as I think that maybe it should be left out there as a 4/5ths finished AddIn which because of its structure is in some ways a better example. Having said that this one is better as I have streamlined it by removing unnecessary varaible, procedures etc.

Changes Since Version 1.0:
Each node is only built when expanded resulting in an apparently faster performance as it will only build what is required at the time. If the node is collapsed and expanded again it will be rebuilt, I left it this way rather than performing a check and cancelling the build if already built to guarantee that the information in each node will be up to date. If you wanted you could alter this to check if the node was built and if so just expand instead of building. This would be done by checking if there was only one child node and then checking if it was the dummy node, if so then build else cancel (I have actually added the code in all you need to do is remove the comments.
I have cut down on the number of procedure level variables in order to improve memory usage. This can be imroved upon even further by removing the CurrentParentNode and LastNodeAdded procedure level varaibles but the code is simpler this way so I have left them in as I don't think the memory usage in holding two nodes in memory is a major one (then again I have 512MB of RAM). The reason for me not usually liking procedure level variables is the lack of control you have on them and memory that is used by them.
I have removed procedures and variables that are no longer used.
I have fixed a error on opening the explorer when no project is open.
I have tried to improve performance but one loop is still slow and has been commented as such in the code.
I have added more comments to this version in an attempt to make this a better tutorial.
I have modified the refresh popup menu to refresh the node the mouse is on or in line with. Because the NodeClick event fires before the MouseUp event if a user right clicks on a node the popup menu will not appear if we implement the NodeClick event. To get around this you have to use HitTest to check if the person clicked on the Node or elsewhere if the user clicked with the left mouse button. Read the code.
I have now also given assigned a Base Address for the DLL (search MSDN to learn more) this is something you should do on all your DLL's. I keep a spreadsheet to keep track of addresses I have used and to decide what address to use
I have removed the office object from the references section and declared it as Object which solves potential compatability problem and a ditribution/packaging problem. 

I think this application can teach more than just how to write an AddIn or build a Tree. I believe that the naming I have used is a good example of how to name your variables and procedures. So many, Too many, developers have this habit of using short, cryptic, non-descriptive names in there code. I also feel that the structure of the code is a good example of how to break up your code into smaller more manageable procedures that perform one function only. The advantages of this are that:
1. Errors are easier to deal with, and locate.
2. You can give procedures better names which ultimately should result in self documenting code.
3. Blocks of code that are repeated ar housed in one procedure meaning that during maintenance you are less likely to forget to update a block in a procedure.
Having said all that I still realize there is much to learn and I am still trying to write the perfect program in terms of naming, design, construction etc. and I suppose the lesson there is as long as we can can recognise and admit that we have not ahieved perfection we can and will continue to grow and get better.

Searching for the string 'TODO' in the code wil take you areas of pieces of code that are either incomplete or need to be analysed and improved upon. If anyone does improve on any of these or enhances this code in any way I would sincerely appreciate an e-mail or similar explaining what was changed so that I can learn from you.
If you chose to use this class module as whole I request that you please keep a refernce to me in it
and not claim this as your own work. Apart from being dishonost (legally called plagarism) it could land you in trouble claiming to be capable of things you are not.
This code was written for the public domain by Darryl Hasieber (www.dazzlingsoftware.com). You are free to use this project in whole or in part for any purpose you see fit (except toilet paper).

Please Check out my other published source code on www.planet-source-code.com or visit my website www.dazzlingsoftware.com where you will find everything I have published so far.
Thanks
Darryl Hasieber
Dazzling Software
Contact Details:
WebSite: 	http://www.dazzlingsoftware.com
E-Mail:		darrylha@dazzlingsoftware.com

For your info incase anyone was wondering I have estimated my total time on this project at approx. 56 hours which includes version 1.0.0
