<div align="center">

## Transparent Form Maker 2 \(VB6 addin\)


</div>

### Description

This is an Add-In Project that creates transparent forms and adds them to your project. This Add-In works almost exactly like the Transparent Form Maker program that I posted a while back on PSC, except that I changed a few functions to add extra LoadByte Procedures to handle the size limit of Procedures (64k is the limit) and I added a way to move the form without using the task bar(see the Form_MouseDown event on the frmRegion).The Created form will also contain the mousedown code.(Note: Some of you may need to change the reference to the Microsoft Office 9.0 Object Library to the version of Office that you have installed.)

To get this to work open the project and compile it then close the project and open a new project. Select Add-In Manager from the Add-Ins Menu. Check the Loaded/Unloaded Check box (You can also check the Startup check box to load the add-in when Visual Basic starts). The add-in will then be listed in the Add-Ins Menu as "SSE Transform". Follow the instructions on the Main Dialog to create and add transparent forms to your project. (Note: The program may take a few minutes creating a form depending on the complexity and size of the picture that you use.)
 
### More Info
 
If you run the add-in from the IDE it will not put the picture on the form.(see code comments in the frmRegion.SaveForm function. If you compile the dll and load it with the Add-In manager it will work fine.)


<span>             |<span>
---                |---
**Submitted On**   |2004-07-30 08:19:00
**By**             |[Robert Gainor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-gainor.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Transparen177653812004\.zip](https://github.com/Planet-Source-Code/robert-gainor-transparent-form-maker-2-vb6-addin__1-55288/archive/master.zip)








