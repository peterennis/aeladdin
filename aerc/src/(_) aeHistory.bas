Option Compare Database
Option Explicit

' !                   Provided under LGPL v3.0
' !                   http://www.gnu.org/licenses/licenses.html
' !                   GPL.txt and LGPL.txt provided
' !
' !                   Original source of V-Tools
' !                   Copyright Skro129 1999-2010, freeware and licensed LGPL v3.0
' !                   http://www.skrol29.com/us/vtools.php#support
' !
' !                   Thanks to Michael Ciurescu (CVMichael) for permission to use indent code
' !                   Ref: http://www.vbforums.com/showthread.php?t=479449
' !
' !                   Credit to Chip Pearson (MVP Excel) who provides public domain example code on his site
' !                   Ref: http://www.cpearson.com/Excel/LegaleseAndDisclaimers.aspx
' !
' !                   Copyright 2011 Peter F. Ennis
' !                   adaept is (TM) by Peter F. Ennis
' !                   adaept information management is (TM) by Peter F. Ennis
' !                   aeladdinT (for Template), aeladdinA (for Access) and aeladdin are (TM) by Peter F. Ennis
' !                   aeladdin elementary add-in for Access (TM)
' !                   - with appreciation for the recursion in GNU
' !
' ! RESEARCH
' ! Using the VBA Extensibility Library = Ref: http://pubs.logicalexpressions.com/pub0009/LPMArticle.asp?ID=307
' ! Ref: http://www.everythingaccess.com/mdeprotector_example.htm
' ! Ref: http://www.everythingaccess.com/vbwatchdog.htm
' ! Ref: http://www.officekb.com/Uwe/Forum.aspx/excel-prog/69455/Print-VBA-code-in-Editor-format-colors
' ! Ref: http://www.programmerworld.net/resources/visual_basic/visual_basic_addin.php
' ! Comment Out VBA Code Blocks = Ref: http://www.ozgrid.com/forum/showthread.php?t=10432&page=1
' ! There is source code here on how to make an extended search: Ref: http://kandkconsulting.tripod.com/VB/VBCode.htm##AddIN
' ! ==========================================================
' ! ribbon Callbacks and Customizing
' ! ==========================================================
' ! Ref: http://msdn.microsoft.com/en-us/library/aa433869.aspx
' ! Reference for Access 2010: "Microsoft Office 14.0 Object Library"
' ! xmlns="http://schemas.microsoft.com/office/2009/07/customui"
' ! MS Office 2010 Schema Ref: http://www.microsoft.com/downloads/en/details.aspx?FamilyID=C2AA691A-8004-46AC-9852-102F1D5BCD18
' ! Reference for Access 2007: "Microsoft Office 12.0 Object Library"
' ! xmlns="http://schemas.microsoft.com/office/2006/01/customui"
' ! MS Office 2007 Schema Ref: http://www.microsoft.com/downloads/details.aspx?familyid=15805380-f2c0-4b80-9ad1-2cb0c300aef9&displaylang=en
' ! Set position of a ribbon tab Ref: http://www.ureader.com/msg/10972241.aspx
' ! Ref: http://office.microsoft.com/en-us/access-help/customize-the-ribbon-HA010341612.aspx
' !
' ! 20110309 - v001 - To turn off auto start - File\Option\Current Database\Display Form:\(none)
' !                   For fixing packaging add-in
' !                   Ref: http://social.msdn.microsoft.com/Forums/en/accessdev/thread/356669c4-31e0-40ab-93ca-d1648b40bc75
' !                   Administrator permission method to install
' !                   Ref: http://www.skrol29.com/us/vtools.php under Troubleshooting:
' !                   Change constants in "This Wizard", DB_SETUP => ae_SETUP
' !                   Use Form1..Form8
' ! 20110310 - v002 - Rename project to aeladdin. Remove all %AccVerE% references
' !                   aeT_SETUP for startup form.
' ! 20110310 - v003 - Add blnLOG and gstr_LOGFILE.
' ! 20110311 - v004 - More debug info.
' ! 20110313 - v005 - More debug info. Create msi folder and test package
' ! 20110314 - v006 - aeladdinT. Create and test package
' !                   Creating Installable Add-ins for Access
' !                   Ref: http://msdn.microsoft.com/en-us/library/aa140937(v=office.10).aspx
' ! 20110316 - v006 - Access 2010 packaging appears to conflict with this add-in method so do not use
' !                   Set version to 0.01 - do not use 0.0.0 format at this time
' !                   DB_SETUP => aeT_SETUP in table USysMultiLanguage
' !                   Only show Search form. Deep Search => Search
' !                   Licensed under LGPL v3.0, zipped and sent to Skrol29
' ! 20110316 - v007 - Db_SearchThrougthObjects => Db_SearchThroughObjects,
' !                   Set it first in list, use Search as the test
' !                   NOTE: It only finds one reference in the code. Table name and references in
' !                   SysMultiLanguage are NOT changed - do manually - Bug?
' !                   Make Const c_MainForm = "aeT_ABOUT" as global
' ! 20110317 - v007 - m_AccVersNum => zzzm_AccVersNum, m_AccVer => zzzm_AccVer and fix all references
' ! 20110406 - v008 - Import "(_) _COMMANDS_" and aeNumClass and test
' !                 - Ref: http://www.databasedev.co.uk/access-add-ins.html -- Load and configure the USysRegInfo table
' !                 - Ref: http://answers.microsoft.com/en-us/office/forum/office_2007-office_install/access-2007-add-in-manager-security-issues/9ac7fa8e-9a3f-4af8-96f2-b4162c55cf63
' ! 20110408 - v009 - Test v009 as add-in works using how to references of v008. Added marker in comment col to stop re-align
' ! 20110411 - v010 - Test v009 with Access 2007 32 bit on Vista 64. Does not work. Ref: http://technet.microsoft.com/en-us/library/ee681792.aspx
' !                 - Add USysRibbons Ref: http://msofficeuser.com/pages/access/how-to-create-the-usysribbons-table-in-microsoft-access-2007
' !                 - Ref: http://msdn.microsoft.com/en-us/library/aa338202(v=office.12).aspx#OfficeCustomizingRibbonUIforDevelopers_Callbacks
' ! 20110415 - v011 - Install IDBE RibbonCreator 2010. Testing for error 50289
' ! 20110421 - v012 - Add aeladdinT ribbon with hookup to Command 1, 2, 3, 4,
' ! 20110421 - v013 - Test on Vista and Access 2007. It worked when code was recompiled and resulting accde addp-in installed
' ! 20110422 - v014 - Add aeRibbon with 0~9, A~F command buttons. Use 2007 as baseline for ribbon features
' ! 20110423 - v015 - Hook up Comand 1, 2, 3, 4 and test. Erl=62 in IndentCodeModule. Move callbacks to module "(_) _COMMANDS_" and test again
' ! 20110424 - v016 - Set default ribbon to aeRibbon1to4
' ! 20110425 - v017 - Add PADDING to name of modules that are not being used for aeladdin
' !                 - Ref: http://www.excelguru.ca/blog/category/the-ribbon/ => Hint about PRIVATE callbacks for Excel add-in !!!
' !                 - ControlId's for Office 2010:
' !                 - Ref: http://www.microsoft.com/downloads/details.aspx?displaylang=en&FamilyID=3f2fe784-610e-4bf1-8143-41e481993ac6
' !                 - Ref: http://www.utteraccess.com/forum/Simplifying-Ribbon-Callba-t1916108.html
' !                 - Application for Access 2003 and 2007+: Substitute missing class in reference with custom MDE
' !                 - Ref: http://accessblog.net/2010_11_01_archive.html
' !                 - Dynamic Cursor with sp_executesql
' !                 - http://accessblog.net/2010_11_01_archive.html
' !                 - Ref: http://www.smccall.demon.co.uk/Downloads.htm > Stuart McCall's Microsoft Access Pages - Downloads
' ! 20110428 - v018 - Add info and link for Michael Ciurescu with permission to use indent code
' !                 - Clean up ribbon XML to use aeNtryPoint function
' !                 - Loss of state of the global IRibbonUI Ribbon object
' !                 - Ref: http://www.rondebruin.nl/ribbonstate.htm
' !
' !

' http://www.access.qbuilt.com/html/gem_tips.html#VBEOptions



Public Function f_MyDocuments() As String
'    Ref: http://snippets.dzone.com/posts/show/7622
    
       Dim objFSO As Object
       Dim objShell As Object
       Dim objFolder As Object
       Dim objFolderItem As Object
    
       Const MY_DOCUMENTS = &H5&
    
      Set objFSO = CreateObject("Scripting.FileSystemObject")
      Set objShell = CreateObject("Shell.Application")
    
      Set objFolder = objShell.Namespace(MY_DOCUMENTS)
      Set objFolderItem = objFolder.Self
      f_MyDocuments = objFolderItem.Path
    
End Function