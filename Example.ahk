#include Excel2Email.ahk

msgbox, Open Excel Select Cells Press ^F7/^F8 to create EMail, Make sure outlook is running
^F7::

try XL	:= ComObjActive("Excel.Application")
catch e
{
	msgbox unable to detect MS Excel
	return
}	
try OL	:= ComObjActive("Outlook.Application")
catch e
{
	msgbox unable to detect MS Outlook
	return	
}

; Creating Setting up New Email
NewMail	:= OL.CreateItem(0)
Recipient := NewMail.Recipients.Add("Mr. xyz <xyz@abc.com>")
Recipient.Type := 1
Msg := "Hi Test,`n`nPlease check Following Data"
NewMail.Subject :=	"Test"
NewMail.Display

Excel2Email(XL.Selection,NewMail,Msg)
msgbox, Please check Email 
return

^F8::
try XL	:= ComObjActive("Excel.Application")
catch e
{
	msgbox unable to detect MS Excel
	return
}	
try OL	:= ComObjActive("Outlook.Application")
catch e
{
	msgbox unable to detect MS Outlook
	return	
}

NewMail	:= OL.CreateItem(0)
Recipient := NewMail.Recipients.Add("Mr. xyz <xyz@abc.com>")
Recipient.Type := 1
Msg := "Hi Test,`n`nPlease check Following Data"
NewMail.Subject :=	"Test"
NewMail.Display

clipsave := clipboard
XL.Selection.SpecialCells(12).copy ; copying selection excluding hidden cells
ClipWait
XLCopy2Email(NewMail,Msg)
clipboard := clipsave
msgbox, Please check Email 
return

