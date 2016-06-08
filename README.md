# Outlook-Redirect

This is a short VB function I made a while ago; I probably pulled parts of it from various VB sites around the Internet, but it was so long ago now that I don't recall them well enough to give credit where credit is due. My apologies if I've unintentionally slighted anyone.


### Implementation:

I dropped this file into a VS Outlook addin project, and built a ribbon customization for both the Explorer and ReadMail windows using an appropriate icon. I've added the icon I used to the repository as `Redirect.png`.


## Usage:

	Redirect.AutoReply(RedirectionAddress as String, EmailToBeRepliedTo as MailItem, Debug as Boolean)
	

## Examples:

#### Explorer:
	
	Private Sub Redirect_Explorer_Ribbon_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles Redirect_Explorer_Ribbon_Button.Click

        Redirect.AutoReply("distribution@company.com", Globals.ThisAddIn.Application.ActiveExplorer.Selection(1), False)

    End Sub

#### ReadMail:

	 Private Sub Redirect_Ribbon_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles Redirect_Ribbon_Button.Click

        Redirect.AutoReply("distribution@company.com", Globals.ThisAddIn.Application.ActiveInspector.CurrentItem, False)

    End Sub