# Outlook-Redirect

This is a short VB function I made a while ago; I probably pulled parts of it from various VB sites around the Internet, but it was so long ago now that I don't recall them well enough to give credit where credit is due. My apologies if I've unintentionally slighted anyone.


### Implementation:

The code can be deployed as either a macro or as an add-in. There are slight code changes between the two; the add-in version is `Outlook-Redirect.vb`, whereas the macro version is `Outlook-Redirect.macro`.


#### As an Add-In:

Note: this approach uses Ribbon cutomizations, which is a feature of MS Office 2007 and higher. To the best of my knowledge, add-ins should still work in MS Office 2003, although investigation would be required. As far as I know, add-ins do not work in Office XP or earlier.
Using an add-in should simplify larger deployments.

* Create a new Visual Studio Outlook addin project. You may require the [Visual Studio Tools for Office add-in](https://msdn.microsoft.com/en-us/library/d2tx7z6d.aspx).

* Add the `Outlook-Redirect.vb` file to the project.

* Build ribbon customizations as required. `Redirect.png` is the icon that I used.

* Add the `Redirect.AutoReply` function to the Click events of the new ribbons. See "Examples" for, well, examples.

* Compile, build, save, deploy, etc.


#### As a Macro:

Note: this approach uses Visual Basic for Applications, so *should* be compatible with most versions of MS Outlook. However, it also requires more work to deploy on a large scale, so may be better suited to individuals.


* Open the VBA interface of Outlook (Alt-F11).

* Create a new module.

* Insert the contents of `Outlook-Redirect.macro` into the new module.

* Create a new Public Sub that calls `AutoReply` with the appropriate parameters. See "Usage" and "Examples" for more information.

* Save changes and return to Outlook.

* Customize the Ribbon.

* Change the dropdown from "Popular Commands..." to "Macros".

* Select a custom group, or create on if a suitable one does not already exist.

* Select the new created sub, and click "Add > >".

* Select the newly added macro (in the right-hand list) and click "Rename" to change the displayed name and icon.


### Usage:

	Redirect.AutoReply(RedirectionAddress as String, EmailToBeRepliedTo as MailItem, (Optional) Debug as Boolean)
	

### Examples:

Note: If using the macro version, remove `Globals.ThisAddin` from the function call.

#### Explorer:
	
	Private Sub Redirect_Explorer_Ribbon_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles Redirect_Explorer_Ribbon_Button.Click

        Redirect.AutoReply("distribution@company.com", Globals.ThisAddIn.Application.ActiveExplorer.Selection(1), False)

    End Sub

#### ReadMail:

	 Private Sub Redirect_Ribbon_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles Redirect_Ribbon_Button.Click

        Redirect.AutoReply("distribution@company.com", Globals.ThisAddIn.Application.ActiveInspector.CurrentItem, False)

    End Sub