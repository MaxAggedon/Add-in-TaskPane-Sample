#Office Add-in Task Pane

>**Note:**  We will be removing this sample from the site on December 15, 2016. If you’d like to keep a copy of this sample for your own reference, please download or clone the repo.

**Table of contents**

* [Run in the playground for Office Add-ins](#run-in-the-playground-for-office-add-ins)
* [Run in desktop version of Office](#run-in-desktop-version-of-Office)
* [Run in Office Online](#run-in-office-online)
* [More resources](#more-resources)

If you are getting started building your first Office Add-in, this sample will help you out. 
This sample is an Office Add-in that gets selected data from a document and displays it. It runs
as a task pane add-in that works in Word, Excel, and PowerPoint.

##Run in the playground for Office Add-ins
To get running quickly:

1. Go to the [playground for Office Add-ins](http://aka.ms/Fb7hmn). The sample will open automatically.
2. Choose the **Run Project** button to run the sample in Excel Online.
3. Log in using a Microsoft account.

**Note:** The playground will only run the add-in in Excel Online. You will need to sign in to a Microsoft account.

##Run in desktop version of Office
Follow the steps below to get the sample running in the desktop version of Office 2013. 

1. Host the files on a local network share.
2. Open Excel, Word, or PowerPoint and open a document.
3. Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs** (this is named **Trusted App Catalogs** in Office 2013).
4. Enter the location of the local network share in the **Catalog Url** text field, and choose **Add Catalog**. Make sure the **Show in Menu** check box is selected.
5. Choose **OK** and then **OK** again to close the dialog boxes. 
6. Close the Office application and start it again so that the changes take effect.
7. Open a document.
8. Choose **Insert** > **Add-ins** > **My Add-ins**, and then choose **SHARED FOLDER**. Note: On Office 2013, choose **Insert** > **My Apps** > **SHARED FOLDER**.
9. Select **My First Taskpane Add-in**. If you don't see the add-in, choose **Refresh**.
10. Choose **Insert** and the task pane add-in will open next to your document.


##Run in Office Online
Follow the steps below to get the sample running in Office Online. You will need a subscription to Office 365. 
If you don't have one, [join the Office 365 Developer Program and get a free 1 year subscription to Office 365](https://aka.ms/devprogramsignup).

1. Host these files locally (on localhost) or online (e.g. AWS, Azure, Heroku, etc). In the manifest.xml file, 
change the **DefaultValue** of the **SourceLocation** to point to the URL where the index.html file is hosted.
2. Go to the [Office 365 portal](https://portal.office.com) and choose **Admin** from the app launcher in the top-left corner.
3. Choose **SharePoint**. This will take you to the SharePoint admin center.
4. Choose **apps** > **App Catalog** > **Apps for Office**.
5. Choose the **Upload** button, and choose the manifest.xml file from your local directory.
6. Choose **OK** and the add-in will install.
7. Open the app launcher in the top-left corner and select an Office application (Excel or Word).
8. Open a document.
9. Choose **Insert** > **Office Add-ins** and then choose **MY ORGANIZATION**.
10. Select **My First Taskpane Add-in**. If you don't see the add-in, choose **Refresh**.
11. Choose **Insert** and the task pane add-in will open next to your document.

##More Resources
* [Office Dev Center](http://dev.office.com/)
* More [code samples](http://dev.office.com/codesamples)
* [Create a network shared folder catalog for task pane and content add-ins](https://msdn.microsoft.com/EN-US/library/office/fp123503.aspx)
* [Publish task pane and content add-ins to an add-in catalog on SharePoint](https://msdn.microsoft.com/EN-US/library/office/fp123517.aspx)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
