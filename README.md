# Outlook VBA Macro: Download Attachments from Selected Emails

## Overview
This project provides a VBA macro for Microsoft Outlook that downloads all attachments from selected emails to a specified folder on your computer. Once set up, you can run this macro by selecting emails and clicking a button in the Outlook ribbon, which will save all attachments from those emails to a designated folder.

## Features
- Automatically downloads all attachments from selected emails.
- Can be easily assigned to a button in the Outlook ribbon for quick access.
- Saves attachments to a specified folder on your local machine.
  
## Setup Instructions

### Step 1: Download and Extract
1. Download the ZIP File.
2. Extract **emailAttachmentDownload.bas**.

### Step 3: Add the VBA Macro
1. Open **Outlook**.
2. Go to **File** > **Options** > **Customize Ribbon**.
3. Check **Developer** on the right side.
4. In the Options Menu, go to **Trust Center**, and press **Trust Center Settings**.
5. In the Trust Center Settings, go to **Macro Settings**, and **Enable All Macros**.**(DOING THIS MAY POTENTIALLY ALLOW DANGEROUS CODE TO RUN)**

### Step 3: Add the VBA Macro
1. In Outlook, press `Alt + F11` to open the **VBA Editor**.
2. In the VBA Editor, go to **File** > **Import File**.
3. Select the extracted **emailAttachmentDownload.bas**.
4. Change the **saveFolder** from the default **"C:\YourFolderPath\"** to where you want to save the attachments.
5. Save the file with `Ctrl + S` and close the VBA Editor.

### Step 4: Assign the Macro to a Button in the Ribbon
1. Right-click anywhere on the ribbon and select **Customize the Ribbon**.
2. Choose the tab (e.g., **Home**) where you want to add the button, and create a new group:
   - Select the tab (e.g., **Home**), and click **New Group**.
   - Rename the group (optional) by selecting it and clicking **Rename**.
3. From the dropdown on the left, choose **Macros**.
4. Select the macro you just created (`DownloadAttachmentsFromSelectedEmails`).
5. Click **Add** to move the macro to your new group.
6. Optionally, click **Rename** to change the button’s name and icon.
7. Click **OK** to save the changes.

## Notes
- Ensure that the folder you specify in the `saveFolder` path exists. If not, the macro may fail.
- The macro skips any non-email items in the selection.
- Ensure that no file system restrictions are in place (e.g., file permission issues or invalid file names).

## Donations
I did this for my own personal use because I absolutely hate tedious tasks. I figured if this was a problem for me, it was for others aswell. I would appreciate any donations. [Paypal](https://paypal.me/meesterbaig?country.x=US&locale.x=en_US)
