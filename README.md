# Draft a Word document with multiple user
This example illustrates how to draft a Word document by combining content from multiple users in Syncfusion&reg; Word processor (a.k.a.) Document editor component using bookmarks and track changes protection type.

1. Using bookmarks, we have marked the content of each author.
2. Using track changes protection type, we have stored the changes done by other users as tracked revisions in the document.

At present Syncfusion&reg; Word processor component doesn't support collaborative editing functionality. So in this example, we have bookmarked the user specific content in the master document and synced the bookmarked content once a user saves their changes.

## Demo illustration

We have added the bookmark for User1 and User2 in the master document which is in Files folder.

### Login Page

First page request the user name and password to navigate to next page. By default, we have provided "User1" for first user and password also same. Similarly, "User2" for second user and password also same.

### Choose the document to view or edit page

Following button will display:
* "View Document" button  to view the whole document.
* "Open my content" - Current user(for eg. User1) will extract the User1 content from master document and view in Document editor.
* "Open user1 or user2 content" - Another user (for eg. User2) will extract the User2 content from master document and view in Document editor. 
* Using Close icon - you can navigate to first page.

### Document editor page 

* This page will view the content in Document editor, based on button clicked in second page. Current user can make any changes in document but another user can view the document with track changes protection type. So, another user(for eg. user2) cannot accept/reject changes.
* Save button click will sync the modified content in master document.
* Close icon will navigate to previous page.
