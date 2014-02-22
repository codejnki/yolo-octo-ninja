yolo-octo-ninja
===============

My wife messaged me one day saying that every time they e-mail a customer they have to save a copy of the message and upload it in to their CRM software.  She asked me if I could write an Outlook macro to do the file saving.  I wrote the macro but ended up having issues with macro security.  So I decided to try and see if I could rewrite it as an Outlook Add In.  This is the results.

Outlook Add-in to automatically save e-mail

Every time an e-mail is sent via Outlook, a copy of the messages in .msg format will be saved in the user's My Documents\MailSave folder.  The add in will create a new folder per day to help keep things organized.  The files are named as the first recipient's email address and a time stamp.

There are no options or settings, it all just happens in the background.
