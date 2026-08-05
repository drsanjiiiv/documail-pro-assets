# DocuMail Pro Support 🛠️

Thank you for using DocuMail Pro! We want to ensure your document automation and mail merge workflows run seamlessly. Because our application runs entirely inside your Google Workspace environment, we track all feedback and system bugs publicly on GitHub.

---

## 🔐 Managing & Deleting Your Data

DocuMail Pro does not store your data. All recipient data, template content, and merge results live inside your own Google Workspace files (your Google Sheet and the documents/PDFs generated into the destination folder you choose).

**To remove DocuMail Pro data:**

1. **Merge/Email status:** Clear the DocuMail Pro status columns in your sheet ("Merged Doc Status", "Merged Doc ID", "Merged Doc URL", and any "Sent Mail Status - <Template Name>" columns). You can also delete those columns entirely.
2. **Saved template configurations:** Use the delete option on the template card in the DocuMail Pro sidebar, or remove the add-on to clear the spreadsheet-scoped configuration.
3. **Generated files:** Delete any generated documents/PDFs from the destination folder in your Google Drive.
4. **Scheduled runs:** Deleting a template automatically removes its time-driven trigger. You can also review and remove triggers at any time under **Extensions > Apps Script > Triggers** in your spreadsheet.
5. **Remove the add-on entirely:** Right-click the add-on icon in Google Sheets (or Docs) and choose "Remove", then revoke access under your Google Account permissions (myaccount.google.com > Security > Third-party access).

Once the add-on is removed and its status columns are cleared, no trace of DocuMail Pro processing remains in your files.

---

## 🐛 How to Report an Issue or Bug

If a template isn't running correctly, or if you encounter a system error, please follow these steps to log an issue:

1. Click on the **Issues** tab at the top of this GitHub repository page.
2. Click the green **New Issue** button.
3. Provide a clear title (e.g., *"Preview button not loading layout"*).
4. Describe the problem, including any error messages you see in your sidebar.
5. Click **Submit new issue**.

---

## 🔒 Data & Privacy Note
When reporting bugs, please **never** paste screenshots or logs that contain sensitive or personal business data (such as client email addresses or private spreadsheet data).
