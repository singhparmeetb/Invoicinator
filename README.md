# Invoicinator
An automated Python script to automatically import data from invoices to Excel sheets.

# Basic WorkFlow:
![Screenshot 2024-01-03 221959](https://github.com/singhparmeetb/Invoicinator/assets/80666749/3363d565-8949-4480-b493-d1e13cf9e09a)


# Different Edge Cases Identified (Along with their solutions)
**Location of PDF Invoices:** Depending on the location of the invoices we might have to change the code for automation
    
- **Online Common Storage (Gdrive/OneDrive):**
    In this case, we can simply use the shared link of the drive or perhaps we can give proper authentication permission to the script to access the files and do the pre-processing.

- **Email:** In this scenario we can use the *label* feature of mailis. We can create a filter to label emails containing PDF attachments and containing the words bills or invoice as "Not Processed". We can then access those emails using the filter and then after processing we can change the filer to "Processed".

**Location of Excel Sheet**
* **Online Common Storage (Gdrive/OneDrive):**
    Again we can provide the script with proper authentication to access these files
* **Local Storage:**
    Append the data to the local Excel sheet.

**Alternative Solution For a complete online experience:** We can rewrite the script in ***Google AppScript.*** By doing so we can easily integrate all the Google services like accessing mails, drives and sheets without the hassle of managing credential keys, which might be unsafe as if this file gets stolen, it can give uncontrolled access to anyone.

We can use the same flow of accessing the files via Google Drive (or mails based on the filter tag). We can then pre-process the data using the API and then find relevant information using regex.

Cons: The user will be limited to using only the G-Suite and working with AppScript is very unintuitive.

