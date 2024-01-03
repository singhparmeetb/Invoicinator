# Invoicinator
An automated python script to automatically import data from invoices to excel sheets.

# Different Edge Cases Identified (Along with thier solutions)
**Location of PDF Invoices:** Depending on the location of the invoices we might have to change the code for automation
    
- **Online Common Storage (Gdrive/OneDrive):**
    In this case we can simply use the shared link of the drive or perhaps we can give proper authentication permission to the script to access the files and do the pre-processing.

- **Email:** In this scenario we can use the *label* feature of mailis. We can create a filter to label emails containing pdf attachments and containing the words bills or invoice as "Not Processed". We can then access those emails using the filter and then after processing we can simply change the filer to "Processed".

**Location of Excel Sheet**
* **Online Common Storage (Gdrive/OneDrive):**
    Again we can simply provide the script with proper authentication to access these files
* **Local Storage:**
    Simply append the data to the local excel sheet.

**Alternative Solution For a comoplete online experience:** We can simply rewrite this whole script in ***Google AppScript.*** By doing so we can easily integrate all the google services like accessing mails, drive and sheets without the hassle of managing credential keys, which might be unsafe, as if this file gets stolen, it can give uncontrolled access to anyone.

We can simply use the same flow of accessing the files via Google-drive (or mails based on the filer tag). We can then preprocess the data using the API and then find the relevant information using regex.

Possilbe cons: The user will be limited to use only the G-Suite and working with AppScript is very unintuitive.
