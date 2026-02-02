# Office Add-in Challenge with Redaction and Confidential Header

This Word add-in adds a taskpane with a Redact document button that:
- Redacts sensitive info in the document:
  - Emails, Phones, SSNs, Credit Cards, Employee IDs, MRNs, INSs
- Inserts a CONFIDENTIAL DOCUMENT in the document header (with a fallback banner if headers aren’t available).
- Enables Track Changes when WordApi 1.5+ is supported.

---

## Run

From the project root folder run the following 
```bash
npm install
```
then only when running for the first time run this command 
```bash
npm run install-dev-certs
```
then run 
```bash
npm start
```