**Policy Compliance Checker**
Automated tool to check e-commerce websites for required policies and compliance.

**Quick Start**

1. Install Requirements

Install Node.js (version 14 or higher)

2. Setup Project

# Open terminal in project folder and run:
npm install
npx playwright install chromium

3. Create Input File

Create input.xlsx in the project folder with these columns:

Website	                  MerchantType	     EntityType	   LegalName	      Email
https://example.com 	     goods	         proprietor	    John Doe	merchant@example.com


4. Run Checker

node index.js


**What Gets Checked**

For Goods Merchants:

Privacy Policy
Terms & Conditions
Shipping Policy
Returns Policy
Refund Policy
Cancellation Policy
Legal Name (if proprietor)
For Services Merchants:

Privacy Policy
Terms & Conditions
Refund Policy
Cancellation Policy
Legal Name (if proprietor)


**Output Files**

After running, you'll get:

output.xlsx - Complete compliance report
emails/ folder - Email drafts for merchants with missing requirements




**Troubleshooting**

Common Issues:

bash
# If "Cannot find module" error:
npm install

# If "Playwright browsers not installed":
npx playwright install chromium

# If Excel file not found:
# - Make sure input.xlsx is in same folder as index.js
# - Check column names match exactly


**Output File Results Guide**

Understanding the Output Excel

The output.xlsx file contains all results from your compliance check. Here's what each column means:

Key Columns Explained

ComplianceStatus: Overall result. PASS = all requirements met, FAIL = something missing.

MissingPolicies: List of what's missing. Can include policy names (privacy, terms, shipping, returns, refund, cancellation) and/or legal name.

LegalNamePresent: Shows if legal name was found on website:

true = Found
false = Not found (proprietorships only)
NOT RELEVANT = Not a proprietorship
IsProprietorship: Whether business is a proprietorship (true) or not (false).

DeterminedMerchantType: What the code used for checking:

goods = If "good" appeared in your input
services = If no "good" in input
ManualCheckingRequired:

YES = Crawler found NO policies at all (needs human review)
NO = At least some policies found
Error: Any errors during checking. Blank means no errors.

Individual Policy Columns

For each policy type (privacy, terms, shipping, returns, refund, cancellation):

{Policy}Status: FOUND, MISSING, or NOT RELEVANT

{Policy}URL: If found, shows exact webpage URL where policy is located

Legal Name Columns

LegalNameStatus: FOUND, MISSING, or NOT RELEVANT

LegalNameURL: If legal name found, shows where it was located

Email Columns

EmailGenerated: YES if email created, NO if not

EmailFilename: Name of email file in emails/ folder (if generated)

How to Interpret Results

PASS Website

ComplianceStatus: PASS
MissingPolicies: (empty)
All relevant policies: FOUND
EmailGenerated: NO
Action: No action needed. Merchant is compliant.

FAIL Website with Specific Issues

ComplianceStatus: FAIL
MissingPolicies: Lists missing items
Some policies: MISSING
EmailGenerated: YES (usually)
Action: Check emails/ folder for draft email to send merchant.

Manual Checking Needed

ComplianceStatus: FAIL
ManualCheckingRequired: YES
All policies: MISSING
EmailGenerated: NO
Action: Human must manually check website. Crawler found nothing.

Non-Proprietorship

IsProprietorship: false
LegalNamePresent: NOT RELEVANT
Legal name not in MissingPolicies
Action: Only check policy compliance.

Common Scenarios

New Store: Has basic policies (privacy, terms) but missing shipping/returns
Services Business: Has required 4 policies, shipping/returns show NOT RELEVANT
Proprietor Missing Name: All policies found but legal name missing → FAIL
Goods vs Services: If input says goods_services, code treats as goods
Using Results

PASS: Activate merchant
FAIL + EmailGenerated=YES: Send email, wait for fixes
ManualCheckingRequired=YES: Investigate manually
Error present: Check website accessibility
