# Overview  

Streamline is a Microsoft Word VBA script for filling out template documents. 

By filling out a brief configuration file, Streamline can quickly...  

- Open specific templates for the task at hand
- Perform multiple find-and-replaces
- Format and fill tables
- Add conditional text
- Copy to clipboard
- Save and export to pdf

Streamline can also do many small quality of life formats like:

- Fix encoding issues (change '\&amp;' to '&' etc.)
- Add states to addresses
- Format currencies with commas
- Space out phone numbers for readability
- Capitalise sentences
- Sum related amounts

This is a generalised version of the software. Developed for my job where i found i was filling out the same templates many times everyday. Code that handles internal procedures has been removed.  


---

# Installation

1. Open "Streamline.dotm"
2. If shown a Security Warning banner, click 'Enable Content'
3. Open "instruction_file.txt" in a notepad.

* If you would like to use the modules and classes on their own, a few extra runtime libraries are required. See requirements.txt for details.

---

# Getting Started

Ensure both Streamline.dotm and instruction_file.txt are open
In Streamline.dotm, Navigate to View > Macros > View Macros

To copy notes to clipboard, Select Complete_Notes and click Run
A description and a detailed note will be copied to clipboard successively.

To create a letter, select Complete_Letters and click Run
A complete letter as both a .docx and .pdf will be saved in the "Output" folder.

To change the content of the letters and notes, make changes to instruction_file.txt with the schema described below.

---

# Schema

The layout used in instruction_file.txt

### (in) inbound
client number
- The client's unique identifier
- must be 8 digits long

date received
- in the form DD-MM-YYYY


### (adr) address
full name
company/person
- either a 'c' or a 'p'
- this affects whether the greeting will be 'Hello <Name>', or 'To Whom It May Concern'

address (multi-line)
- This can be as many lines as required
- The macro will add the state abbreviation on the last line, based on the postcode


### (res) reason
reason (multi-line)
- This can be as many lines as required
- This is the client's reasons from previous correspondence paraphrased.
- Common encoding issues are replaced here.


### (acc) account
acc name
- An account name, must include a number separated by a space at the end
- for example: 'Some Account 3'
- shorthand can also be used for common accounts
- In this version two abreviations are available
- 'exp' for Expenses Account 1
- 'rec' for Receivables Account 2
- the account number for the abbreviations can be changed with a dot, for example exp.3

value
- any number representing a monetary amount
- $ and commas should not be included, they are added by the macro.

payment number
- any number

*Multiple accounts can be included by repeating account name, value and payment number in order.*


### (out) outbound
outbound code
- determines whether to approve or deny the client
- 'y' for approved
- 'n' for denied
- 'ya' for escalation required for approval
- 'na' for escalation required for denial

outbound flags (multi)
- This can be as many lines as required, or none
- Key words to describe special details about the client. In this version of the software, only one is included though.
- include 'details' to prompt the client to update their details.


### (not) note
notes (multi-line)
- This can be as many lines as required
- These are notes from the user to add any extra information to the notes copied.
- Common encoding issues are replaced here
- functions similar to <res>


### (rpy) reply
job name
- The name that the file should be saved as
- Ensure that it is a valid filename
- Any existing file with the same name in the 'Output' folder will be overwritten with no prompt.
- suffixes like .docx or .pdf should not be included, they are added by the macro.

reference number
- A reference number included in the letter header
- It should be numeric.

