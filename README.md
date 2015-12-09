# Resume-parser
Python program to analyze resume (word document) using a generated set of keywords
#Working
    1. Place both the python executable files keywords_generator.py and resume_identifier.py in the directory 
    in which the word documents that are to be operated upon are present.
    2. Execute keywords_generator.py
    3. Execute resume_identifier.py
**keywords_generator.py**
* Finds all the files in the current directory which are of types '.doc' , '.docx' , '.rtf' and parses them individually.
* Gets the common keywords from the files for software keyterms and management keyterms depending upon the input.
* Generates two word documents 'software keyterms.docx' and 'management keyterms.docx' in the same directory which 
contain the found common keyterms.

**resume_identifier.py**
* Parses 'software keyterms.docx' and 'management keyterms.docx' present in the same directory.
* parses other word documents of types '.doc' , '.docx' , '.rtf' depending upon the user request and finds the percentage 
of keyterms covered for both the management and software resume.
* identifies each resume as software or management depending upon this percentage.

#Requirements
    1. Microsoft Word.
    2. PyWin32 library.
