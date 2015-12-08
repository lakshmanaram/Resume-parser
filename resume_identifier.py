from time import sleep
import win32com.client as win32
import glob,os
import sys
#Make sure that this file is in the same directory as the files that are to be identified as Software or Management resume.
Files = []
extensions = ['*.docx','*.doc','*.rtf']
software = []
management = []
m_map = {}
s_map = {}
print 'Are the resume that are to be checked in this directory: '+os.getcwd()+' (Y/N)'
if raw_input(':') == 'N':
    print 'Please paste resume_identifier.py in the same directory as the word documents that are to be identified and Try again.'
    print 'This program is being terminated.'
    sleep(5)
    sys.exit()
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
for e in extensions:
    for infile in glob.glob( os.path.join('',e) ):
        Files.append(infile)
if ('Management keyterms.docx' in Files) & ('Software keyterms.docx' in Files):
    for infile in ['Management keyterms.docx','Software keyterms.docx']:
        m = {}
        doc = word.Documents.Open(os.getcwd()+'\\'+infile)
        for each_word in doc.Words:
            w = ""
            text = each_word.Text
            for i in text:
                if ((i>='a') & (i<='z')) |((i>='A') & (i<='Z')) | (i == '-') :
                    w+=i
                else:
                    break
            if w!="":
                m[w] = 1
        if infile == 'Software keyterms.docx':
            s_map = m
        else:
            m_map = m
        print "Done Parsing  ",infile
        doc.Close(False)
        sleep(1)
    for infile in Files:
        m = {}
        if '~' in infile:
            os.remove(infile)
            continue
        if infile in ['Management keyterms.docx','Software keyterms.docx']:
            continue               
        else:
            print 'Do you wish to identify '+infile+' resume? (Y/N) '
            choice = raw_input(':')
            if choice == 'N':
                continue
            elif choice != 'Y':
                print 'Invalid input'
                print infile + ' is being identified'
            scount = 0
            mcount = 0
            doc = word.Documents.Open(os.getcwd()+'\\'+infile)
            for each_word in doc.Words:
                w = ""
                text = each_word.Text
                for i in text:
                    if ((i>='a') & (i<='z')) |((i>='A') & (i<='Z')) | (i == '-') :
                        w+=i
                    else:
                        break
                if w!="":
                    if w in m.keys():
                        m[w] += 1
                    else:
                        m[w] = 1
                        if w in s_map.keys():
                            #print 'software',w
                            scount += 1
                        if w in m_map.keys():
                            #print 'management',w
                            mcount += 1
            print "Done Parsing  ",infile
            doc.Close(False)
            sleep(1)           
            spercent = int(float(scount)/float(len(s_map.keys()))*100)
            mpercent = int(float(mcount)/float(len(m_map.keys()))*100)
            print mcount,' words from ',infile,' matches out of ',len(m_map.keys()),' management keywords which is          ',mpercent,'%'
            print scount,' words from ',infile,' matches out of ',len(s_map.keys()),' software keywords which is             ',spercent,'%'
            if spercent > mpercent:
                software.append(infile)
            else:
                management.append(infile)
else:
    print 'Software keyterms.docx Or Management keyterms.docx Not Found'
    print 'Execute keyword_generator.py before executing this program'
    print 'This program is being terminated.'
    sleep(5)
    sys.exit()
word.Application.Quit(-1)
print 'Resume identified as software resume using "software keyterms.docx" are:'
for i in software:
    print i
print 'Resume identified as management resume using "management keyterms.docx" are: '
for i in management:
    print i
sleep(10)
