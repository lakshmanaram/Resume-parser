#keyword_generator.py
from time import sleep
import win32com.client as win32
import glob,os
Files = []
map_list = []
extensions = ['*.docx','*.doc','*.rtf']
#software = ['Shailender Dabodiya_BA Continuum_5.06_yrs.docx','Manisha Bhayana elhi 8.03 rs.doc']
#management = ['Nitin Bailey_Sales Manager -- India.docx','Philip Sales.docx','Matiur_rahiman786@yahoo.com.doc','PAWAN KUMAR KANDPAL _Admin.doc','SAMEER LEEKHA_ CTO.rtf']
software = []
management = []
m_maps = []
s_maps = []
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
for e in extensions:
    for infile in glob.glob( os.path.join('',e) ):
        if '~' in infile:
        #make sure that word doesn't contain ~ symbols.
        #This is to avoid the script from picking upfiles that are being
        #edited which give an exception of improper word document
            continue
        if infile in ['Management keyterms.docx','Software keyterms.docx']:
        #to remove the already existing key terms word documents if existing
            os.remove(infile)
        else:
            sleep(1)
            m = {}
            Files.append(infile)
            print 'Do you want to consider '+infile+' for generating the key terms? (Y/N) '
            choice = raw_input(':')
            if choice == 'N':
                continue
            elif choice != 'Y':
                print 'Invalid input'
                print infile + 'is being used for generating key terms'
            print 'Is '+infile+' a management resume or a software resume? (M/S)'
            category = raw_input(':')
            if category == 'S':
                software.append(infile)
            elif category == 'M':
                management.append(infile)
            else:
                print 'Invalid input'
                print infile + 'is considered to be a software resume'
                software.append(infile)
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
            if infile in software:
                s_maps.append(m)
            elif infile in management:
                m_maps.append(m)
            else:
                map_list.append(m)
            print "Done Parsing  ",infile
            doc.Close(False)
            sleep(1)
m_common = m_maps[0]
doc = word.Documents.Add()
sleep(1)
pointer = doc.Range(0,0)
for i in range(1,len(m_maps)):
    temp_map = m_maps[i]
    for k in m_common.keys():
        if k in temp_map:
            continue
        else:
            del m_common[k]
#    print "------------------------------------------------------\n\n\n\n\n\n\n\n\n\n\n\n\n"
for j in m_common.keys():
    pointer.InsertAfter(j + '\n')
#print "common management terms are: "
#for i in m_common.keys():
#    print i
doc.SaveAs(os.getcwd()+'\\Management keyterms.docx')
print 'Management common terms word document created in the present directory'
doc.Close(False)
s_common = s_maps[0]
doc = word.Documents.Add()
sleep(1)
pointer = doc.Range(0,0)
for i in range(1,len(s_maps)):
    temp_map = s_maps[i]
    for k in s_common.keys():
        if k in temp_map:
            continue
        else:
            del s_common[k]
#    print "------------------------------------------------------\n\n\n\n\n\n\n\n\n\n\n\n\n"
for j in s_common.keys():
    pointer.InsertAfter(j + '\n')
doc.SaveAs(os.getcwd()+'\\Software keyterms.docx')
print 'Software common terms word document created in the present directory'
doc.Close(False)
word.Application.Quit(-1)
print 'Key words generated'
print 'Now the script for the identification of resume can be used'
