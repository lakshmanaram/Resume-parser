import win32com.client as win32
import glob,os
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
Files = []
for infile in glob.glob( os.path.join('', '*.rtf') ):
    Files.append(infile)
    doc = word.Documents.Open(os.getcwd()+'\\'+infile)
    for each_word in doc.Words:
        w = ""
        text = each_word.Text
        for i in text:
            if ((i>='a') & (i<='z')) |((i>='A') & (i<='Z')) | (i == '-') :
                w+=i
            else:
                break
        print w
word.Application.Quit(-1)
