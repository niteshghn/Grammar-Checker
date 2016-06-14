import win32com.client, os, re

wdDoNotSaveChanges = 0
path = os.path.abspath('snippet7.txt')

filename = "C:\\Users\\Ayush\\AppData\\Roaming\\nltk_data\\corpora\\genesis\\speech.txt"
with open(filename) as f:
    text = f.read()
    sentences = re.split(r' *[\.\?!][\'"\)\]]* *', text)
    mystring = text.replace('\n','')
#print(mystring)

snippet =text
#snippet += 'You is selfish.'
file = open(path, 'w')
file.write(snippet)
file.close()

app = win32com.client.gencache.EnsureDispatch('Word.Application')
doc = app.Documents.Open(path)
print(snippet)
print ("Grammar: %d" % (doc.GrammaticalErrors.Count,))
print ("Spelling: %d" % (doc.SpellingErrors.Count,))

app.Quit(wdDoNotSaveChanges)
#os.remove(path)