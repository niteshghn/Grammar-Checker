import win32com.client, os

wdDoNotSaveChanges = 0
path = os.path.abspath('snippet3.txt')

snippet = 'Myself John. I studies mathematics. This is an correct sentence. '
snippet += 'You is selfish.'
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