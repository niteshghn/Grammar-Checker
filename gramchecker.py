import ATD, re,time
from nltk.tokenize import sent_tokenize
filename = "C:\\Users\\Ayush\\AppData\\Roaming\\nltk_data\\corpora\\genesis\\speech.txt"
with open(filename) as f:
    text = f.read()

sentences= sent_tokenize(text)
#    mystring = text.replace('\n','')
leng=len(sentences)
a=[0]*leng
ATD.setDefaultKey("GranCheck#123")
metrics = ATD.stats(text)
print([str(m) for m in metrics])
for i in range(0,leng):
    time.sleep(1)
    errors = ATD.checkDocument(sentences[i])
    #metrics= ATD.stats(sentences[i])
    for error in errors:
        print ("%s error for: %s **%s**" % (error.type, error.precontext, error.string))
        print ("some suggestions: %s" % (", ".join(error.suggestions),))
        a[i]+=1
        # print(error.string,"@@@@@@")
    print (sentences[i], a[i])
print ("Total ",sum(a))
