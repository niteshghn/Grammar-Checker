import win32com.client, os, re
from matplotlib import pyplot as plt
import numpy as np
from textstat.textstat import textstat

wdDoNotSaveChanges = 0
path = os.path.abspath('snippet7.txt')

filename = "C:\\Users\\Ayush\\AppData\\Roaming\\nltk_data\\corpora\\genesis\\Vasu.txt"
with open(filename) as f:
    text = f.read()
    sentences = re.split(r' *[\.\?!][\'"\)\]]* *', text)

    mystring = text.replace('\n','')
    leng=len(sentences)

snippet =text
#snippet += ' They are selfish.'
file = open(path, 'w')
file.write(snippet)
file.close()

app = win32com.client.gencache.EnsureDispatch('Word.Application')
doc = app.Documents.Open(path)
print(snippet)
gram = doc.GrammaticalErrors.Count
#senti=doc.GetSpellingSuggestions.Count

avg_gram = gram/leng
print(gram)

app.Quit(wdDoNotSaveChanges)
#os.remove(path)
def plot1 (val):
    print ('''felsch easy reading score is as folows
            * 90-100 : Very Easy
            * 80-89 : Easy
            * 70-79 : Fairly Easy
            * 60-69 : Standard
            * 50-59 : Fairly Difficult
            * 30-49 : Difficult
            * 0-29 : Very Confusing''')

    A = [[30,],[40],[50],[60],[70],[80],[90],[100]]

    X = range(1)
    width= 5
    plt.figure(1)
    p1= plt.bar(X, A[0][0], color = '0.25', width=width )
    p2= plt.bar(X, A[1][0], color = 'w', bottom = A[0][0],width=width)
    p3= plt.bar(X, A[2][0], color = 'c', bottom = A[1][0],width=width)
    p4= plt.bar(X, A[3][0], color = 'g', bottom = A[2][0],width=width)
    p5= plt.bar(X, A[4][0], color = 'r', bottom = A[3][0],width=width)
    p6= plt.bar(X, A[5][0], color = 'b', bottom = A[4][0],width=width)
    p7= plt.bar(X, A[6][0], color = 'm', bottom = A[5][0],width=width)
    p8= plt.bar(X, A[7][0], color = 'k', bottom = A[6][0],width=width)
    t = np.arange(0.,5.,0.2)
    p0= plt.plot(t,(0*t + val), 'ys')
    plt.yticks(np.arange(0, 100, 5))
    plt.ylim(0., 100.0)
    plt.legend((p0[0],p1[0],p2[0],p3[0],p4[0],p5[0],p6[0],p7[0],p8[0]), ('Your Score', 'Very Confusing', 'Difficult','Fairly Difficult', 'Standard','Fairly Easy',"Easy", "Very Easy"))
    plt.title("The Flesch Readng Ease Scores")

def plot2(smog):
    t = np.arange(0.,5.,0.2)
    plt.figure(2)
    p10= plt.plot(t,(0*t + smog), 'r^')



def run_analysis( test_data ):
    print ("Your Flesch Reading Ease score is : ",textstat.flesch_reading_ease(test_data))
    val = textstat.flesch_reading_ease(test_data)
    plot1(val)
    print ("The smog index: ", textstat.smog_index(test_data))
    smog=textstat.smog_index(test_data)
    #plot2(smog)
    print ("The flesch kincaid grade",textstat.flesch_kincaid_grade(test_data))
    print ("Coleman liau index",textstat.coleman_liau_index(test_data))
    print ("Automated Readability Index",textstat.automated_readability_index(test_data))
    print ("Dale Chall Readability Score",textstat.dale_chall_readability_score(test_data))
    print ("Difficult Words",textstat.difficult_words(test_data))
    print ("Linsear write formula",textstat.linsear_write_formula(test_data))
    print ("gunning fog", textstat.gunning_fog(test_data))
    plt.show()
run_analysis(text)