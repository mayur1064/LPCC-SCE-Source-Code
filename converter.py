import docx2txt #For Docx to Text Conversion
from nltk.corpus import stopwords

#Stopwords in English
stop_words = set(stopwords.words('english')) 


# Docx to text File Conversion
MY_TEXT = docx2txt.process("InputDoc.docx")


with open("InputTxt.txt", "w") as text_file:
    print(MY_TEXT, file=text_file)
    

#Removing Stopwords from the document
lst = []
file = open('InputTxt.txt', 'r')
for x in file:
    temp1 = x.split()
    for t in temp1:
        temp = t.strip()
        if temp not in lst:
            lst.append(temp)
file.close()

#Writing Final Data to text File
lst1 = []
for w in lst:
    #Removing non alphanumeric characters
    alphanumeric_filter = filter(str.isalnum, w)
    temp = "".join(alphanumeric_filter)
    if temp not in stop_words and len(temp) > 2: 
         lst1.append(temp) 

#Writing Final Data to text File       
file = open('InputTxt.txt', 'w')
for x in lst1:
    file.write(x+"\n")
file.close()