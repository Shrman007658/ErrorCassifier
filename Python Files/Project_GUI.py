import tkinter as tk
from tkinter import *
import pandas as pd
import os
import subprocess
from pandas import ExcelWriter
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfTransformer
from sklearn.feature_extraction.text import CountVectorizer
from sklearn import metrics
from sklearn.metrics import accuracy_score
from tkinter import filedialog
import nltk
from nltk.corpus import stopwords 
from nltk.tokenize import word_tokenize
from nltk.stem.wordnet import WordNetLemmatizer 
import re
import time
import warnings
from sklearn.externals import joblib
from sklearn.ensemble import RandomForestClassifier
from sklearn.svm import SVC

#nltk.download('wordnet')
rf = RandomForestClassifier(n_estimators = 1500, random_state = 12)
joblib_file="joblib_classifier_model.pkl"
warnings.filterwarnings('ignore')
path=''
fpath=''
lem=WordNetLemmatizer()
cnt_vectorizer = CountVectorizer()
transtfidf=TfidfTransformer()
X_train=[]
X_test=[]
y_train=[]
y_test=[]
root=tk.Tk()
root.title("Organizer")
root.geometry("750x500")
C = tk.Canvas(bg='#7DCEA0',bd=20,height=750, width=750).place(x=0,y=0)
df = pd.DataFrame()
def preprocessor(Title,dataframe):
    l=len(Title)
    for x in range(0,l):
        stop_words = stopwords.words('english')
        stop_words.remove('not')
        lw=Title[x].lower()
        #removing special character    
        sub = re.sub(r'\W', ' ', str(lw))
        
        #tokenization
        tokenized_word=word_tokenize(sub) 
        filtered_word=[]
        for w in tokenized_word:
            if w not in stop_words:
                filtered_word.append(w)
        #lemmatization  
        lemmed_word=[]
        for w in filtered_word:
            lemmed_word.append(lem.lemmatize(w,"v"))
        #print("Lematized_Sentence:",lemmed_word)
        s=' '.join(lemmed_word)
        dataframe['Title'][x]=s

def FileSelect():
    try:
        global path
        StatusLabel.configure(bg='#E1E809',text="Status:Stemming..Lematizing...Tossing..Turning..Please Wait.")
        path=filedialog.askopenfilename(filetypes=(("Template files","*.xlsx"),("All files","*")))
        global df
        df=pd.read_excel(path)
        preprocessor(df['Title'],df)
        StatusLabel.configure(bg='#51D406',text="Status:File Path Configured...")
        df.Title.fillna(df.Title.dropna().max(),inplace =True)
        df.Category.fillna(df.Category.dropna().max(),inplace =True)
        root.mainloop()
    except Exception as e:
        StatusLabel.configure(bg='#FF0000',text=e)
        root.mainloop()

def SplitData():
    value=float(splt.get())
    try:
            global df
            global X_train,X_test,y_test,y_train
            X=df['Title']
            y=df['Category']
            X_train,X_test,y_train,y_test=train_test_split(X,y,test_size=value,random_state=1)
            testStatus.configure(bg='#51D406',text="Splitted and Ready to work☺")
            root.mainloop()
    except Exception as e:
            testStatus.configure(bg='#FF0000',text=e)
            root.mainloop()

def TrainMachine():
        try:

                
                global X_train,X_test,y_train,y_test
                X_train_tf=cnt_vectorizer.fit_transform(X_train)
                X_train_tfidf=transtfidf.fit_transform(X_train_tf)
                rf.fit(X_train_tfidf,y_train)
                trainStatus.configure(bg='#51D406',text="The model is trained :)")
                root.mainloop()
        except Exception as e:
                trainStatus.configure(bg='#FF0000',text=e)
                root.mainloop()

def TestMachine():
        try:
                global X_train,X_test,y_train,y_test
                X_test_tf=cnt_vectorizer.transform(X_test)
                X_test_tfidf=transtfidf.transform(X_test_tf)
                predictedTest=rf.predict(X_test_tfidf)
                acc=accuracy_score(y_test,predictedTest)
                res="The accuracy is "+str(acc)
                MTestStatus.configure(bg='#51D406',text=res)
                root.mainloop()
        except Exception as e:
                MTestStatus.configure(bg='#FF0000',text=e)
                root.mainloop()

def FinTrain():
        try:
                tottrainStatus.configure(bg='#E1E809',text="Status:Stemming..Lematizing...Tossing..Turning..Please Wait.")
                trainfile=filedialog.askopenfilename(filetypes=(("Template files","*.xlsx"),("All files","*")))
                if path==trainfile:
                        dft=df
                        time.sleep(1)
                        tottrainStatus.configure(bg='#51D406',text="Status:Identified same file as evaluation.Initialization skipped.")
                else:
                        dft=pd.read_excel(trainfile)
                        preprocessor(dft['Title'],dft)
                        tottrainStatus.configure(bg='#51D406',text="Status:File Configured...")
                        dft.Title.fillna(dft.Title.dropna().max(),inplace =True)
                        dft.Category.fillna(dft.Category.dropna().max(),inplace =True)
                        
                X_tot=dft['Title']
                y_tot=dft['Category']
                X_train_tftot=cnt_vectorizer.fit_transform(X_tot)
                X_train_tfidftot=transtfidf.fit_transform(X_train_tftot)
                rf.fit(X_train_tfidftot,y_tot)
                status=tk.Label(root,bg='#51D406',text="The final model has been trained").place(x=350+40,y=275)
                joblib.dump(rf,joblib_file)
                root.mainloop()
        except Exception as e:
                tottrainStatus.configure(bg='#FF0000',text=e)
                root.mainloop()

def Findxl():
        try:
                global df,fpath
                fpath=filedialog.askopenfilename(filetypes=(("Template files","*.xlsx"),("All files","*")))
                df=pd.read_excel(fpath)
                tarbutStatus.configure(bg='#51D406',text="File Locked and Loaded")
        except Exception as e:
                tarbutStatus.configure(bg='#FF0000',text=e)

def PutXl():
        try:
                global df
                #load stored model
                model=joblib.load(joblib_file)
                X_xl=df['Title']
                df['Pred_category']=0#initializing null column
                X_xl_vect=cnt_vectorizer.transform(X_xl)
                X_xl_tfidf=transtfidf.transform(X_xl_vect)
                df['Pred_category']=model.predict(X_xl_tfidf)
                engine='xlsxwriter'
                #writer=ExcelWriter(fpath,engine=engine)
                df.to_excel("output.xlsx")
                c=os.getcwd()
                c=c+'\\output.xlsx'
                os.startfile(c)
                #os.open(c)
                #subprocess.Popen(r'explorer /select,"C:\\Users\SHRAMAN\\output.xlsx"')
                predbutStatus.configure(bg='#51D406',text="Task Completed!")
                root.mainloop()
        except Exception as e:
                predbutStatus.configure(bg='#FF0000',text=e)
                root.mainloop()


NameLabel=tk.Label(root,text="Enter the file for training the model")
NameLabel.pack()
NameLabel.place(x=0+40,y=20)


#CREATING A BROWSE FOR FIILE PATH FOR MACHINE TRAIN

#adding a finalize button with the entry to store the path
browse=tk.Button(root,bg='grey', text = "Browse", command =FileSelect, width = 10)
browse.place(x=200+40,y=20)
StatusLabel=tk.Label(root,bg='#E1E809',text="Status:Waiting.....")
StatusLabel.place(x=200+40,y=55)


#TRAIN_TEST RATIO PART
TrainTest=tk.Label(root,text="Enter the train test ratio for the split:")
TrainTest.place(x=0+40,y=87)
splt=tk.Entry(root,width=25)
splt.grid(column=0,row=1)
splt.place(x=200+40,y=85)
test=tk.Button(root,bg='grey',text="Split Data",command=SplitData, width=10)
test.place(x=360+40,y=83)
testStatus=tk.Label(root,bg='#E1E809',text="Status:Waiting")
testStatus.place(x=200+40,y=110)

#TRAINING AND TESTING THE MODEL
train=tk.Button(root,bg='grey',text="TRAIN MODEL",command=TrainMachine,width=20)
train.place(x=50+40,y=140)
trainStatus=tk.Label(root,bg='#E1E809',text="TrainStatus:Waiting..")
trainStatus.place(x=270+40,y=140)
test=tk.Button(root,bg='grey',text="TEST MODEL",width=20,command=TestMachine)
test.place(x=50+40,y=170)
MTestStatus=tk.Label(root,bg='#E1E809',text="TestStatus:Waiting...")
MTestStatus.place(x=270+40,y=170)

##TRAINING FINAL MODEL FOR EXECUTION
chlabel=tk.Label(root,text="Please Select the File to Train Final Machine")
chlabel.place(x=0+40,y=250)
tottrain=tk.Button(root,bg='grey',text="Browse File",command=FinTrain,width=10)
tottrain.place(x=250+40,y=248)
tottrainStatus=tk.Label(root,bg='#E1E809',text="Status:Waiting for File..")
tottrainStatus.place(x=345+40,y=249)

#INPUT FOR TARGET EXCEL FILE
tarfile=tk.Label(root,text="Select the target excel file for model to predict")
tarfile.place(x=0+40,y=310)
tarbut=tk.Button(root,bg='grey',text='Browse File',command=Findxl,width=10)
tarbut.place(x=250+40,y=310)
tarbutStatus=tk.Label(root,bg='#E1E809',text='Status:Waiting for file..')
tarbutStatus.place(x=345+40,y=310)

##ADD BUTTON FOR PREDICTING FOR EXCEL

predbut=tk.Button(root,bg='grey',text="Predict Values To File",command=PutXl,width=20)
predbut.place(x=250+40,y=350)
predbutStatus=tk.Label(root,bg='#E1E809',text="Status: Waiting..")
predbutStatus.place(x=250+40,y=390)
root.mainloop()

