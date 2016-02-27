# -*- coding: utf-8 -*-
"""
Created on Tue Oct 27 15:36:10 2015

@author: Aska
"""
def main():
    #JobDesc1211.xlsx
    #r1df_5 = r1df.loc[~r1df['words'].isin(['', 'asd'])]
    import pandas as pd
    import progressbar
    #import nltk
    from time import sleep
    #from nltk.probability import FreqDist
    from nltk.tokenize import word_tokenize
    from time import time
    import numpy as np
    import os
    import sys
    import unicodedata
    
    def Word_list2(A,k):
        if k<=len(A):
            A2=[]
            for w in range(len(A)-(k-1)):
                w2=''        
                t=0
                while t<=(k-1):
                    w2=w2+' '+A[w+t]
                    t=t+1
                A2.append(w2)
            return A2
        else:
            return A
    
    def Word_list(A,k):
        A2=[]
        punc=[':','.',',','!',';','?','{','}',']','[','(',')','"',"'"]
        for sett in A:
            for sett2 in sett:
                A2.append(sett2)
        
        if k==1:
            A3=[]
            for w in A2:
                if w not in punc:
                    A3.append(w.lower())
            return A3
        elif k>=2 and k<=len(A2):
            A3=[]
            
            for sen in A:
                for w in range(len(sen)-(k-1)):
                    nope=0
                    w2=sen[w]  
                    if sen[w] in punc:
                        nope=1
                    t=1
                    while t<=(k-1) and nope==0:
                        if sen[w+t] in punc :
                            nope=1
                        w2=w2+' '+sen[w+t]
                        t=t+1
                    if nope==0:
                        A3.append(w2.lower())
            return A3
        else:
            if k==0:
                print 'Error...'
            else:
                print 'Error.. Too much Word'
        
    
    Mas=1
    while Mas==1:
        Mas=0
        print 'Masukan File,, '
        fold='D:\Keyword_Density\Filee'
        no=1
        FNm=[]
        for fname in os.listdir(fold):
            print no,'. ',fname
            FNm.append(fname)
            no=no+1
        
        File=0
        while File==0:
            try:
                File=input('\nPilih File :')
                if File>=no+1 or type(File)!=int or File<=0:
                    print '\nERROR.. Tak ada di pilihan. Coba lagi'
                    File=0
            except :
                print '\nERROR.. Harus angka. Coba lagi'
                File=0
        
        Finy=FNm[File-1]
        path='D:\Keyword_Density\Filee\\'+Finy
        try:
            r=pd.read_excel(path)
        except:
            Mas=1
            print 'Error.. File is not in excel extention'
            sleep(1.0)
            print '='*15
            print 'Coba Lagi'
            print '='*15

    yakin='n'
    while yakin=='n' or yakin == 'N':
        try:
            yakin='y'
            numb=input('\nSampai berapa kata :')
            if numb>=5:
                print 'Mana ada kata berpadan sampai sebanyak ', numb, '(-__-)'
                yakin=raw_input('Yakin masih mau lanjut (Y/N)??')
        except:
            print '\nERROR.. Harus angka. Coba lagi'
            yakin='N'
    #try:
    print 'Processing..'    
    
    r_test1=r.dropna()
    r_lis2=list(r_test1.reset_index().values)
    r_m2=np.array(r_lis2)
    r_m_t2=np.transpose(r_m2)
    r22=list(r_m_t2[1])
    
    print '\n Tokenizing..'
    
    tw=[]
    
    for w in r22:
        try :
            tw.append(word_tokenize(str(w)))
        except:
            print w
            print type(w)
            #u=unicode(w, "utf-8")
            u=unicodedata.normalize('NFKD', w).encode('ascii','ignore')
            print ' ',u
            try:
                tw.append(word_tokenize(str(u)))
            except:
                print 'Masih tak bisa'
                    
                

    #Cleansing
    print "\n Cleansiing.. Clean everything between '<' and '>'"
    bar = progressbar.ProgressBar()
    tw2=[]
    for i in bar(range(len(tw))):
        sleep(0.0000000000000001)
        ww=[]
        m=0
        for w in tw[i]:
            if w=='<':
                m=1
            elif w=='>':
                m=0
            
            if m==0 and w not in ['/','-','*','<','>','\\','|','+','_','~','`','+','=','%','#','^','nan']:                         
                ww.append(w)
        tw2.append(ww)
    
    AllWords=tw2
    p_wo='D:\Keyword_Density\Filter.xlsx'
    wo=pd.read_excel(p_wo)
    wo2=list(wo.reset_index().values)
    wo=[]
    for i in wo2:
        wo.append(i[1])
        
            
    writer = pd.ExcelWriter('D:\Keyword_Density\File_Result\\'+Finy[:len(Finy)-5]+'_result.xlsx')
    
    w=1
    print '\n Counting..'
    while w<=numb:
        if w==1:
            AllWords2=Word_list(AllWords,w)
            time1=time()
            df_words=pd.DataFrame(pd.Series(AllWords2),columns=[('words')])
            df_words=df_words.loc[~df_words['words'].isin(wo)]
            g1=df_words.groupby('words').size()
            df_res=pd.DataFrame(g1)
            df_res=df_res.sort([0],ascending=0)
            print '\n  Untuk',w,'word, Waktunya :',time()-time1
            df_res.to_excel(writer,'Sheet1')
            
        else:
            sht='Sheet'+str(w)
            AllWords2=Word_list(AllWords,w)
            time1=time()
            df_words=pd.DataFrame(pd.Series(AllWords2),columns=[('words')])
            df_words=df_words.loc[~df_words['words'].isin(wo)]
            g1=df_words.groupby('words').size()
            df_res=pd.DataFrame(g1)
            df_res=df_res.sort([0],ascending=0)
            print '\n  Untuk',w,'word, Waktunya :',time()-time1
            df_res.to_excel(writer,sht)
        
        w=w+1
    
    print '\n Exporting..'
    writer.save()
    print 'Done'
    print '--------------Thank You-------------'
    sleep(3.0)
    #except:
    #   print('Ada Error.. Tanya yang buat.. Masa Error? (-____-) Malu-maluin')
    #   sleep(5.0)

if __name__ == "__main__":
    main()




