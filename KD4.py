# -*- coding: utf-8 -*-
"""
Created on Tue Oct 27 15:36:10 2015

@author: Aska
"""
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
        print colored('ERROR_7.. Too much Word','red')
        A3=[0,1]
        return A3

def main():
    import pandas as pd
    from time import sleep
    from nltk.tokenize import word_tokenize
    from time import time
    import os
    import sys
    import unicodedata
    from progressbar import ProgressBar
    from nltk.probability import FreqDist
    from termcolor import colored
                
    reload(sys)  
    sys.setdefaultencoding('Cp1252')
    Mas=1
    while Mas==1:
        Mas=0
        print '\nInput Your File,, '
        fold='Filee'
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
                    print colored('\nERROR_2.. Not in the choice.','red')
                    print colored('=','red')*15
                    print colored('Try Again','red')
                    print colored('=','red')*15
                    File=0
                
                Finy=FNm[File-1]
                path='Filee\\'+Finy
                try:
                    if path[len(path)-4:]=='xlsx' or path[len(path)-3:]=='xls':
                        r=pd.read_excel(path)
                    else:
                        r=pd.read_csv(path)
                except:
                    Mas=1
                    print colored('ERROR_3.. This is not excel file or csv file','red')
                    sleep(1.0)
                    File=0
                    print colored('=','red')*15
                    print colored('Try Again','red')
                    print colored('=','red')*15
                
                
            except :
                print colored('\nERROR_1.. Integer input','red')
                File=0
                print colored('=','red')*15
                print colored('Try Again','red')
                print colored('=','red')*15
    
    print '*'*60    
    print '\nList of Columns:'
    collist=r.columns.tolist()
    for i in collist : 
        print i
    
    cc1='a'
    while cc1 not in collist:
        print('\nWhich column(s) that you want to analyze? ')
        cc=raw_input('(Seperate by comma if you want to analyze two or more columns) : ')
        cc2=cc.split(',')
        rs=pd.DataFrame()
        for cc in cc2:
            rs2=pd.DataFrame()
            if cc not in collist :
                print colored('\nERROR_4.. There is no ','red'),cc,colored('in Column list','red')
                print colored('=','red')*15
                print colored('Try Again','red')
                print colored('=','red')*15
                cc1=[0,0]
                break
            else:
                rs2['texts']=r[cc].tolist()
                rs=rs.append(rs2)
                cc1=cc
    r=rs
    r.index=[i for i in range(len(r))]
    
    print '*'*60
    yakin='n'
    while yakin=='n' or yakin == 'N':
        try:
            yakin='y'
            print "Consider a word that contain two or more word like 'Good Morning'"
            numb=input('You want to count words that contain (until) how many word :')
            if numb>=5:
                print 'There is too many word. ', numb
                yakin=raw_input('Do you want continue (Y/N)??')
            elif numb==0:
                print colored('\nERROR_5.. You cant count zero word','red')
                print colored('=','red')*15
                print colored('Try Again','red')
                print colored('=','red')*15
                yakin='n'
        except:
            print colored('\nERROR_1.. Integer input','red')
            yakin='N'
            print colored('=','red')*15
            print colored('Try Again','red')
            print colored('=','red')*15
        
	print '*'*60
    #try:
	
	print 'Processing..'
	r22=r.dropna()
	r22=r22.texts.tolist()

	print '\n Tokenizing..'

	tw=[]
	pbar = ProgressBar()   
	for w in pbar(r22):
		sleep(0.000000000000000000001)
		try :
			tw.append(word_tokenize(str(w)))
		except:
			u=unicodedata.normalize('NFKD', w).encode('ascii','ignore')            
			try:
			    tw.append(word_tokenize(str(u)))
			except:
				print 'ERROR 6. There is one text which cannot be tokenized'
				print '-'*15
				print w
				print '-'*15
					
	#Cleansing
	print "\n Cleansiing.."
	bar = ProgressBar()
	tw2=[]
	for i in bar(range(len(tw))):
		sleep(0.000000000000000000001)
		ww=[]
		m=0
		for w in tw[i]:
			if w=='<':
				m=1
			elif w=='>':
				m=0
			
			if m==0 and w not in ['/','-','*','<','>','\\','|','+','_','~','`','+','=','%','#','^','nan','NAN']:                         
				ww.append(w)
		tw2.append(ww)

	AllWords=tw2
	p_wo='Filter.xlsx' #File with un-use word 
	wo=pd.read_excel(p_wo)
	wo=wo['filter'].tolist()
		
	writer = pd.ExcelWriter('File_Result\\'+Finy[:len(Finy)-5]+'_result.xlsx')

	w=1
	print '\n Counting..'
	while w<=numb:
		if w==1:
			AllWords2=Word_list(AllWords,w)
			if AllWords2!=[0,1]:
				time1=time()
				df_words=pd.DataFrame(pd.Series(AllWords2),columns=[('words')])
				df_words=df_words.loc[~df_words['words'].isin(wo)]
				g1=FreqDist(df_words.words.tolist())
				df_res=pd.DataFrame(g1,index=[0]).transpose()
				df_res=df_res.sort([0],ascending=0)
				print '\n  For',w,'word, Counting Time :',time()-time1 
			else:
				df_res=pd.DataFrame(AllWords2)
				print colored('\n  For','red'),colored(w,'red'),colored('word, Counting Time : ERROR_7.. Too much Word','red')
			
			df_res.index=[unicode(x) for x in df_res.index.tolist()]
			df_res.to_excel(writer,'Word_1')
			
		else:
			sht='Word_'+str(w)
			AllWords2=Word_list(AllWords,w)
			if AllWords2!=[0,1]:
				time1=time()
				df_words=pd.DataFrame(pd.Series(AllWords2),columns=[('words')])
				df_words=df_words.loc[~df_words['words'].isin(wo)]
				g1=FreqDist(df_words.words.tolist())
				df_res=pd.DataFrame(g1,index=[0]).transpose()
				df_res=df_res.sort([0],ascending=0)
				print '\n  For',w,'word, Counting Time :',time()-time1                
			else:
				df_res=pd.DataFrame(AllWords2)
				print colored('\n  For','red'),colored(w,'red'),colored('word, Counting Time : ERROR_7.. Too much Word','red')
			df_res.index=[unicode(x) for x in df_res.index.tolist()]
			df_res.to_excel(writer,sht)
		w=w+1

	print '\n Exporting..'
	writer.save()
	print '*'*60
	print 'Done'
	print '--------------Thank You-------------'
	sleep(3.0)
    #except:
    #   print('-PROGRAMME ERROR-')
    #   sleep(5.0)

if __name__ == "__main__":
    from termcolor import colored
    print '-'*50
    print 'Welcome to Keyword Density 4.0'
    print '-'*50    
    print 'This Programme was built to count any word in the text. For our purpose, the programme only works for excel or csv file.'
    print "If your text in ms.word file you can copy your text in excel cell A2 and type A1 as your column name such as 'Text'. "
    print 'Input your excel file on folder "Filee" from this programme. The Result would be end up on folder "File_Result"'
    print '\nHave you input your file in that folder? (press enter button if you have)'
    a=raw_input()
    print '*'*60
    main()