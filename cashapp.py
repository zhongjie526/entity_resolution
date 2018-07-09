#!/usr/bin/env python

import pandas as pd
import math
import numpy as np
import re
import distance
import nltk
import time 

def timeit(method):
    def timed(*args, **kw):
        ts = time.time()
        result = method(*args, **kw)
        te = time.time()

        if 'log_time' in kw:
            name = kw.get('log_name', method.__name__.upper())
            kw['log_time'][name] = int((te - ts) * 1000)
        else:
            print '%r  %2.2f ms' % \
                  (method.__name__, (te - ts) * 1000)
        return result

    return timed


def tokenize(sent):
    """
    When passed in a sentence, tokenizes and normalizes the string,
    returning a list of lemmata.
    """
    lemmatizer = nltk.WordNetLemmatizer() 
    for token in nltk.wordpunct_tokenize(sent):
        token = token.lower()
        yield lemmatizer.lemmatize(token)

def normalized_jaccard(*args):
    try:
        return distance.jaccard(*[tokenize(arg) for arg in args])
    except UnicodeDecodeError:
        return 1.0
    
def wordsim(*args):
    return 1.0-normalized_jaccard(*args)

def write_to_excel(df,name):
    writer = pd.ExcelWriter(name,  datetime_format='MM/dd/yyyy')
    df.to_excel(writer,index=False)
    writer.close()

def days_between(d1, d2):
    return abs(d2 - d1)/np.timedelta64(1, 'D')


invoice = pd.read_excel("CCN 4189 FBL5N_Jul-Dec 2017.XLSX",dtype={'Reference':'str','Document Number':'str','Clearing Document':'str'})

invoice_1 = invoice[(invoice['Clearing date'] < '20171231') & (invoice['Posting Date'] > '20170701') & (invoice['Amount in doc. curr.'] > 0)]


bank = pd.read_excel("CCN 4189 Bank Statements_JPM_Oct-Dec 2017.xls",sheet_name="Details",\
                     dtype={'Remarks 1':'str','Remarks 2':'str','Remarks 3':'str','Remarks 4':'str',\
                            'Remarks 5':'str','Remarks 6':'str'})

bank_1 = bank[(bank['Description'] != 'Debit Summary') & (bank['Description'] != 'Check Summary')]


matches=[]

@timeit
def entity_resolution():
	n_customer=0
	n_no_customer=0
	n_match=0
	lst_customer=[]
	lst_name_match=[]
	global matches

	for index,row in bank_1[:10].iterrows():
	    print "========================================================================="
	    remarks = row['Remarks 1']+"|"+row['Remarks 2']+"|"+row['Remarks 3']+"|"+row['Remarks 4']+"|"+\
		      row['Remarks 5']+"|"+row['Remarks 6']
	    remarks = re.sub( '[\s\.]+', ' ',remarks).strip()
	    remarks = remarks.replace('PETRON1/AS','PETRONAS')
	    remarks = remarks.replace('PTE LT1','PTE LTD 1')
	    remarks = remarks.replace('PTE LTDPAY','PTE LTD PAY')
	    remarks = remarks.replace('PRIVATE LTD','PTE LTD')
	    remarks = remarks.replace('LTD','LTD|')
	    remarks = remarks.replace('PLC','PLC|')
	    match_result = re.match(r".*B\/O CUSTOMER[\W\d\d]*([a-zA-Z\s\(\)\&]+).*",remarks)
	    if match_result==None:
		#print "no customer name: "+ remarks 
		n_no_customer+=1

	    else:
		cust_name = match_result.group(1).replace("PTE","").replace("LTD","").replace("PLC","").strip()
		#print"customer name found: "+cust_name
		n_customer+=1
		lst_customer.append(cust_name)
		amt = row['Credit Amount']
		date = row['Transaction Date']
		#print "{1:'%Y-%m-%d'}:Credit amount {0}".format(amt,date)
		
		date_matched = invoice_1[(invoice_1['Net due date']>=date) & (invoice_1['Posting Date']<=date)]
		n_date_matched= len(date_matched)
		
		if(n_date_matched>0):
		    match = date_matched.to_dict('records')
		    for each in match:
			acc_inv = each['Account Name'].replace("PTE","").replace("LTD","").replace("PLC","").strip()
			amt_inv = each['Amount in doc. curr.']
			date_inv = each['Net due date']
			name_score = wordsim(acc_inv,cust_name)
			amt_score = 1.0-abs(amt_inv/amt-1.0)
			date_score =  math.exp(-days_between(date,date_inv))
			final_score = name_score*0.3+amt_score*0.4+date_score*0.3
			

			if(final_score>0.7):
			    print "name score: {0}  amount score: {1}  date score: {2}".format(name_score,amt_score,date_score)
			    print "invoice matched: amount {0} from account {2} due on {1:'%Y-%m-%d'}".format(amt_inv,date_inv,acc_inv)
			    n_match+=1
			    matches.append({"Bank acc":cust_name,\
					    "Bank date":date,"Bank amt":amt,\
					    "Invoice acc":acc_inv,"Invoice amt":amt_inv,\
					    "Invoice due":date_inv, "Name score":name_score,\
					    "Amount score":amt_score,"Date score":date_score,\
					    "Final score":final_score
					    })
		    
		else:
		    print "no date range"

		    #print date_matched[['Account Name','Net due date','Amount in doc. curr.']]
		

	print "customer names found: {0}".format(n_customer)
	print "customer names not found: {0}".format(n_no_customer)
	print "final match count: {0}".format(n_match)


entity_resolution()

pd_matched = pd.DataFrame(matches)[['Bank acc','Bank date','Bank amt','Invoice acc','Invoice due','Invoice amt','Name score',\
                       'Date score','Amount score','Final score']]

write_to_excel(pd_matched,"matched_transactions.xls")
