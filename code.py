
from sklearn.externals import joblib
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd
import xlwt 
from xlwt import Workbook 

df = pd.read_csv('part1.csv')
event = df.iloc[:,3:]


events = joblib.load('events.pickle')
domains = joblib.load('domains.pickle')

e = []
d = []
sent = input('Enter your sentence: ')

for i in e:
	if i in sent:
		e.append(i)

for j in d:
    if j in sent:
        d.append(j)

print(d)


event1 = input("Enter Event 1: ")
event2 = input("Enter Event 2: ")



event1 = events[event1.lower()]
event2 = events[event2.lower()]

similar_met = cosine_similarity(event)

output = np.where(event == [event1,event2])[0][0]

output = list(enumerate(similar_met[output]))
output = sorted(output,key= lambda x: x[1],reverse=True)
top = output[:5]

lst = []

for i in top:
  c = (df['Name'].iloc[i[0]])
  lst.append(c)

print(sent, '|', lst[0], ',' ,lst[1], ',' ,lst[2], ',', lst[3], ',', lst[4])

df1 = pd.DataFrame([s], columns=['Events Name'])
df1

df1['Employee Name'] = [lst]
df1

def converter(input_seq, seperator):
   final_str = seperator.join(input_seq)
   return final_str

seperator = ', '
print("Recommended Employees: ", converter(lst, seperator))


  
sheets = Workbook() 
  
sheet1 = sheets.add_sheet('Sheet 1')

sheet1.write(0, 0, 'Event Name') 
sheet1.write(0, 1, 'Employee Name') 
sheet1.write(1, 0, sent)
sheet1.write(1, 1, converttostr(lst, seperator))

sheets.save('result.xls')
