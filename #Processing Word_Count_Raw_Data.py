#Processing Word_Count_Raw_Data.txt
import xlsxwriter


f=open("Word_Count_Raw_Data.txt", "r")

infile=f.readlines()

ssn=[]
eps=[]
name_lst=[]
count_lst=[]

ssn_num=0
eps_num=0
name=""
count=0

for i in range(len(infile)):
	my_str=infile[i].strip().split(":")
	#print (my_str)

	if my_str[0]=='"seasonNum"':
		ssn_num=int(my_str[1].strip().strip(","))
		
	elif my_str[0]=='"episodeNum"':
		eps_num=int(my_str[1].strip().strip(","))

	elif my_str[0]=='"name"':
		name=my_str[1].strip().strip(",")

	elif my_str[0]=='"count"':
		count=int(my_str[1])

		ssn.append(ssn_num)
		eps.append(eps_num)
		name_lst.append(name)
		count_lst.append(count)


		

	#	print(i)


f.close()

	#print(ssn_num, eps_num,name,count)
#print(ssn,eps,name_lst,count_lst)


# Workbook() takes one, non-optional, argument  
# which is the filename that we want to create. 
workbook = xlsxwriter.Workbook('Word_Count_Raw_Data_Processed.xlsx') 
  
# The workbook object is then used to add new  
# worksheet via the add_worksheet() method. 
worksheet = workbook.add_worksheet() 
  
# Use the worksheet object to write 
# data via the write() method. 
worksheet.write('A1', 'Season') 
worksheet.write('B1', 'Episode') 
worksheet.write('C1', 'Name') 
worksheet.write('D1', 'Word Count') 

for i in range(len(ssn)):
	mystr=str(i+2)
	A='A'+ mystr
	B='B'+ mystr
	C='C'+ mystr
	D='D'+ mystr

	worksheet.write(A, ssn[i]) 
	worksheet.write(B, eps[i]) 
	worksheet.write(C, name_lst[i].replace('"','')) 
	worksheet.write(D, count_lst[i]) 



  
# Finally, close the Excel file 
# via the close() method. 
workbook.close() 
