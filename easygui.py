#imports
import numpy
import pandas as pd
import math
import statistics
from itertools import combinations
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import *
from tkinter import messagebox
import math
import collections
from fpdf import FPDF
import xlsxwriter


class wd:
    def __init__(self,parent):
        self.parent = parent
        self.filename = None
        self.data = None
        self.text = tk.Text(self.parent)
        self.text.pack()
        
        self.same_hours = False;
        self.factors=[]
        self.pdf = False;
        self.excel = False;
        self.groupnum = 0
        self.groups ={}
        self.weekday_open = 0
        self.weekday_close = 0
        self.sat_open = 0
        self.sat_close = 0
        self.sun_open = 0
        self.sun_close = 0
        tk.messagebox.showinfo("Note:","If the store is closed for an entire day, please enter typical store hours for that day and then discard that day when the schedule is created.")        
        self.button = tk.Button(self.parent,text="Store Hours",command = self.open_hour)
        self.button.pack()        
        self.button = tk.Button(self.parent,text="Excel File Input",command=self.load)
        self.button.pack()
        
        self.button =tk.Button(self.parent,text = "Number of Groups",command =self.group)
        self.button.pack()
        
        self.button = tk.Button(self.parent,text="View the New Groups Created by the Algorithm",command = self.analyze)
        self.button.pack()
        
        self.button = tk.Button(self.parent,text="File Type Output",command = self.output)
        self.button.pack()
        
    def open_hour(self):
            
        def hour_error(day, message):
            while (True):
                try:    
                    if (day[-2:] == "am" or day[-2:] == "pm") and (int(day[:-2])>=1 and int(day[:-2])<=12):
                        break;
                    else:
                        day = simpledialog.askstring(message, "What is the "+message+" hour?   (ex: 5pm)", parent=application_window)
                except:
                    day = simpledialog.askstring(message, "What is the "+message+" hour?   (ex: 5pm)", parent=application_window)
            return day
        
        def time_convert(time):
            # this function converts times in form "5pm" into 17, or "3am" into 3, or "12am" into 0
            hour = int(time[:-2])
            Ampm = time[-2:]
            new_time = 0
            
            if Ampm == "am":
                if hour < 12:
                    new_time = hour
                else: #if the time is 12am
                    new_time = 0
            else:
                if hour < 12:
                    new_time = hour + 12
                else: # if the time is 12pm
                    new_time = hour
            return new_time

        same_hours = messagebox.askyesno("Store Hours","Are the store hours the same every day?")
        application_window = tk.Tk()
        self.same_hours = same_hours
        if not(same_hours):    
            weekday_open = simpledialog.askstring("Weekday Opening", "What is the Weekday OPENING hour?   (ex: 5pm)", parent=application_window)
            weekday_open = hour_error(weekday_open, "Weekday OPENING")
            weekday_close = simpledialog.askstring("Weekday Closing", "What is the Weekday CLOSING hour?   (ex: 5pm)", parent=application_window)
            weekday_close = hour_error(weekday_close, "Weekday CLOSING")
            
            sat_open = simpledialog.askstring("Saturday Opening", "What is the Saturday OPENING hour?   (ex: 5pm)", parent=application_window)
            sat_open = hour_error(sat_open, "Saturday OPENING")
            sat_close = simpledialog.askstring("Saturday Closing", "What is the Saturday CLOSING hour?   (ex: 5pm)", parent=application_window)
            sat_close = hour_error(sat_close, "Saturday CLOSING")
            
            sun_open = simpledialog.askstring("Sunday Opening", "What is the Sunday OPENING hour?   (ex: 5pm)", parent=application_window)
            sun_open = hour_error(sun_open, "Sunday OPENING")
            sun_close = simpledialog.askstring("Sunday Closing", "What is the Sunday CLOSING hour?   (ex: 5pm)", parent=application_window)
            sun_close = hour_error(sun_close, "Sunday CLOSING")        
            
            self.weekday_open = time_convert(weekday_open)
            self.weekday_close = time_convert(weekday_close)
            self.sat_open = time_convert(sat_open)
            self.sat_close = time_convert(sat_close)
            self.sun_open = time_convert(sun_open)
            self.sun_close = time_convert(sun_close)
            
        else:
            open_hour = simpledialog.askstring("Opening", "What is the OPENING hour?   (ex: 5pm)", parent=application_window)
            open_hour = hour_error(open_hour, "OPENING")
            close_hour = simpledialog.askstring("Closing", "What is the CLOSING hour?   (ex: 5pm)", parent=application_window)
            close_hour = hour_error(close_hour, "CLOSING")
            
            
            self.weekday_open = time_convert(open_hour)
            self.weekday_close = time_convert(close_hour)
            self.sat_open = time_convert(open_hour)
            self.sat_close = time_convert(close_hour)
            self.sun_open = time_convert(open_hour)
            self.sun_close = time_convert(close_hour)           
        
        #calcualates the possible group sizes that the user can choose based on how many hours the store is open
        num_hours = abs(self.weekday_close - self.weekday_open)
        factors = []
        check = 0 
        for a in list(range(num_hours)):
            if num_hours in [3,5,7,11,13,17,19,23]:
                if num_hours <= 10:
                    factors.append(num_hours)
                num_hours-=1
                check = 1
            if (num_hours % (a+1) == 0 and a <= 10):
                if check == 1:
                    factors.insert(-1,a+1)
                elif  check == 0:
                    factors.append(a+1)
        if factors == []:
            factors.append(1)
        factors.sort()
        self.factors=factors
    def group(self):
        answer = simpledialog.askinteger("Input", "Please choose from :"+",".join(list(map(str,self.factors))),parent=self.parent)
        while (answer not in self.factors):
            answer = simpledialog.askinteger("Input", "Please choose from :"+",".join(list(map(str,self.factors))),parent=self.parent)
        self.groupnum = answer
    def load(self):
        file = filedialog.askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx'))])
        if file:
            self.data = pd.read_excel(file)
            self.filename = file
    def analyze(self):
        if self.data is None:
            self.load()
        else:
            df = pd.DataFrame(self.data, columns= ['Last Name'])       
            #create list of letters
            last_name = df.to_numpy()
            d = {}

            for i in list(range(26)):
                d[chr(65+i)]=0
            for name in last_name:
                if type(name[0])==str:
                    if name[0][0].upper() in d:
                        d[name[0][0].upper()] += 1
                    else:
                        d[name[0][0].upper()] = 1
            # up to this point we have a dictionary d that stores the count of each first letter of each last name (also taking lower case last names into account)
            # {'S': 373, 'J': 394, 'U': 385, 'T': 353, 'R': 375, 'A': 355, 'O': 403, 'E': 376, 'C': 378, 'F': 405, ...........
            
            #create list of counter for each letter, in alphabetical order
            count_order = []
            for i in list(map(chr,range(ord('A'),ord('Z')+1))):
                count_order.append(d[i])
            
            #count_order: [355, 389, 378, 385, 376, ..........
            
            #calculate total number of people in excel file
            total_num = sum(count_order)

            #calculate group sizes
            
            group_num  = self.groupnum
            decimal_group_size = total_num / group_num
            
            def f(k):
                l = list(range(1,26))
                return list(combinations(l,k))
            def g(l,t):
                temp = []
                start = 0
                for i in t:
                    end = i
                    temp.append(sum(l[start:end]))
                    start = i
                temp.append(sum(l[start:]))
                return statistics.variance(temp)
                 
            tuple_l = f(group_num-1)
            big_var = list(map(lambda item:g(count_order,item),tuple_l))
            min_var = min(big_var)
            min_index = big_var.index(min_var)
            best_tuple = tuple_l[min_index]
            final  = []
            start = 0
            for i in best_tuple:
                end  = i
                final.append(count_order[start:end])
                start = i
            final.append(count_order[start:])
            
            alphabet_result = {}
            curlen = -1
            for i in list(range(1,len(final)+1)):
                alphabet_result[i] = [chr(65+curlen+1)]
                curlen += len(final[i-1])
                alphabet_result[i].append(chr(65+curlen))
            for i in alphabet_result:
                if (len(set(alphabet_result[i]))==1): 
                    self.text.insert('end',"Group "+str(i)+": "+alphabet_result[i][0]+"\n")
                else:
                    self.text.insert('end',"Group "+str(i)+": "+"-".join(alphabet_result[i])+"\n")
            self.groups = alphabet_result
    def output(self):
        '''
        subroot = tk.TK()
        subroot.title("determine the output file type")
        subroot.geometry("500x500")
        '''
        def reverse_time_convert(time):
            # this function converts times in form 17 into "5pm", or 3 into "3am", or 0 into "12am"
            new_time = ""
            if time <= 11:
                if time == 0:
                    new_time = "12am"
                else:
                    new_time = str(time)+"am"
            elif time <= 23:
                if time == 12:
                    new_time = "12pm"
                else:
                    new_time = str(time - 12) + "pm"
            return new_time        
        
        def submit():
            if (pdf.get()==1):
                self.pdf = True
            if (excel.get()==1):
                self.excel = True
                
        pdf = IntVar()  
        excel= IntVar()
        
        pdfbutton = Checkbutton(self.parent, text = "PDF (.pdf)", variable = pdf, onvalue = 1, offvalue = 0, height = 2, width = 10)  
        excelbutton= Checkbutton(self.parent, text = "Excel (.xlsx)", variable = excel, onvalue = 1, offvalue = 0, height = 2, width = 10)  
        pdfbutton.pack()  
        excelbutton.pack()  
        
        submit_button = Button(self.parent,text="Enter",command = submit).pack()
        self.parent.mainloop()
        
        #print("after submission we output: ",self.pdf,self.excel)
        
        if (self.pdf ==True): print("pdf")
        if (self.pdf == True and top.excel == True): print(" and ")
        if (self.excel ==True): print("excel")          
        
        num_hours = abs(self.weekday_close - self.weekday_open)
        
        schedule = {}
        #schedule = {"Monday":[["14","Group1"],["15","Group1"],["16","Group2"],["17","Group2"],["18","Group3"],["19","Group3"]],"Tuesday":[]}
        cur_hour = self.weekday_open
        schedule["Monday"]=[]
        quotient = math.floor(num_hours / self.groupnum)
        check = 0
        if (num_hours in [3,5,7,11,13,17,19,23]):
            schedule["Monday"].insert(0,"Free Hour")
            cur_hour+=1
            check = 1
        while cur_hour < self.weekday_close:
            for i in list(range(quotient)):
                schedule["Monday"].append("Group "+str(math.floor((cur_hour-self.weekday_open-check)/quotient)+1))
                cur_hour+=1
        
        if (check == 0):
            p=collections.deque(schedule["Monday"])
            p.rotate(quotient)
            schedule["Tuesday"]=list(p)
            p.rotate(quotient)
            schedule["Wednesday"]=list(p)
            p.rotate(quotient)
            schedule["Thursday"]=list(p)
            p.rotate(quotient)
            schedule["Friday"]=list(p)
        if (check == 1):
            d=collections.deque(schedule["Monday"][1:])
            d.rotate(quotient)
            schedule["Tuesday"]=["Free Hour"]+list(d)
            d.rotate(quotient)
            schedule["Wednesday"]=["Free Hour"]+list(d)
            d.rotate(quotient)
            schedule["Thursday"]=["Free Hour"]+list(d)
            d.rotate(quotient)
            schedule["Friday"]=["Free Hour"]+list(d)
        schedule["Saturday"]=[]
        schedule["Sunday"]=[]
        sat=[]
        
        sat_hours = abs(self.sat_close - self.sat_open)
        if (sat_hours<self.groupnum):
            sat=["Free Hour"]*sat_hours
        elif ((sat_hours/2 <= self.groupnum) and (sat_hours >= self.groupnum)):
            for i in list(range(1,self.groupnum+1)):
                sat.append("Group "+str(i))
            p=sat_hours-self.groupnum
            w = collections.deque(sat)
            w.rotate(math.floor(sat_hours/self.groupnum)*2)
            sat = list(w)            
            if sat_hours == 2 * self.groupnum:
                sat += sat
            else:
                while p > 0:
                    if (p%2==0):
                        sat.append("Free Hour")
                    else:
                        sat.insert(0,"Free Hour")
                    p-=1

        else:
            loop = math.floor(sat_hours /(self.groupnum*2))
            n = loop
            while loop >0:
                for i in list(range(1, self.groupnum+1)):
                    sat.append("Group "+str(i))
                    sat.append("Group "+str(i))
                loop-=1
            p=sat_hours - self.groupnum*2*n
            r = collections.deque(sat)
            r.rotate(math.floor(sat_hours/self.groupnum)*2)
            sat = list(r)            
            while p > 0:
                if (p%2==0):
                    sat.append("Free Hour")
                else:
                    sat.insert(0,"Free Hour")
                p-=1
        
        schedule["Saturday"]=sat
        #print("Saturday opens at "+str(self.sat_open)+" and closes at "+str(self.sat_close)+" and we have "+str(self.groupnum)+" groups")
        sun=[]
        sun_hours = abs(self.sun_close - self.sun_open)
        if (sun_hours<self.groupnum):
            sun=["Free Hour"]*sun_hours
        elif ((sun_hours/2 <= self.groupnum) and (sun_hours >= self.groupnum)):
            for i in list(range(1,self.groupnum+1)):
                sun.append("Group "+str(i))
            sun+=["Free Hour"]*(sun_hours-self.groupnum)
            c=sun_hours-self.groupnum
            p=sun_hours-self.groupnum
            while p > 0:
                if (p%2==0):
                    sun.append("Free Hour")
                else:
                    sun.insert(0,"Free Hour")
                p-=1            

        else:
            loop = math.floor(sun_hours /(self.groupnum*2))
            n=loop
            while loop >0:
                for i in list(range(1, self.groupnum+1)):
                    sun.append("Group "+str(i))
                    sun.append("Group "+str(i))
                loop-=1
            c=sun_hours - self.groupnum*2*n
            p=sun_hours - self.groupnum*2*n
            while p > 0:
                if (p%2==0):
                    sun.append("Free Hour")
                else:
                    sun.insert(0,"Free Hour")
                p-=1

        if (self.sat_open==self.sun_open and self.sat_close==self.sun_close):
            sun=list(filter(lambda item:item!="Free Hour",schedule["Saturday"]))
            d = collections.deque(sun)
            open_hours = abs(self.sun_open-self.sun_close)
            q = math.floor(open_hours/self.groupnum)
            if (q%2==1):q-=1
            n = 0
            if (self.groupnum >= 4):
                n = q
            else:
                n=math.floor(open_hours/2)
                if (n%2==1):n-=1
            d.rotate(n)
            sun = list(d)
            c = sat.count("Free Hour")
            while c>0:
                if (c%2==0):
                    sun.append("Free Hour")
                else:
                    sun.insert(0,"Free hour")
                c-=1
        schedule["Sunday"]=sun
        early = min(self.weekday_open,self.sat_open,self.sun_open)
        close = max(self.weekday_close,self.sat_close,self.sun_close)
        if (self.sat_open==self.sun_open and self.sat_close==self.sun_close):
            sun=list(filter(lambda item:item!="Free Hour",schedule["Saturday"]))
            d = collections.deque(sun)
            open_hours = abs(self.sun_open-self.sun_close)
            q = math.floor(open_hours/self.groupnum)
            if (q%2==1):q-=1
            d.rotate(q)
            sun = list(d)
            c = sat.count("Free Hour")
            while c>0:
                if (c%2==0):
                    sun.append("Free Hour")
                else:
                    sun.insert(0,"Free hour")
                c-=1
        schedule["Sunday"]=sun
        early = min(self.weekday_open,self.sat_open,self.sun_open)
        close = max(self.weekday_close,self.sat_close,self.sun_close)        
        if (self.pdf == True):
            data = [["Hours","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]]           
            for i in list(range(early,close)):
                data.append([reverse_time_convert(i)+" - " +reverse_time_convert(i+1)])
            count = 0
            cur = early
            while (cur < close):
                if (cur<self.weekday_open or cur >= self.weekday_close):
                    data[cur-early+1].append("Closed")
                    data[cur-early+1].append("Closed")
                    data[cur-early+1].append("Closed")
                    data[cur-early+1].append("Closed")
                    data[cur-early+1].append("Closed")
                else:
                    data[cur-early+1].append(schedule["Monday"][count])
                    data[cur-early+1].append(schedule["Tuesday"][count])
                    data[cur-early+1].append(schedule["Wednesday"][count])
                    data[cur-early+1].append(schedule["Thursday"][count])
                    data[cur-early+1].append(schedule["Friday"][count])
                    count+=1
                cur+=1
            cur = early
            count = 0
            while (cur<close):
                if (cur<self.sat_open or cur >= self.sat_close):
                    data[cur-early+1].append("Closed")
                else:
                    data[cur-early+1].append(schedule["Saturday"][count])
                    count+=1
                cur+=1
            cur = early
            count = 0
            while (cur<close):
                if (cur<self.sun_open or cur >=self.sun_close):
                    data[cur-early+1].append("Closed")
                else:
                    data[cur-early+1].append(schedule["Sunday"][count])
                    count+=1
                cur+=1
              
            pdf = FPDF()    
            spacing = 3
            pdf.set_font("Arial", size=15)
            pdf.add_page()
            pdf.set_left_margin(0)
            pdf.set_right_margin(0)
            pdf.set_font('Arial','B',10)
            s = "The Shopping Schedule              "
            #print(self.groups)
            pdf.set_font('Arial','B',7)
            for i in self.groups:
                s+="Group "+str(i)+": "+self.groups[i][0]
                if len(self.groups[i])>1:
                    s+="-"+self.groups[i][-1]+" "
            pdf.cell(50,50,s,0,1)    
            pdf.set_font("Arial", size=10)
            col_width = pdf.w / 8
            row_height = pdf.font_size
            for row in data:
                for item in row:
                    pdf.set_left_margin(0)
                    pdf.set_right_margin(0)
                    pdf.cell(col_width, row_height*spacing,
                             txt=item, border=1)
                pdf.ln(row_height*spacing)
                
            pdf.output('schedule.pdf')    
      
        if (self.excel == True):
            ex_workbook = xlsxwriter.Workbook('Shopping Schedule.xlsx')
            ex_worksheet = ex_workbook.add_worksheet()
            
            days = ["Hours","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
            
            
            total_hours = close - early
            row = 0
            column = 0
            current_hour = early
            count = 0
            for c in days:
                ex_worksheet.write(0, column, c)       
                row = 1
                for r in list(range(early,close)):
                    if c == "Hours":
                        ex_worksheet.write(row, 0, reverse_time_convert(r)+" - "+reverse_time_convert(r+1))
                    else:
                        if (c == "Saturday"):
                            if (self.sat_open > current_hour) or (self.sat_close <= current_hour):
                                ex_worksheet.write(row, column, "Closed")
                            else:
                                ex_worksheet.write(row, column, schedule[c][count])
                                count += 1
                        elif (c == "Sunday"):
                            if (self.sun_open > current_hour) or (self.sun_close <= current_hour):
                                ex_worksheet.write(row, column, "Closed")
                            else:
                                ex_worksheet.write(row, column, schedule[c][count])
                                count += 1
                        else:
                            if (self.weekday_open > current_hour) or (self.weekday_close <= current_hour):
                                ex_worksheet.write(row, column, "Closed")
                            else:
                                ex_worksheet.write(row, column, schedule[c][count])
                                count += 1                            
                    row += 1
                    current_hour += 1
                count = 0
                column += 1
                current_hour = early
            
            row = 0
            for d in self.groups:
                ex_worksheet.write(row, 10, "Group "+str(d)+":")
                if (self.groups[d][0] == self.groups[1]):
                    ex_worksheet.write(row, 11, str(self.groups[d][0]))
                else:
                    ex_worksheet.write(row, 11, str(self.groups[d][0])+"-"+str(self.groups[d][1]))
                row += 1
            
            ex_workbook.close()


if __name__ == '__main__':
    file = open("Instructions.txt","w")
    L=["Instructions:\n\n","1. First Pop-up Window\"\n","  -> Read the note\n","  -> Press 'OK'\n","2. Click on \"Store Hours\"\n","  -> Input your store's hours\n","  -> If you input an hour format that is not acceptable, you will be prompted to re-enter that time\n","3. Click on \"Excel File Input\"\n","  -> Find the excel file with the list of names\n","  -> Click on the excel file, and press \"Open\"\n","4. Click on \"Number of Groups\"\n","  -> Input how many different groups you want\n","  -> If you don't choose from the given list, you will be prompter to re-enter a given group number\n","5. Click on \"View the New Groups Created \" *OPTIONAL*\n","  -> This will show you the groups that were created\n","6. Click on \"File Type Output\"\n","  -> Choose which file(s) you would like the new schedule on\n","  -> Press Enter\n","7. Close any window(s) that are open by clicking the \"x\" in the top right corner\n","8. A new Excel and/or PDF file(s) of your new schedule will now be saved\n","  -> Open the file that your original Excel file was saved in to view the new Excel and/or PDF file(s) named \"Shopping Schedule\"\n"]
    file.writelines(L)
    file.close()
    root = tk.Tk()
    top = wd(root)
    root.mainloop()




'''
   TESTS:
   
   same hours everyday:
   				num hours:
          1, 7, 12
   
   saturday and sunday the same (differet from weekday):
          sat/sun hours:
          4, 9
          weekday hours:
          1, 2, 14
   
   saturday and sunday differenet (differerent from weekday):
          sat hours:
          1, 6
          sun hours:
          5, 10
          weekday hours:
          3, 8, 13
   
   
   # DO LATER
    - error checking
    - comment out any lines of code we don't need (print statements for example)
    - write out exactly what our algorithm does and why it's so useful and how it handles errors (cannot have excel or pdf file open while running the code, don't close the extra windows until you've filled out the other ones, you can go back and re-enter any value instead of having to rerun the program again)


Instructions Before Running Code:

What This Algorithm Does:

  
1. Excel File:
  -> Save an excel file on the computer with a list of names
  -> The code needs to read the names to create a schedule with evenly distribution group sizes based on last names

2. Excel File Format:
  -> Have ONE column header labeled "Last Names"
  -> In that column underneath "Last Names", have a LAST NAME ONLY (do not have first names in that column) in each cell

3. In Command Prompt:
  -> 




Instructions:

1. First Pop-up Window
  -> Read the note
  -> Press 'OK'

2. Click on "Store Hours"
  -> Input your store's hours
  -> If you input an hour format that is not acceptable, you will be prompted to re-enter that time
  
3. Click on "Excel File Input"
  -> Find the excel file with the list of names
  -> Click on the excel file, and press "Open"

4. Click on "Number of Groups"
  -> Input how many different groups you want
  -> If you don't choose from the given list, you will be prompter to re-enter a given group number

5. Click on "View the New Groups Created" *OPTIONAL*
  -> This will show you the groups that were created 

6. Click on "File Type Output"
  -> Choose which file(s) you would like the new schedule on
  -> Press Enter

7. Close any window(s) that are open by clicking the "x" in the top right corner

8. A new Excel and/or PDF file(s) of your new schedule will now be saved
  -> Open the file that your original Excel file was saved in to view the new Excel and/or PDF file(s) named "Shopping Schedule"

'''