import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import *
import tkinter
import zam
import datetime


class app:
    
    def __init__(self, master):
        self.time = datetime.datetime.now() # datetime object to be used in filename
        frame = Frame(master)
        frame.pack(side='top', fill='x', expand=False)
        

        self.button = Button(frame, text='Dosya seç', command= self.button_command)
        self.button.pack(side='left')

        self.filename_label = Label(frame, text='')    
        self.filename_label.pack(side='left', padx=20, fill='x', expand=True)


        
        frame_bottom = Frame(master)
        frame_bottom.pack(side='top', fill='both', expand=True, pady=50)
                
        
        self.label1 = Label(frame_bottom, text='Zam Oranı')
        self.label1.grid(row=0, column=0, sticky='w')

        self.rise_percent = Entry(frame_bottom)
        self.rise_percent.grid(row=0, column=1, padx=15, sticky='w')
        
        
        self.round_down = IntVar()
        self.cb2 = Checkbutton(frame_bottom, text= 'Asagi yuvarla', variable=self.round_down).grid(row=1,column=0, sticky='w')
        
        self.is_dollar = IntVar()
        self.cb3 = Checkbutton(frame_bottom, text= 'Dolar Listesi', variable=self.is_dollar).grid(row=2, column=0, sticky='w')

        self.todays_date_into_filename = Button(frame_bottom, text='Bugünün Tarihini Ekle', command=self.add_date).grid(row=2, column=1, sticky='e')

        self.output_file_name_label = Label(frame_bottom, text='Yeni Dosya Adı').grid(row=3, column=0, sticky='w', pady=20)
        self.output_file_name = Entry(frame_bottom)
        self.output_file_name.grid(row=3, column=1, sticky='w', padx=15, ipadx=50, pady=20)
        
        self.submit_button = Button(frame_bottom, text='Zam Yap', command=self.submit_button_func).grid(row=4, column=1, sticky='w', padx=15, pady=10)

        self.result_label = Label(frame_bottom, text='')
        self.result_label.grid(row=5, column=1, ipady=25, sticky='w')
        
    
    def add_date(self):
        if self.output_file_name.get() != '':
            date = self.time.strftime('%d') + '.' + self.time.strftime('%m') + '.' + self.time.strftime('%Y')
            output_file_with_date = self.output_file_name.get().split('.docx')[0] + '_' + date
            self.output_file_name.delete(0, 'end')
            self.output_file_name.insert(0, output_file_with_date)
            
        else:
            messagebox.showinfo(title='Hata', message='Dosya Adı Boş Olamaz')
        
    def button_command(self):
        self.file_path = filedialog.askopenfilename()       # get filepath
        filename = self.file_path.split('/')[-1]            # filename
        self.path = self.file_path.split(filename)[0]       # path without filenamem - used for saving new file into same location
        self.filename_label["text"] = filename              # show filename on the top of the screen
        
        self.temp_output_file = filename.split('.docx')[0]  # filename without .docx shows on the entry temporarily
        self.output_file_name.delete(0, 'end')
        self.output_file_name.insert(0, self.temp_output_file)
    
    def submit_button_func(self):
        
        if self.output_file_name.get() == self.temp_output_file:            # returns an error if output filename is same with the old version
            messagebox.showinfo(title='Hata', message='Yeni Dosya Adi Eskisi İle Aynı Olamaz!')
            self.result_label['text'] = 'Dosya Keydedilemedi. Lütfen Tekrar Deneyiniz'
            return

        if self.rise_percent.get() == '':
            messagebox.showinfo(title='Hata', message='Zam Orani Boş Olamaz!')
            return
        
        
        if self.round_down.get() == 0:      # checks if round up or down
            round_up = True
        else:
            round_up = False
        
        if self.is_dollar.get() == 1:   # checks wheter it's dollar list
            para_birimi = '$'
        else:
            para_birimi = '₺'
        
        try:
            output_file_name = self.output_file_name.get() + '.docx'        # output file name to be saved
            
            output_file_name = self.path + output_file_name                 # output file name with full path
            new_doc = zam.change_doc(self.file_path,int(self.rise_percent.get()), round_up, para_birimi)    # send file into function which will return a new file

            new_doc.save(output_file_name)      # save new file
            
            self.result_label['text'] = 'Dosya Kaydedildi'

        except:
             self.result_label['text'] = 'Dosya Keydedilemedi. Lütfen Tekrar Deneyiniz'


if __name__ == "__main__":
    
    root = tk.Tk()
    root.title('Otomatik Zam')
    root.geometry('400x450')
    obj = app(root)
    root.mainloop()
