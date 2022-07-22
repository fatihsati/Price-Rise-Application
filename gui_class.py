import tkinter as tk
from tkinter import filedialog
from tkinter import *
import zam

# isimde tarih varsa onu atip yeni tarih koymali regex ile tarihi bul ve at
# kodu exe dosyasi haline getir
# github repo acip pushla



class app:
    
    def __init__(self, master):
        self.file_path = ''
        frame = Frame(master)
        frame.pack(side='top', fill='x', expand=False)
        

        self.label_text = ''
        self.button = Button(frame, text='Dosya seç', command= self.button_command)
        self.button.pack(side='left')

        self.filename_label = Label(frame, text=self.label_text)    # label style eklenecek
        self.filename_label.pack(side='left', padx=20, fill='x', expand=True)


        
        
        frame_right = Frame(master)
        frame_right.pack(side='top', fill='both', expand=True, pady=50)
                
        
        self.label1 = Label(frame_right, text='Zam Oranı')
        self.label1.grid(row=0, column=0, sticky='w')

        self.zam_orani = Entry(frame_right)
        self.zam_orani.grid(row=0, column=1, padx=15, sticky='w')
        
        
        self.asagi_yuvarla = IntVar()
        self.cb2 = Checkbutton(frame_right, text= 'Asagi yuvarla', variable=self.asagi_yuvarla).grid(row=1,column=0, sticky='w')
        
        self.dolar_listesi = IntVar()
        self.cb3 = Checkbutton(frame_right, text= 'Dolar Listesi', variable=self.dolar_listesi).grid(row=2, column=0, sticky='w')

        self.submit_button = Button(frame_right, text='Zam Yap', command=self.submit_button_func).grid(row=3, column=1, sticky='w', padx=15)

        self.result_label = Label(frame_right, text='')
        self.result_label.grid(row=5, column=1, ipady=25, sticky='w')
        
        
    def button_command(self):
        self.file_path = filedialog.askopenfilename()
        filename = self.file_path.split('/')[-1]
        self.filename_label["text"] = filename
    
    
    def submit_button_func(self):
        
        if self.asagi_yuvarla.get() == 0:   
            round_up = True
        else:
            round_up = False
        
        if self.dolar_listesi.get() == 1:
            para_birimi = '$'
        else:
            para_birimi = '₺'
        
        try:
            output_file_name = self.file_path.split('.docx')[0] + '23.07.2022.docx'
            new_doc = zam.change_doc(self.file_path,int(self.zam_orani.get()), round_up, para_birimi)

            new_doc.save(output_file_name)
            
            self.result_label['text'] = 'Dosya Kaydedildi'

        except:
             self.result_label['text'] = 'Dosya Keydedilemedi. Lütfen Tekrar Deneyiniz'


if __name__ == "__main__":
    
    root = tk.Tk()
    root.title('Otomatik Zam')
    root.geometry('400x400')
    obj = app(root)
    root.mainloop()




