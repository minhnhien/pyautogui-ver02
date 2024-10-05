import pyautogui as pag
import time
import pandas as pd
import os
import customtkinter
from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
#import database
#import os
#import shutil
#import pandas as pd
#im = pag.screenshot("fullscreen.png")
#im = pag.screenshot("partialscreen.png",region=(0,0,300,300))

app = customtkinter.CTk()
app.title('TỰ ĐỘNG GỬI SMS - VER 01 - HOÀNG HÙNG')
app.geometry('480x550')
app.config(bg='#161C25')
app.resizable(False,False)

font1 = ('Arial', 14,'bold')
font2 = ('Arial',10,'normal')

# Lấy tọa độ
def GET_POS():
    pass
#1. Thư mục gốc của phần mềm
current_dir = os.getcwd()

#print(current_dir)
# Kiểm tra hệ thống và bản quyền
def write_admin_file(number):
    try:
        with open(current_dir + '/image/hcl30784.txt', "w", encoding="utf-8") as f:
            f.writelines('314159\n')
            f.writelines(str(number))
        f.close()
    except:
        return 0
    
def read_admin_file():
    try:
        with open(current_dir + '/image/hcl30784.txt', "r", encoding="utf-8") as f:
            text = f.readlines()
        f.close()
        return [int(text[0]),int(text[1])]
    except:
        return [0,0]
def check_admin():
    admin_file = current_dir + '/image/hcl30784.txt'
    if (os.path.isfile(admin_file)):
        return True
    else:
        messagebox.showerror('Lỗi','Phần mềm đã hết hạn sử dụng!')
        return False
#2. Kiểm tra file excel chứa số điện thoại
EXCEL_FILE_PHONE = current_dir + '/sdt.xlsx'
#df = pd.read_excel(EXCEL_FILE_PHONE)
#mang_sdt = df['sdt'].dropna() 
#mang_content = df['content'].dropna()
def check_data():
    if (os.path.isfile(EXCEL_FILE_PHONE)):
        return True
    else:
        messagebox.showerror('Lỗi','Chưa có file dữ liệu')
        return False

#3. Kiểm tra danh sách dữ liệu cần chạy
def check_input():
    if (FROM_entry.get() and TO_entry.get()):
        try:            
            k_1 = int(FROM_entry.get())
            k_2 = int(TO_entry.get())
            df = pd.read_excel(EXCEL_FILE_PHONE)
            mang_sdt = df['sdt'].dropna()
            #mang_content = df['content'].dropna()
            if len(mang_sdt)<=0:
                messagebox.showerror('Lỗi', 'Danh sách dữ liệu trống!') 
                return False
            else:
                if(k_1>0 and k_2>k_1 and k_2 <= len(mang_sdt)):
                    return True
                else:
                    messagebox.showerror('Lỗi','Số thứ tự không hợp lệ!')
                    return False
        except ValueError:
            messagebox.showerror('Lỗi', 'Bạn cần nhập vào một số nguyên!')
            return 0
    else:
        messagebox.showerror('Lỗi', 'Bạn chưa nhập số thứ tự cần lấy')
#4. Kiểm tra thời gian delay
def check_delay():
    if DELAY_entry.get():
        try:
            delay = int(DELAY_entry.get())
            if ( delay >= 5):
                return True
            else:
                messagebox.showerror('Lỗi','Thời gian chờ >= 5 giây!')
        except ValueError:
            messagebox.showerror('Lỗi', 'Bạn cần nhập vào một số nguyên!')
            return 0
    else:
        messagebox.showerror('Lỗi','Chưa nhập thời gian chờ!')
        return False
#5. Load dữ liệu
def LOAD_DATA():
    if check_input():
        k_1 = int(FROM_entry.get())
        k_2 = int(TO_entry.get())
        df = pd.read_excel(EXCEL_FILE_PHONE)
        mang_sdt = df['sdt'].dropna()
        mang_content = df['content'].dropna()
        INFO_entry.delete(0,END)
        INFO_entry.insert(0,'Ready')
        tree.delete(*tree.get_children())
        for i in range(k_2-k_1+1):
            tree.insert('', END, values=[i+k_1,mang_sdt[i+k_1-1],mang_content[i+k_1-1]])
        return True
    else:
        #messagebox.showerror('Lỗi','Tải dữ liệu không thành công. Vui lòng kiểm tra các điều kiện và thử lại!')
        return False
#6. Hàm check hình ảnh
def click_img(loc_img):    
        btn = pag.locateCenterOnScreen(loc_img,confidence=0.8)
        pag.moveTo(btn,duration=1)
        pag.leftClick()   
def show_image_error():
        messagebox.showerror('Lỗi','Không tìm thấy hình ảnh. Chưa mở cửa sổ SMS hoặc thời gian chờ quá nhanh. Vui lòng thử lại!')  
      
#6. Hàm thực hiện vòng lặp
def VONG_LAP():
    time_delay = int(DELAY_entry.get())
    df = pd.read_excel(EXCEL_FILE_PHONE)
    mang_sdt = df['sdt'].dropna()
    mang_content = df['content'].dropna()
    k_1 = int(FROM_entry.get())
    k_2 = int(TO_entry.get())
    try:
        dem_sms = 0
        click_img('image/btn0.png')
        for i in range(k_2-k_1+1):                    
            time.sleep(1)
            try:
                click_img('image/btn_sms.png')                 
                time.sleep(2)
                pag.typewrite('0'+str(mang_sdt[i+k_1-1]))
                try:
                    click_img('image/btn_content.png')
                    time.sleep(2)
                    pag.typewrite(mang_content[i+k_1-1])
                    time.sleep(1)
                    try:
                        click_img('image/btn_gui.png')                    
                        time.sleep(time_delay)
                        dem_sms+=1
                    except:
                        show_image_error()
                        break
                except:
                    show_image_error()
                    break
            except:
                show_image_error()
                break
        return [dem_sms,k_2-k_1+1]
    except:
        show_image_error()
        return [0,-1]
#6. Hàm hiển thị kết quả sau khi kết thúc vòng lặp
def THE_END():
    INFO_entry.delete(0,END)
    INFO_entry.insert(0,'FINISH')
    messagebox.showinfo('Success', 'Đã chạy hết dữ liệu!')
#7. Hàm khởi đông quá trình
def START_IMAGE():
    if check_admin():
        [key_number,dem] = read_admin_file()
        if(dem>=0 and dem<100 and key_number==314159):
            if check_data():
                if check_delay():
                    if (INFO_entry.get()=='Ready'):
                        result = VONG_LAP()
                        if (result[0]==result[1]):
                            THE_END()
                        elif (result[0]<result[1]):
                            INFO_entry.delete(0,END)
                            INFO_entry.insert(0,'Đã gửi xong số thứ '+str(result[0]))
                            messagebox.showinfo('Lỗi', 'ĐÃ GỬI XONG SỐ THỨ '+str(result[0]))
                        else:
                            pass
                    else:
                        messagebox.showerror('Lỗi','Chưa load dữ liệu!')
            dem += 1
            write_admin_file(dem)   
        else:
            messagebox.showerror('Lỗi','Phần mềm đã hết hạn sử dụng!')
            os.remove(current_dir+'/image/hcl30784.txt')

def START_POS():
    pass
def START_HOTKEY():
    pass
#8. Hàm check điều kiện image
def display_data(event):
    pass
#Xây dựng giao diện

#POS1_label = customtkinter.CTkLabel(app,font=font1,text='P1:',text_color='#fff',bg_color='#161C25')
#POS1_label.place(x=5,y=45)

#POS1_entry = customtkinter.CTkEntry(app,font=font1,text_color='#000',fg_color='#fff',border_color='#0C9295',border_width=0,width=80)
#POS1_entry.place(x=30,y=45)

#POS2_label = customtkinter.CTkLabel(app,font=font1,text='P2:',text_color='#fff',bg_color='#161C25')
#POS2_label.place(x=5,y=85)
#POS2_entry = customtkinter.CTkEntry(app,font=font1,text_color='#000',fg_color='#fff',border_color='#0C9295',border_width=0,width=80)
#POS2_entry.place(x=30,y=85)

#POS3_label = customtkinter.CTkLabel(app,font=font1,text='P3:',text_color='#fff',bg_color='#161C25')
#POS3_label.place(x=5,y=125)
#POS3_entry = customtkinter.CTkEntry(app,font=font1,text_color='#000',fg_color='#fff',border_color='#0C9295',border_width=0,width=80)
#POS3_entry.place(x=30,y=125)

#POS4_label = customtkinter.CTkLabel(app,font=font1,text='P4:',text_color='#fff',bg_color='#161C25')
#POS4_label.place(x=5,y=165)
#POS4_entry = customtkinter.CTkEntry(app,font=font1,text_color='#000',fg_color='#fff',border_color='#0C9295',border_width=0,width=80)
#POS4_entry.place(x=30,y=165)

#GET_POS_button = customtkinter.CTkButton(app,command=GET_POS,font=font1,text_color='#fff',text='GET_POS',fg_color='#05A312',hover_color='#00850B',bg_color='#161C25',cursor='hand2',corner_radius=0,width=80)
#GET_POS_button.place(x=30,y=230)



FROM_label = customtkinter.CTkLabel(app,font=font1,text='FROM:',text_color='#fff',bg_color='#161C25')
FROM_label.place(x=5,y=455)

FROM_entry = customtkinter.CTkEntry(app,font=font1,text_color='#000',fg_color='#fff',border_color='#0C9295',border_width=0,width=45)
FROM_entry.place(x=65,y=455)

TO_label = customtkinter.CTkLabel(app,font=font1,text='TO:',text_color='#fff',bg_color='#161C25')
TO_label.place(x=125,y=455)

TO_entry = customtkinter.CTkEntry(app,font=font1,text_color='#000',fg_color='#fff',border_color='#0C9295',border_width=0,width=45)
TO_entry.place(x=160,y=455)

# Các nút chức năng
#PAUSE_button = customtkinter.CTkButton(app,command=PAUSE,font=font1,text_color='#fff',text='DỪNG',fg_color='#E40404',hover_color='#AE0000',bg_color='#161C25',border_color='#F15704',border_width=2,cursor='hand2',corner_radius=15,width=60)
#PAUSE_button.place(x=10,y=400)

INFO_entry = customtkinter.CTkEntry(app,font=font1,text_color='#000',fg_color='#fff',border_color='#0C9295',border_width=0,width=160)
INFO_entry.place(x=230,y=455)

LOAD_DATA_button = customtkinter.CTkButton(app,command=LOAD_DATA,font=font1,text_color='#fff',text='LOAD',fg_color='#05A312',hover_color='#00850B',bg_color='#161C25',cursor='hand2',corner_radius=0,width=60)
LOAD_DATA_button.place(x=410,y=455)


DELAY_label = customtkinter.CTkLabel(app,font=font1,text='DELAY:',text_color='#fff',bg_color='#161C25')
DELAY_label.place(x=5,y=500)

DELAY_entry = customtkinter.CTkEntry(app,font=font1,text_color='#000',fg_color='#fff',border_color='#0C9295',border_width=0,width=45)
DELAY_entry.place(x=65,y=500)

START_IMAGE_button = customtkinter.CTkButton(app,command=START_IMAGE,font=font1,text_color='#fff',text='START',fg_color='#05A312',hover_color='#00850B',bg_color='#161C25',cursor='hand2',corner_radius=15,width=150)
START_IMAGE_button.place(x=170,y=500)

#START_POS_button = customtkinter.CTkButton(app,command=START_POS,font=font1,text_color='#fff',text='BY POS',fg_color='#05A312',hover_color='#00850B',bg_color='#161C25',cursor='hand2',corner_radius=15,width=80)
#START_POS_button.place(x=150,y=500)

#START_HOTKEY_button = customtkinter.CTkButton(app,command=START_HOTKEY,font=font1,text_color='#fff',text='BY HOTKEY',fg_color='#05A312',hover_color='#00850B',bg_color='#161C25',cursor='hand2',corner_radius=15,width=80)
#START_HOTKEY_button.place(x=290,y=500)

####
style = ttk.Style(app)

style.theme_use('clam')
style.configure('Treeview',font=font2,foreground='#fff',background='#000',fieldbackground='#313837')
style.map('Treeview',bacground=[('selected', '#1A8F2D')])

tree = ttk.Treeview(app,height=20)

tree['columns'] = ('STT', 'PHONE', 'CONTENT')

tree.column('#0', width=0, stretch=tk.NO) # Hide the default first column
tree.column('STT', anchor=tk.CENTER,width=30)
tree.column('PHONE', anchor=tk.CENTER,width=120)
tree.column('CONTENT', anchor=tk.CENTER,width=310)

tree.heading('STT',text='STT')
tree.heading('PHONE',text='SỐ ĐIỆN THOẠI')
tree.heading('CONTENT',text='NỘI DUNG TIN NHẮN')

tree.place(x=5,y=5)

tree.bind('<ButtonRelease>', display_data)

app.mainloop()

