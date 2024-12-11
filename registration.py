from tkinter import *
import tkinter
import openpyxl
screen=tkinter.Tk()

screen.title('Registration')
screen.geometry('600x500')
screen.configure(bg='#ececec')
#kiểm tra thông tin đăng kí có trong dữ liệu không
def check_duplicate(lst,lst_tong):
    if lst in lst_tong:
        return True
    return False
#kiểm tra  email hợp lệ hay không
def check_email(em):
    lst_mail=['@gmail.com','@fpt.edu','@outlook.com']
    for i in lst_mail:
        if i in em:
            return False
    return True

#kiểm tra phone chỉ chứa int
def check_phone(ph):
    if ph.isnumeric():
        return False 
    return True

def register():
    try:
        workbook=openpyxl.load_workbook('Register Form.xlsx')
        sheet=workbook.active
    except FileNotFoundError:
        workbook=openpyxl.Workbook()
        sheet= workbook.active   
    sheet['A1']='Name'
    sheet['B1']='Phone'
    sheet['C1']='Gender'
    sheet['D1']='Email'
    
    nm=nameValue.get()
    ph=phoneValue.get()
    gen=genderValue.get()
    em=emailValue.get()
    print(type(em))
    #kiểm tra trùng lặp
    lst=[nm,ph,gen,em]  
    lst_tong=[]
    for cell_rows in sheet.rows:
        giatri=','.join(str(cell.value) for cell in cell_rows)
        lst_tong.append(giatri.split(','))
    if check_email(em):
        success_info.place_forget()
        duplicate_info.place_forget()
        phone_int.place_forget()
        email_type.place(x=210,y=100) 
    elif check_phone(ph):
        success_info.place_forget()
        duplicate_info.place_forget()
        email_type.place_forget()
        phone_int.place(x=170,y=100)
    else:   
        if check_duplicate(lst,lst_tong):
            success_info.place_forget()
            email_type.place_forget()
            phone_int.place_forget()
            duplicate_info.place(x=130,y=100)
        else:
            sheet.append(lst)
            duplicate_info.place_forget()
            email_type.place_forget()
            phone_int.place_forget()
            success_info.place(x=210,y=100)
    #canh chỉnh cột
    for column_cells in sheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        sheet.column_dimensions[column_cells[0].column_letter].width=length + 2
    
    workbook.save('Register Form.xlsx')
    
#thiết kê giao diện đăng kí
regis=Label(screen,text='Python registation form',font=('Arial',30),bg='#ececec').pack(side=tkinter.TOP,pady=30)


nameValue=StringVar()
name=Label(screen,text='Name',font=('Arial',15),bg='#ececec').place(x=100,y=150)
name_input=Entry(screen,font=('Arial',15),width=30,textvariable=nameValue).place(x=200,y=150)

phoneValue=StringVar()
phone=Label(screen,text='Phone',font=('Arial',15),bg='#ececec').place(x=100,y=200)
phone_input=Entry(screen,font=('Arial',15),width=30,textvariable=phoneValue).place(x=200,y=200)

genderValue=StringVar()
gender=Label(screen,text='Gender',font=('Arial',15),bg='#ececec').place(x=100,y=250)
gender_input=Entry(screen,font=('Arial',15),width=30,textvariable=genderValue).place(x=200,y=250)

emailValue=StringVar()
email=Label(screen,text='Email',font=('Arial',15),bg='#ececec').place(x=100,y=300)
email_input=Entry(screen,font=('Arial',15),width=30,textvariable=emailValue).place(x=200,y=300)

checkValue=IntVar()
check_button=Checkbutton(screen,text='Remember me?',font=('Arial',10),bg='#ececec',variable=checkValue).place(x=210,y=350)

#kiểm tra thông tin đăng kí
duplicate_info=Label(screen,bg='#ececec',text='Your form has been already registered',fg='red',font=('arial',14))
success_info=Label(screen,bg='#ececec',text='Register success',fg='green',font=('arial',14))
email_type=Label(screen,bg='#ececec',text='Invalid Email',fg='red',font=('arial',14))
phone_int=Label(screen,bg='#ececec',text='Invalid Phone number',fg='red',font=('arial',14))

Button(screen,text='Register',font=('Arial',20),bg='#ececec',command=register).pack(side=tkinter.BOTTOM,pady=50)


screen.mainloop()

