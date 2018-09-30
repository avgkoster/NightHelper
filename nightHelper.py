from tkinter import *
from datetime import *
from tkinter import messagebox
from sendmails import OtprEmail
import openpyxl


class Helper:
    filename = 0
    mainFrame = Tk()   
    VidProd = Entry(width=30)
    VidProd.place(x=15,y=110)
    Dogovor = Entry(width=25)
    Dogovor.place(x=220,y=110)
    prices = Entry(width=10)
    prices.place(x=395,y=110)
    Zalob = Text(width=70,height=5)
    Zalob.place(x=15,y=160)
    
    def __init__(self,filename):
           #Глобальные переменные
            self.filename=filename 
            self.CreateElem(self.mainFrame,'calibri 12')
            self.mainFrame.geometry("600x350+500+300")
            self.mainFrame.resizable(False,False)
            self.mainFrame.mainloop()

    #Выгрузка таблиц
    def ReadExcel(self,variant):
        wbb = openpyxl.load_workbook(filename=self.filename,data_only=True)
        sheets = wbb['Лист1']
        if variant==0:
            smena = sheets['AG2'].value
            return smena
        elif variant==1:
            smena = sheets['AI2'].value
            return smena
        elif variant==2:
            smena = sheets['AK2'].value
            return smena
    
    

    def RunEmail(self):
        OtprEmail('Ночной отчет','kostyruzavin@mail.ru','avgkoster@gmail.com',self.filename,
                  'smtp.mail.ru','kostyruzavin@mail.ru','34650qqQQqwerty')
     #Добавление договора
    def NewDogovor(self):
        i=1
        if (len(self.VidProd.get())!=0 and len(self.Dogovor.get())!=0 and len(self.prices.get())!=0):
            wb = openpyxl.load_workbook(filename=self.filename,data_only=False)
            sheet = wb['Лист3']
            while(sheet['B'+str(i)].value!=None):
                i+=1
            sheet['A'+str(i)].value = datetime.now().strftime("%d.%m.%y")
            sheet['B'+str(i)].value = 'Рузавин К.В '
            sheet['C'+str(i)].value = self.VidProd.get()
            sheet['D'+str(i)].value = self.Dogovor.get()
            sheet['E'+str(i)].value = self.prices.get()
            sheet['F'+str(i)].value = float(self.prices.get())*0.04
            sheet['G'+str(i)].value = sheet['E'+str(i)].value
            wb.save(self.filename)
            self.CreateElem(self.mainFrame)
        else:
            messagebox.showerror("Ошибка", "Пожалуйста заполните все поля")



    #Счет процентов
    def FindProc(self):
        sum=0;i=1
        wb = openpyxl.load_workbook(filename=self.filename,data_only=True)
        sheet = wb['Лист3']
        while(sheet['B'+str(i)].value!=None):
            if sheet['B'+str(i)].value=='Рузавин К.В ':
                sum+=sheet['F'+str(i)].value
            i+=1
        return int(sum)


 
    #Элементы
    def CreateElem(self,Frame,typefont='calibri 12'):
        SmenButton = Button(Frame,width=70,text='Новая смена',
                        font=typefont,command=self.NewSmena)
        SmenButton.place(x=15,y=45)
        #Сегодняшняя дата
        DateText = Label(Frame,font=typefont,
                         text="Сегодня: "+
                         str(datetime.now().strftime("%d.%m.%y"))+" |")
        DateText.place(x=10,y=10)
        #Сегодняшняя смена
        SmenaText = Label(Frame,font=typefont,
                         text="Смена: "+
                         str(self.ReadExcel(0))+"  |")
        SmenaText.place(x=145,y=10)
        #Текущий заработок
        Zarab = Label(Frame,font=typefont,
                      text="Текущий заработок: "+str(self.ReadExcel(1))+"руб. |")
        Zarab.place(x=235,y=10)
        Procent = Label(Frame, font=typefont,
                      text="Процент: " + str(self.FindProc())+"руб.")
        Procent.place(x=460, y=10)
        #Ввод нового договора   
        VidProdLabel = Label(Frame, font=typefont,
                      text="Вид продажи: ")
        VidProdLabel.place(x=15, y=85)
        DogovorLabel = Label(Frame, font=typefont,
                      text="Номер договора: ")
        DogovorLabel.place(x=225, y=85)
        PriceLabel = Label(Frame,font=typefont,
                           text='Цена: ')
        PriceLabel.place(x=405,y=85)
        DogovorButton = Button(Frame,width=11,height=-5,text='Добавить',
                            font='calibri 12',command=self.NewDogovor)
        DogovorButton.place(x=485,y=95)
        #Очистка договоров
        self.VidProd.delete(0,END)
        self.Dogovor.delete(0,END)
        self.prices.delete(0,END)
        #Добавление жалоб
        ZalobLabel = Label(Frame,font = typefont,
                           text="Добавление замечаний: ")
        ZalobLabel.place(x=15,y=130)
        ZalobButton = Button(text="Добавить замечание(Новых замечаний нет)",width=79,
                             command=self.NewZalob)
        ZalobButton.place(x=15,y=255)
        #Отправка почты
        EmailButton = Button(text="Отправить отчет",width=79,command=self.RunEmail)
        EmailButton.place(x=15,y=295)

        #Загрузка таблиц
    def NewSmena(self):
        bools = 0
        wb = openpyxl.load_workbook(filename=self.filename,data_only=True)   
        sheet = wb['Лист1']
        for get in sheet["A2:AE2"]:
            if bools==1:
                break
            for i in range(1,30):
                if get[i].value=='x':
                    sheet.cell(row=2, column=i+1).value = 1
                    sheet["AG2"].value=sheet["AG2"].value+1
                    sheet["AI2"].value=sheet["AG2"].value*sheet["AH2"].value
                    wb.save('fick1.xlsx')
                    self.CreateElem(self.mainFrame)
                    bools=1
                    break

    #Добавление жалоб
    def NewZalob(self):
        i=1
        wb = openpyxl.load_workbook(filename=self.filename,data_only=False)
        sheet = wb['Лист2']
        while(sheet['B'+str(i)].value!=None):
            i+=1
        if(len(self.Zalob.get("1.0",'end-1c'))!=0):   
            sheet['A'+str(i)].value = datetime.now().strftime("%d.%m.%y")
            sheet['B'+str(i)].value = "Рузавин К.В"
            sheet['C'+str(i)].value = self.Zalob.get("1.0",'end-1c')
        else:
            sheet['A'+str(i)].value = datetime.now().strftime("%d.%m.%y")
            sheet['B'+str(i)].value = "Рузавин К.В"
            sheet['C'+str(i)].value = "Новых замечаний нет"    
        wb.save(self.filename)
        self.CreateElem(self.mainFrame)  

           

newWindow = Helper('fick1.xlsx')

    

    


