'''
Program pozwala wykorzystać otwartą sesję MS Outlook w celu zlecenia automatycznej wysyłki wiadomości. Każda z nich będzie zawierała załącznik z jednym, kolejnym plikiem ze wskazanego folderu.

Aktualnie w toku: funkcja zapisywania szablonów
'''

from os import walk
from tkinter import StringVar, IntVar, Frame, Tk, END, SUNKEN, Text, filedialog, messagebox
from tkinter.ttk import Label, Button, Entry, Radiobutton, Progressbar
from win32com.client import Dispatch

#Wczytanie szablonów i kont [importing templates nad Outlook accounts]:
lista_sz = []
for (dirpath, dirnames, filenames) in walk('Szablony'):
    lista_sz.extend(filenames)
szablon1 = ''
szablon2 = ''
szablon3 = ''
if len(lista_sz) >= 1:
    plik = open('Szablony\{}'. format(lista_sz[0]), encoding='utf8')
    for line in plik:
        szablon1 += line
if len(lista_sz) >= 2:
    plik = open('Szablony\{}'.format(lista_sz[1]), encoding='utf8')
    for line in plik:
        szablon2 += line
if len(lista_sz) >= 3:
    plik = open('Szablony\{}'.format(lista_sz[2]), encoding='utf8')
    for line in plik:
        szablon3 += line
outlook = Dispatch('outlook.application')
konta = []
for n in outlook.Session.Accounts:
    konta.append(n)

#Funkcje [Functions]:
def reset_text():
    txt_window.delete('1.0',END)
    if szablon.get() == 0:
        txt_window.insert(END, szablon1)
    if szablon.get() == 1:
        txt_window.insert(END, szablon2)
    if szablon.get() == 2:
        txt_window.insert(END, szablon3)
    reset_button()

def reset_button(*args):
    if ent_receiver.get() != '' and ent_folder.get() != '':
        btn_send['state'] = "enabled"
    else:
        btn_send['state'] = "disabled"

def quit():
    window.destroy()

def save_template():
    pass

def send():
    global konta
    pgr_suwak.grid(row=1, column=0, sticky='nsew')
    pliki = []
    progres.set(0)
    path = (fold.get())
    for (dirpath, dirnames, filenames) in walk(r'{}'.format(path)):
        pliki.extend(filenames)
    step = (100 / (len(pliki) + 1))
    progres.set(step)
    outlook = Dispatch('outlook.application')
    for i in range(0,len(pliki)):
        inf.set('Wysyłam {} z {} maili'.format(i + 1, len(pliki)))
        lbl_komunikat.update()
        mail = outlook.CreateItem(0)
        mail.To = receiver.get()
        mail.CC = receiver1.get()
        mail.BCC = receiver2.get()
        mail.Subject = '{}[cz. {}/{}]'. format(subject.get(), i+1, len(pliki))
        mail.Body = txt_window.get("1.0",END)
        attachment = '{}/{}'. format(path, pliki[i])
        mail.Attachments.Add(attachment)
        mail.SendUsingAccount = konta[wyb_konto.get()]
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, konta[wyb_konto.get()]))
        if sender.get() != '':
            mail.SentOnBehalfOfName = sender.get()
        mail.Send()
        progres.set((i+2) * step)
    messagebox.showinfo(title="WYSSSYŁACZ @", message='SSSkończyłem!!!')
    window.destroy()

def wybierz_folder():
    dirname = filedialog.askdirectory()
    if not dirname:
        return
    fold.set(r'{}'.format(dirname))
    reset_button()
    return dirname

def zdefiniuj_pliki():
    pass

#GUI:
window = Tk()
window.title("WYSSSYŁACZ @ v.BETA")
window.resizable(width=False, height=False)

#Zmienne:
inf = StringVar(value="* - pola wymagane")
sender = StringVar()
receiver = StringVar()
receiver1 = StringVar()
receiver2 = StringVar()
subject = StringVar()

#Ramki:
ramka_konto1 = Frame(borderwidth=5)
ramka_konto1.rowconfigure(0, minsize=25, weight=2)
ramka_konto1.columnconfigure(0, minsize=100, weight=1)
ramka_konto1.grid(row=0, column=0, sticky="nsew")
ramka_konto2 = Frame(borderwidth=5)
ramka_konto2.rowconfigure(0, minsize=25, weight=2)
ramka_konto2.columnconfigure([0,1], minsize=175, weight=1)
ramka_konto2.grid(row=0, column=1, sticky="nsew")
ramka = Frame(borderwidth=5)
ramka.rowconfigure([0, 1, 2, 3, 4, 5], minsize=25, weight=2)
ramka.columnconfigure(0, minsize=100, weight=1)
ramka.grid(row=1, column=0, sticky="nsew")
ramka1 = Frame(borderwidth=5, relief=SUNKEN)
ramka1.rowconfigure([0, 1, 2, 3, 4, 5], minsize=25, weight=2)
ramka1.columnconfigure(0, minsize=350, weight=1)
ramka1.grid(row=1, column=1, sticky="nsew")
ramka2 = Frame(borderwidth=5)
ramka2.rowconfigure([0, 1], minsize=25, weight=2)
ramka2.columnconfigure([0, 1, 2], minsize=150, weight=1)
ramka2.grid(columnspan=2, sticky="nsew")
ramka3 = Frame(borderwidth=5, relief=SUNKEN)
ramka3.rowconfigure(0, minsize=25, weight=2)
ramka3.columnconfigure(0, minsize=450, weight=1)
ramka3.grid(columnspan=2, sticky="nsew")
ramka4 = Frame(borderwidth=5)
ramka4.rowconfigure([0, 1], minsize=25, weight=2)
ramka4.columnconfigure(0, minsize=450, weight=1)
ramka4.grid(columnspan=2, sticky="nsew")

#Labelki:
lbl_konto = Label(master=ramka_konto1, text="Konto:")
lbl_konto.grid(row=0,column=0)
lbl_sender = Label(master=ramka, text="Nadawca:")
lbl_sender.grid(row=0, column=0)
lbl_receiver = Label(master=ramka, text="Odbiorca*:")
lbl_receiver.grid(row=1, column=0)
lbl_receiver1 = Label(master=ramka, text="DW:")
lbl_receiver1.grid(row=2, column=0)
lbl_receiver2 = Label(master=ramka, text="UDW:")
lbl_receiver2.grid(row=3, column=0)
lbl_subject = Label(master=ramka, text="Temat:")
lbl_subject.grid(row=4, column=0)
btn_folder = Button(master=ramka, text="Wybierz Folder*:", command=wybierz_folder)
btn_folder.grid(row=5, column=0, sticky="nsew")

#Konto i Enterki:
wyb_konto = IntVar(value=0)
if len(konta) >= 1:
    rb_konto1 = Radiobutton(master=ramka_konto2, text=konta[0].DisplayName, variable=wyb_konto, value=0)
    rb_konto1.grid(row=0,column=0)
if len(konta) >= 2:
    rb_konto1 = Radiobutton(master=ramka_konto2, text=konta[1].DisplayName, variable=wyb_konto, value=1)
    rb_konto1.grid(row=0, column=1)
ent_sender = Entry(master=ramka1, textvariable=sender)
ent_sender.grid(row=0, column=0, sticky="nsew")
ent_receiver = Entry(master=ramka1, textvariable=receiver)
ent_receiver.grid(row=1, column=0, sticky="nsew")
ent_receiver1 = Entry(master=ramka1, textvariable=receiver1)
ent_receiver1.grid(row=2, column=0, sticky="nsew")
ent_receiver2 = Entry(master=ramka1, textvariable=receiver2)
ent_receiver2.grid(row=3, column=0, sticky="nsew")
ent_subject = Entry(master=ramka1, textvariable=subject)
ent_subject.grid(row=4, column=0, sticky="nsew")

fold = StringVar(value="")
ent_folder = Entry(master=ramka1, textvariable=fold)
ent_folder.grid(row=5, column=0, sticky="nsew")

#Radio:
szablon = IntVar(value=0)
if len(lista_sz) == 0:
    lbl_brak_szablonow = Label(master=ramka2, text="Nie zdefiniowano szablonów")
    lbl_brak_szablonow.grid(row=0, column=0, columnspan=3)
if len(lista_sz) >= 1:
    rb_sz1 = Radiobutton(master=ramka2, text=lista_sz[0][:-4], variable=szablon, value=0, command=reset_text)
    rb_sz1.grid(row=0,column=0)
if len(lista_sz) >= 2:
    rb_sz2 = Radiobutton(master=ramka2, text=lista_sz[1][:-4], variable=szablon, value=1, command=reset_text)
    rb_sz2.grid(row=0, column=1)
if len(lista_sz) >= 3:
    rb_sz3 = Radiobutton(master=ramka2, text=lista_sz[2][:-4], variable=szablon, value=2, command=reset_text)
    rb_sz3.grid(row=0, column=2)

#Przyciski:
btn_send = Button(master=ramka2, text="Wyślij", state="disabled", command=send)
btn_send.grid(row=1, column=0, sticky="nsew")
btn_save = Button(master=ramka2, text="Zapisz szablon", command=save_template)
btn_save.grid(row=1, column=1, sticky="nsew")
btn_exit = Button(master=ramka2, text="Wyjdź", command=quit)
btn_exit.grid(row=1, column=2, sticky="nsew")

#Text:
txt_window = Text(master=ramka3)
txt_window.grid(row=0, column=0, columnspan=3, sticky="nsew")

#Dolna ramka:
lbl_komunikat = Label(master=ramka4, foreground="#8B0000", textvariable=inf)
lbl_komunikat.grid(row=0, column=0)
progres = IntVar(value=0)
pgr_suwak = Progressbar(master=ramka4, length=450, variable=progres, mode='determinate')

#Domyślny stan aplikacji
reset_text()
reset_button()
sender.trace_add("write", reset_button)
receiver.trace_add("write", reset_button)
receiver1.trace_add("write", reset_button)
receiver2.trace_add("write", reset_button)
subject.trace_add("write", reset_button)
fold.trace_add("write", reset_button)

window.mainloop()