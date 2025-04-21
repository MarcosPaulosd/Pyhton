from gettext import find
from tkinter import *
from tkcalendar import DateEntry
from tkinter import ttk
from lendo_planilha import lendo
from telas import telas_cadastro

def reservar():
    lista = []
    error = False
    tec = eidtecnico.get()
    fer = eidferramenta.get()
    if tec == 'Selecione o Tecnico     -- > ':
        error = True
        errorreservar['text'] = 'Preencha o campo corretamente'
        return
    else:
        errorreservar['text'] = ''
        e = tec.find(' ')
        r = tec[:e]
        lista.append(r)

    if fer == 'Selecione a ferramenta     -- > ':
        error = True
        errorreservar['text'] = 'Preencha o campo corretamente'
        return
    else:
        errorreservar['text'] = ''
        ee = fer.find(' ')
        rr = fer[:ee]
        lista.append(rr)

    datab = edatabusca.get()
    datae = edataentrega.get()

    lista.append(datab)
    lista.append(datae)


    lendo.reservando(lista)


    eidferramenta.set('Selecione a ferramenta     -- > ')
    eidtecnico.set('Selecione o Tecnico     -- > ')


    reservaconcluida['text'] = 'Reserva concluida'

    for item in tv.get_children():
        tv.delete(item)

    lendo.lendo_planilha(tv)
    
    
def devolucao_ferramenta():

    
        


    error = False
    lista = []


    ferdev = eidferramentadev.get()
    if ferdev == 'Selecione a ferramenta':
        error = True
        errordev['text'] = 'Ferramenta error'
    else:
        errordev['text'] = ''
        lista.append(ferdev)

    motivo = emotivo.get()
    if motivo == 'Selecione o motivo':
        error = True
        errormotivo['text'] = 'Motivo error'
    else:
        errormotivo['text'] = ''
        lista.append(motivo)

    
    lendo.devolvendo(lista,tv)
    eidferramentadev.set('Selecione a ferramenta')
    emotivo.set('Selecione o motivo')
    listadev = lendo.combobox_devolucao()

    return listadev




    
    

    





tl = Tk()
tl.title('Controle de reserva')
tl.geometry('1024x700')

tl.resizable(width=False,height=False)

fce = Frame(width=318,height=104,bg='#FFFFFF',relief='flat')
fce.grid(column=0,row=0)

fcd = Frame(width=706,height=104,bg = '#836FFF' , relief='flat')
fcd.grid(column=1,row=0)

fbe = Frame(width=318,height=596,bg = '#DEDBE8', relief='flat')
fbe.grid(column=0,row=1)

fbd = Frame(width=706,height=596,bg = '#4169E1',relief='flat')
fbd.grid(column=1,row=1)

tp = Label(fce, text='Controle de reserva',bg = '#FFFFFF', font=('Consolas 21 bold'))
tp.place(x = 12,y=35)

Label(fbe,text='Id do Técnico',bg='#DEDBE8', fg = '#000000',font=('Consolas 15 bold')).place(x=16,y=10)

listatec = lendo.comboboxtecnico()
eidtecnico = ttk.Combobox(fbe,width=40,values=listatec)
eidtecnico.set('Selecione o Tecnico     -- > ')
##eidtecnico = Entry(fbe,width=45)
eidtecnico.place(x=20,y=52)

listaferr = lendo.comboboxferramenta()
Label(fbe,text='Id da ferramenta',bg='#DEDBE8', fg = '#000000',font=('Consolas 15 bold')).place(x=16,y=95)
eidferramenta= ttk.Combobox(fbe,width=40,values=listaferr)
eidferramenta.set('Selecione a ferramenta     -- > ')
##eidferramenta = Entry(fbe,width=45)
eidferramenta.place(x=20,y=139)

Label(fbe,text='Data de busca',bg='#DEDBE8', fg = '#000000',font=('Consolas 12 bold')).place(x=16,y=169)
edatabusca = DateEntry(fbe)
edatabusca.place(x=20,y=196)

Label(fbe,text='Data de entrega',bg='#DEDBE8', fg = '#000000',font=('Consolas 12 bold')).place(x=169,y=169)
edataentrega = DateEntry(fbe)
edataentrega.place(x=169,y=196)

errorreservar = Label(fbe,text='',bg = '#DEDBE8',fg='#FF0000')
errorreservar.place(x=200,y=260)

botaoreservar = Button(fbe,text='Reservar',width=28,command=reservar)
botaoreservar.place(x=55,y=230)

Label(fbe,text='Devolução',bg='#DEDBE8', fg = '#000000',font=('Consolas 12 bold')).place(x=10,y=257)
listadev = lendo.combobox_devolucao()
eidferramentadev = ttk.Combobox(fbe,values=listadev)
eidferramentadev.set('Selecione a ferramenta')
eidferramentadev.place(x=55,y=290)

errordev = Label(fbe,text='',bg = '#DEDBE8',fg='#FF0000')
errordev.place(x=200,y=288)



listamotivosdev = ['Selecione o motivo','Cancelando','Devolvendo']
emotivo = ttk.Combobox(fbe,values=listamotivosdev)
emotivo.set('Selecione o motivo')
emotivo.place(x=55,y=320)

errormotivo = Label(fbe,text='',bg = '#DEDBE8',fg='#FF0000')
errormotivo.place(x=200,y=320)


botaodevolucao = Button(fbe,text='Devolver',command=devolucao_ferramenta)
botaodevolucao.place(x=90,y=350)



reservaconcluida = Label(fbe,text='',bg = '#DEDBE8',fg='#85bb65')
reservaconcluida.place(x=190,y=260)

Label(fbe,text='Cadastrar',bg='#DEDBE8', fg = '#000000',font=('Consolas 15 bold')).place(x=110,y=426)

bcadastferram = Button(fbe,text='Ferramenta',width=15,command=telas_cadastro.jan_cadastro_ferramenta)
bcadastferram.place(x=17,y=459)

bcadasttec = Button(fbe,text='Técnico',width=15,command=telas_cadastro.janela_cadastro_tecnico)
bcadasttec.place(x=186,y=459)

Label(fbe,text='Extrair',bg='#DEDBE8', fg = '#000000',font=('Consolas 15 bold')).place(x=120,y=499)

beferramenta = Button(fbe,text='Ferramenta',width=10,command=lendo.extraindo_ferramentas)
beferramenta.place(x=10,y=540)

betecnico = Button(fbe,text='tecnico',width=10,command=lendo.extraindo_tecnicos)
betecnico.place(x=120,y=540)

behistorico = Button(fbe,text='historico',width=10, command=lendo.extraindo_historico)
behistorico.place(x=230,y=540)





Label(fcd,text='Ferramentas reservadas',bg = '#836FFF', font=('Consolas 30 bold')).place(x=125,y=29)

tv = ttk.Treeview(fbd,height=496,columns=('id','ferramenta','idtec','nometec','telefonetec','dataentrega'),show='headings')

tv.column('id',minwidth=0,width=70)
tv.column('ferramenta',minwidth=0,width=124)
tv.column('idtec',minwidth=0,width=70)
tv.column('nometec',minwidth=0,width=148)
tv.column('telefonetec',minwidth=0,width=148)
tv.column('dataentrega',minwidth=0,width=146)

tv.heading('id',text='ID ferrmanta')
tv.heading('ferramenta',text='Ferramenta')
tv.heading('idtec',text='ID Tec')
tv.heading('nometec',text='Nome')
tv.heading('telefonetec',text='Telefone Tec')
tv.heading('dataentrega',text='Data entrega')

tv.place(x=0,y=0)

lendo.lendo_planilha(tv)

tl.mainloop()