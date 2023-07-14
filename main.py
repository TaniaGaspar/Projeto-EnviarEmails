from tkinter import * #para importar imagem de fundo e o END no apagar assunto
import customtkinter
import re
from PIL import Image #para importar imagens

import win32com.client as win32 # biblioteca para abrir o e-mail
import pandas as pd

from flask import Flask, render_template, request
import threading
import webbrowser

#aparencia geral escura
customtkinter.set_appearance_mode("dark") #dark | light| system (para MacOs)
#customtkinter.set_default_color_theme("dark-blue")

#configuração da janela inicial
janela = customtkinter.CTk()
janela.geometry("400x550+100+10") # Largura x Altura + distancia da esquerda + distancia da direita
janela.title("Projeto e-mails")
janela.iconbitmap("icon.ico")

imagem_fundo = customtkinter.CTkImage(light_image=Image.open("fundo.png"), dark_image=Image.open("fundo.png"), size=(1600,900))
customtkinter.CTkLabel(janela, text=None, image=imagem_fundo).place(x=0,y=0)

#isto é para o editor de mensagem WYSIWYG funcionar, é para abrir a página web
app = Flask(__name__)
@app.route("/", methods=['GET', 'POST'])
def index(): 
    with open("corpo_email.txt", "r", encoding = "utf-8") as corpo_email_arq:
        corpo_email = corpo_email_arq.read()
    with open("assunto.txt", "r", encoding = "utf-8") as assunto_arq:
        assunto = assunto_arq.read()
    if request.method == 'POST':
        mensagem_recebida=request.form.get("editordata_mensagem")
        assunto_recebido = request.form.get("editordata_assunto")
        #print("mensagem: ",mensagem_recebida)
        with open("corpo_email.txt","w", encoding = "utf-8") as guardar_corpo_texto_arq:       
            guardar_corpo_texto = guardar_corpo_texto_arq.write(mensagem_recebida)
        with open("assunto.txt","w", encoding = "utf-8") as guardar_assunto_arq:        
            guardar_assunto = guardar_assunto_arq.write(assunto_recebido)
        return 'Mensagem Guardada'
    return render_template("index.html", mensagem_mostrar = corpo_email, assunto_mostrar=assunto)

# B O T Ã O   E N V I A R   E M A I L 
def clique_enviar_emails():
    try:
        # T E M P O   E N T R E   E N V I O S   D E   E M A I L S - ciclo infino enquanto for true
        with open("emails_automaticos_tempo_ativo.txt", "r", encoding = "utf-8") as emails_automaticos_tempo_arq:
            emails_automaticos_tempo_ativo = emails_automaticos_tempo_arq.read()
        #while emails_automaticos_tempo_ativo == "True":
        if emails_automaticos_tempo_ativo =="True":
            with open("emails_automaticos_tempo.txt", "r", encoding = "utf-8") as emails_automaticos_tempo_arq:
                tempo_min = emails_automaticos_tempo_arq.read()
            tempo_mil_sec= int(tempo_min)*60*1000
            #print("tempo entre envios:",tempo_min, "minutos")
            #print("tempo entre envios:", tempo_mil_sec, "miliseconds")
            janela.after(tempo_mil_sec,lambda: clique_enviar_emails())

        # primeiro importar a lista de emails
        tabela_emails= pd.read_excel("dados_emails.xlsx")
        emails_enviar = tabela_emails['emails'].tolist()
        
        if emails_enviar[-1] != "@@@@@":
            emails_enviar.append("@@@@@")
        #print("lista com marca final", emails_enviar)
        #print("to list:", emails_enviar)
        outlook = win32.Dispatch('outlook.application') #ligar com o outlook
        
        #variáveis em txt dentro do email
        with open("assunto.txt", "r", encoding = "utf-8") as assunto_arq:
            assunto = assunto_arq.read()
        with open("corpo_email.txt", "r", encoding = "utf-8") as corpo_email_arq:
            corpo_email = corpo_email_arq.read()
        with open("numero_envios.txt", "r", encoding = "utf-8") as numero_envios_arq:
            n_requerido_envios = int(numero_envios_arq.read())

        #variáveis uteis
        regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
        numero_envios=1
        n_requerido_envios2= n_requerido_envios
        #print("numero de envios", n_requerido_envios)
        for email_enviar in emails_enviar:
            #print(email_enviar)
            if re.fullmatch(regex,email_enviar):
                if numero_envios <= n_requerido_envios:
                    numero_envios +=1
                    #criar um email
                    email = outlook.CreateItem(0) #associar o remetente
                    #cofigurar o email
                    email.To = email_enviar
                    email.Subject = f"{assunto}"
                    email.HTMLBody = f"{corpo_email}"

                    #abrir a localização dos anexos
                    with open("anexo1.txt", "r", encoding = "utf-8") as anexo1_arq:
                        anexo1 = anexo1_arq.read()
                    with open("anexo2.txt", "r", encoding = "utf-8") as anexo2_arq:
                        anexo2 = anexo2_arq.read()
                    with open("anexo3.txt", "r", encoding = "utf-8") as anexo3_arq:
                        anexo3 = anexo3_arq.read()
                    with open("anexo4.txt", "r", encoding = "utf-8") as anexo4_arq:
                        anexo4 = anexo4_arq.read()
                    with open("anexo5.txt", "r", encoding = "utf-8") as anexo5_arq:
                        anexo5 = anexo5_arq.read()
                
                    #enviar anexos se eles existirem
                    if anexo1 != " " or None:
                        email.Attachments.Add(anexo1)
                    if anexo2 != " " or None:
                        email.Attachments.Add(anexo2)
                    if anexo3 != " " or None:    
                        email.Attachments.Add(anexo3)                    
                    if anexo4 != " " or None:    
                        email.Attachments.Add(anexo4)
                    if anexo5 != " " or None:    
                        email.Attachments.Add(anexo5)

                    email.Send()
                else:
                    #criar nova tela
                    janela_enviar_email = customtkinter.CTkToplevel(janela)
                    janela_enviar_email.title("Projeto e-mails")
                    janela_enviar_email.geometry("300x150+500+100")
                    #caixa de texto    
                    muda_assunto_label = customtkinter.CTkLabel(master = janela_enviar_email, text = "E-mails enviados !", font = customtkinter.CTkFont(size=16, weight = "bold"))
                    muda_assunto_label.pack(padx=10, pady=(20,10))

                    botao_voltar = customtkinter.CTkButton(janela_enviar_email, text = " Voltar ", corner_radius=0, command = janela_enviar_email.destroy)
                    botao_voltar.pack(padx=10, pady=10)

                    #fechar janela automaticamente
                    janela.after(10000,lambda: janela_enviar_email.destroy())
                    break
            elif re.fullmatch("@@@@@",email_enviar):
                with open("reposicao_automatica.txt", "r", encoding = "utf-8") as reposicao_automatica_arq:
                    reposicao_automatica = reposicao_automatica_arq.read()
                if reposicao_automatica == "True":
                    #abrir excel
                    tabela= pd.read_excel("dados_emails - completo.xlsx")
                    emails = tabela['emails'] #puxar colunas       
                    #guardar excel 
                    file_name = 'dados_emails.xlsx'
                    emails.to_excel(file_name)  

                    janela_reposicao_automatica = customtkinter.CTkToplevel(janela, fg_color="red")
                    janela_reposicao_automatica.title("Projeto e-mails - Repor lista automatica")
                    janela_reposicao_automatica.geometry("300x150+500+100")
                    #caixa de texto    
                    reposicao_automatica_label = customtkinter.CTkLabel(master = janela_reposicao_automatica, text = "Emails enviados \nLista de e-mails terminou!", font = customtkinter.CTkFont(size=16, weight = "bold"))
                    reposicao_automatica_label.pack(padx=10, pady=(20,10))

                    botao_voltar = customtkinter.CTkButton(janela_reposicao_automatica, text = " Voltar ", corner_radius=0, command = janela_reposicao_automatica.destroy)
                    botao_voltar.pack(padx=10, pady=10)
                    break
                else:

                    with open("email_para_avisos.txt", "r", encoding = "utf-8") as email_para_avisos_arq:
                        email_para_avisos = email_para_avisos_arq.read()
                    email = outlook.CreateItem(0) #associar o remetente
                    #cofigurar o email
                    email.To = email_para_avisos
                    email.Subject = f"Lista Terminada"
                    email.HTMLBody = f"A lista de e-mails terminou e por isso o envio parou."
                    email.Send()

                    janela_tabela_vazia = customtkinter.CTkToplevel(janela,fg_color="red")
                    janela_tabela_vazia.title("Projeto e-mails - Repor lista automatica")
                    janela_tabela_vazia.geometry("400x150+500+100")
                    #caixa de texto    
                    tabela_vazia_label = customtkinter.CTkLabel(master = janela_tabela_vazia, text = "Lista terminou, importe uma nova", font = customtkinter.CTkFont(size=16, weight = "bold"))
                    tabela_vazia_label.pack(padx=10, pady=(20,10))

                    botao_voltar = customtkinter.CTkButton(janela_tabela_vazia, text = " Voltar ", corner_radius=0, command = janela_tabela_vazia.destroy)
                    botao_voltar.pack(padx=10, pady=10)
            else: 
                #email critico -- estragado
                n_requerido_envios2 += 1
        
        dicionario_emails_enviar= {'emails':emails_enviar}
        emails_enviar_df = pd.DataFrame(dicionario_emails_enviar )
        #n_requerido_envios2 += 1
        emails_enviar_df= emails_enviar_df.iloc[n_requerido_envios2:, : ]

        file_name_por_enviar = 'dados_emails.xlsx'
        emails_enviar_df.to_excel(file_name_por_enviar)

    except UnboundLocalError:
        #print("letra no numero de emails")
        erro_n_emails = customtkinter.CTkToplevel(janela, fg_color="red")
        erro_n_emails.title("Projeto e-mails - Erro")
        erro_n_emails.geometry("300x150+500+100")

        erro_n_emails_label = customtkinter.CTkLabel(master = erro_n_emails, text = "Erro no numero de e-mails.\n Certifique-se que tem lá um número", font = customtkinter.CTkFont(size=16, weight = "bold"))
        erro_n_emails_label.pack(padx=10, pady=(20,10))

    except KeyError:
        #print("keyerror")
        #reposicao_automatica = True
        with open("reposicao_automatica.txt", "r", encoding = "utf-8") as reposicao_automatica_arq:
            reposicao_automatica = reposicao_automatica_arq.read()
        if reposicao_automatica == "True":
            #abrir excel
            tabela= pd.read_excel("dados_emails - completo.xlsx")
            emails = tabela[['emails','flag']] #puxar colunas       
            #guardar excel 
            file_name = 'dados_emails.xlsx'
            emails.to_excel(file_name)  

            janela_reposicao_automatica = customtkinter.CTkToplevel(janela, fg_color="red")
            janela_reposicao_automatica.title("Projeto e-mails - Repor lista automatica")
            janela_reposicao_automatica.geometry("300x150+500+100")
            #caixa de texto    
            reposicao_automatica_label = customtkinter.CTkLabel(master = janela_reposicao_automatica, text = "Emails enviados \nLista de e-mails reposta!", font = customtkinter.CTkFont(size=16, weight = "bold"))
            reposicao_automatica_label.pack(padx=10, pady=(20,10))

            botao_voltar = customtkinter.CTkButton(janela_reposicao_automatica, text = " Voltar ", corner_radius=0, command = janela_reposicao_automatica.destroy)
            botao_voltar.pack(padx=10, pady=10)
        else:
            tabela= pd.read_excel("marca_final.xlsx")
            emails = tabela[['emails']] #puxar colunas
                
            file_name = 'dados_emails.xlsx'
            emails.to_excel(file_name)  

            with open("email_para_avisos.txt", "r", encoding = "utf-8") as email_para_avisos_arq:
                email_para_avisos = email_para_avisos_arq.read()
            email = outlook.CreateItem(0) #associar o remetente
            #cofigurar o email
            email.To = email_para_avisos
            email.Subject = f"Lista Terminada"
            email.HTMLBody = f"A lista de e-mails terminou e por isso o envio parou."
            email.Send()

            janela_sem_reposicao_automatica = customtkinter.CTkToplevel(janela,fg_color="red")
            janela_sem_reposicao_automatica.title("Projeto e-mails - Repor lista automatica")
            janela_sem_reposicao_automatica.geometry("400x150+500+100")
            #caixa de texto    
            reposicao_automatica_label = customtkinter.CTkLabel(master = janela_sem_reposicao_automatica, text = "Lista terminou, importe uma nova", font = customtkinter.CTkFont(size=16, weight = "bold"))
            reposicao_automatica_label.pack(padx=10, pady=(20,10))

            botao_voltar = customtkinter.CTkButton(janela_sem_reposicao_automatica, text = " Voltar ", corner_radius=0, command = janela_sem_reposicao_automatica.destroy)
            botao_voltar.pack(padx=10, pady=10)

    except IndexError:
        print("indexError")
        with open("reposicao_automatica.txt", "r", encoding = "utf-8") as reposicao_automatica_arq:
            reposicao_automatica = reposicao_automatica_arq.read()
        if reposicao_automatica == "True":
            #abrir excel
            tabela= pd.read_excel("dados_emails - completo.xlsx")
            emails = tabela['emails'] #puxar colunas       
            #guardar excel 
            file_name = 'dados_emails.xlsx'
            emails.to_excel(file_name) 

            with open("email_para_avisos.txt", "r", encoding = "utf-8") as email_para_avisos_arq:
                email_para_avisos = email_para_avisos_arq.read()
            
            outlook = win32.Dispatch('outlook.application') #ligar com o outlook
            email = outlook.CreateItem(0) #associar o remetente
            #cofigurar o email
            email.To = email_para_avisos
            email.Subject = f"Lista Terminada"
            email.HTMLBody = f"A lista de e-mails terminou e foi reposta, por isso o envio continua."
            email.Send() 

            janela_reposicao_automatica = customtkinter.CTkToplevel(janela, fg_color="red")
            janela_reposicao_automatica.title("Projeto e-mails - Repor lista automatica")
            janela_reposicao_automatica.geometry("300x150+500+100")
            #caixa de texto    
            reposicao_automatica_label = customtkinter.CTkLabel(master = janela_reposicao_automatica, text = "Emails enviados \nLista de e-mails reposta!", font = customtkinter.CTkFont(size=16, weight = "bold"))
            reposicao_automatica_label.pack(padx=10, pady=(20,10))

            botao_voltar = customtkinter.CTkButton(janela_reposicao_automatica, text = " Voltar ", corner_radius=0, command = janela_reposicao_automatica.destroy)
            botao_voltar.pack(padx=10, pady=10)
        else:
            with open("email_para_avisos.txt", "r", encoding = "utf-8") as email_para_avisos_arq:
                email_para_avisos = email_para_avisos_arq.read()
            outlook = win32.Dispatch('outlook.application') #ligar com o outlook
            email = outlook.CreateItem(0) #associar o remetente
            #cofigurar o email
            email.To = email_para_avisos
            email.Subject = f"Lista Terminada"
            email.HTMLBody = f"A lista de e-mails terminou e por isso o envio parou."
            email.Send()

            janela_tabela_vazia = customtkinter.CTkToplevel(janela,fg_color="red")
            janela_tabela_vazia.title("Projeto e-mails - Repor lista automatica")
            janela_tabela_vazia.geometry("400x150+500+100")
            #caixa de texto    
            tabela_vazia_label = customtkinter.CTkLabel(master = janela_tabela_vazia, text = "Lista terminou, importe uma nova", font = customtkinter.CTkFont(size=16, weight = "bold"))
            tabela_vazia_label.pack(padx=10, pady=(20,10))

            botao_voltar = customtkinter.CTkButton(janela_tabela_vazia, text = " Voltar ", corner_radius=0, command = janela_tabela_vazia.destroy)
            botao_voltar.pack(padx=10, pady=10)
        
       
    #except Exception as error:   
        

# B O T Ã O   E N V I A R   E M A I L   T E S T E
def clique_enviar_email_teste():
    try:
        outlook = win32.Dispatch('outlook.application') #ligar com o outlook

        #variáveis em txt dentro do email
        with open("assunto.txt", "r", encoding = "utf-8") as assunto_arq:
            assunto = assunto_arq.read()
        with open("corpo_email.txt", "r", encoding = "utf-8") as corpo_email_arq:
            corpo_email = corpo_email_arq.read()
        with open("email_teste.txt", "r", encoding = "utf-8") as teste_email_arq:
            teste_email = teste_email_arq.read()
        
        #criar um email
        email = outlook.CreateItem(0) #associar o remetente
        #cofigurar o email
        email.To = teste_email
        email.Subject = f"{assunto}"
        #utilizar formatação HTML a utilizar no e-mail - Escrever email normal
        email.HTMLBody = f"{corpo_email}"
        
        #abrir a localização dos anexos
        with open("anexo1.txt", "r", encoding = "utf-8") as anexo1_arq:
            anexo1 = anexo1_arq.read()
        with open("anexo2.txt", "r", encoding = "utf-8") as anexo2_arq:
            anexo2 = anexo2_arq.read()
        with open("anexo3.txt", "r", encoding = "utf-8") as anexo3_arq:
            anexo3 = anexo3_arq.read()
        with open("anexo4.txt", "r", encoding = "utf-8") as anexo4_arq:
            anexo4 = anexo4_arq.read()
        with open("anexo5.txt", "r", encoding = "utf-8") as anexo5_arq:
            anexo5 = anexo5_arq.read()
    
        #enviar anexos se eles existirem
        if anexo1 != " " or None:
            email.Attachments.Add(anexo1)

        if anexo2 != " " or None:
            email.Attachments.Add(anexo2)

        if anexo3 != " " or None:    
            email.Attachments.Add(anexo3)
        
        if anexo4 != " " or None:    
            email.Attachments.Add(anexo4)

        if anexo5 != " " or None:    
            email.Attachments.Add(anexo5)
        email.Send()
        
        #janela a confirmar envio
        janela_enviar_email = customtkinter.CTkToplevel(janela)
        janela_enviar_email.title("Projeto e-mails - E-mail de Teste")
        janela_enviar_email.geometry("300x150+500+100")
        #caixa de texto    
        muda_assunto_label = customtkinter.CTkLabel(master = janela_enviar_email, text = "E-mail de teste enviado", font = customtkinter.CTkFont(size=16, weight = "bold"))
        muda_assunto_label.pack(padx=10, pady=(20,10))

        botao_voltar = customtkinter.CTkButton(janela_enviar_email, text = " Voltar ", corner_radius=0, command = janela_enviar_email.destroy)
        botao_voltar.pack(padx=10, pady=10)

    except Exception as error:
        #janela a informar que houve um erro
        #print("O campo do e-mail de teste encontra-se vazio!")
        janela_enviar_teste_vazio = customtkinter.CTkToplevel(janela, fg_color="red")
        janela_enviar_teste_vazio.title("Projeto e-mails - Erro")
        janela_enviar_teste_vazio.geometry("300x200+500+100")
        #caixa de texto    
        muda_assunto_label = customtkinter.CTkLabel(master = janela_enviar_teste_vazio, text = "O campo de email de teste\nencontra-se vazio, ou preenchido \nde forma incorreta,\n ou há um problema com os anexos", font = customtkinter.CTkFont(size=16, weight = "bold"))
        muda_assunto_label.pack(padx=10, pady=(20,10))
        muda_assunto_label2 = customtkinter.CTkLabel(master = janela_enviar_teste_vazio, text = teste_email)
        muda_assunto_label2.pack(padx=10, pady=0)

        botao_voltar = customtkinter.CTkButton(janela_enviar_teste_vazio, text = " Voltar ", corner_radius=0, command = janela_enviar_teste_vazio.destroy)
        botao_voltar.pack(padx=10, pady=10)

# B O T Ã O   M U D A R   M E N S A G E M 
def clique_mudar_email():
    # janela
    janela_mudar_email = customtkinter.CTkToplevel(janela)
    janela_mudar_email.title("Projeto e-mails - Mudar E-mail")
    janela_mudar_email.geometry("650x490+500+50")

    frame_janela = customtkinter.CTkFrame(janela_mudar_email, width=600, height=450)
    frame_janela.place(x=20,y=20)


#o assunto não precisa de formatação etão pode ficar só como entry
    def botao_guardar_assunto():
        assunto_entry1 = str(assunto_entry.get())
        with open("assunto.txt","w", encoding = "utf-8") as guardar_assunto_arq:       
            guardar_assunto = guardar_assunto_arq.write(assunto_entry1)
        
#isto está a pegar a informação que colocou na entry do corpo de texto        
    def botao_guardar_corpo_texto():
        corpo_texto_entry1 = str(inserir_corpo_texto.get())
        with open("corpo_email.txt","w", encoding = "utf-8") as guardar_corpo_texto_arq:       
            guardar_corpo_texto = guardar_corpo_texto_arq.write(corpo_texto_entry1)

#então como é que está a receber a informação?
    # BOTÃO OCNSTRUTOR DE CORPO DE TEXTO
    def botao_construtor_HTML():
        def botao_construtor_HTML1():
            #ao carregar abrir uma determinada pagina web
            webbrowser.open_new_tab("http://127.0.0.1:5000")
            app.run()
            pass
        
        #para abrir a nova pagina ela estava me a criar sempre um novo botão, então eu não coloquei as coordenadas da pagina
        #assim ele cria mas não aparece
        btn_corpo_texto1 = customtkinter.CTkButton(master=coluna2, width = 1, height= 1,text = "", command = threading.Thread(target=botao_construtor_HTML1).start())
        #btn_corpo_texto1.place(x=30,y=110)
        pass
    
    # ANEXAR ANEXOS
    def clique_adicionar_anexos():
        janela_adicionar_anexos = customtkinter.CTkToplevel(janela_mudar_email)
        janela_adicionar_anexos.title("Projeto e-mails - Adicionar anexos")
        janela_adicionar_anexos.geometry("680x550+500+50")

        def clique_anexo_full_path():
            adicionar_anexos_entry1a = adicionar_anexos_entry1.get()
            with open("anexo1.txt","w", encoding = "utf-8") as anexo1_arq:       
                anexo1 = anexo1_arq.write(adicionar_anexos_entry1a)
            pass
        def clique_anexo_full_path2():
            adicionar_anexos_entry2a = adicionar_anexos_entry2.get()
            with open("anexo2.txt","w", encoding = "utf-8") as anexo2_arq:       
                anexo2 = anexo2_arq.write(adicionar_anexos_entry2a)
            pass
        def clique_anexo_full_path3():
            adicionar_anexos_entry3a = adicionar_anexos_entry3.get()
            with open("anexo3.txt","w", encoding = "utf-8") as anexo3_arq:       
                anexo3 = anexo3_arq.write(adicionar_anexos_entry3a)
            pass
        def clique_anexo_full_path4():
            adicionar_anexos_entry4a = adicionar_anexos_entry4.get()
            with open("anexo4.txt","w", encoding = "utf-8") as anexo4_arq:       
                anexo4 = anexo4_arq.write(adicionar_anexos_entry4a)
            pass
        def clique_anexo_full_path5():
            adicionar_anexos_entry5a = adicionar_anexos_entry5.get()
            with open("anexo5.txt","w", encoding = "utf-8") as anexo5_arq:       
                anexo5 = anexo5_arq.write(adicionar_anexos_entry5a)
            pass

        #criar espaço para ter as entrys ao lado dos botões
        adicionar_anexos_coluna1 = customtkinter.CTkFrame(janela_adicionar_anexos, width=440, height=430)
        adicionar_anexos_coluna1.place(x=20,y=20)   

        # ler o caminho onde está o anexo
        with open("anexo1.txt", "r", encoding = "utf-8") as adicionar_anexos1_arq:
            adicionar_anexos1 = adicionar_anexos1_arq.read()
        adicionar_anexos_entry1 = customtkinter.CTkEntry(adicionar_anexos_coluna1, width=400, placeholder_text=adicionar_anexos1)
        adicionar_anexos_entry1.pack(padx=10, pady=10)
        with open("anexo2.txt", "r", encoding = "utf-8") as adicionar_anexos2_arq:
            adicionar_anexos2 = adicionar_anexos2_arq.read()
        adicionar_anexos_entry2 = customtkinter.CTkEntry(adicionar_anexos_coluna1, width=400, placeholder_text=adicionar_anexos2)
        adicionar_anexos_entry2.pack(padx=10, pady=10)
        with open("anexo3.txt", "r", encoding = "utf-8") as adicionar_anexos3_arq:
            adicionar_anexos3 = adicionar_anexos3_arq.read()
        adicionar_anexos_entry3 = customtkinter.CTkEntry(adicionar_anexos_coluna1, width=400, placeholder_text=adicionar_anexos3)
        adicionar_anexos_entry3.pack(padx=10, pady=10)
        with open("anexo4.txt", "r", encoding = "utf-8") as adicionar_anexos4_arq:
            adicionar_anexos4 = adicionar_anexos4_arq.read()
        adicionar_anexos_entry4 = customtkinter.CTkEntry(adicionar_anexos_coluna1, width=400, placeholder_text=adicionar_anexos4)
        adicionar_anexos_entry4.pack(padx=10, pady=10)
        with open("anexo5.txt", "r", encoding = "utf-8") as adicionar_anexos5_arq:
            adicionar_anexos5 = adicionar_anexos5_arq.read()
        adicionar_anexos_entry5 = customtkinter.CTkEntry(adicionar_anexos_coluna1, width=400, placeholder_text=adicionar_anexos5)
        adicionar_anexos_entry5.pack(padx=10, pady=10)

        adicionar_anexos_info_label1= customtkinter.CTkLabel(adicionar_anexos_coluna1, text="Para adicionar o ficheiro:", justify="left")
        adicionar_anexos_info_label1.pack(padx=10, pady=5)
        adicionar_anexos_info_label2= customtkinter.CTkLabel(adicionar_anexos_coluna1, text=" 1) Botão direito em cima do ficheiro a anexar;\n 2) Propriedades;\n 3) Copiar a localização para a barra de anexos, colocar uma barra invertida,\nnome do ficheiro, ponto, extensão/tipo do ficheiro \n (pdf, xlsx, docx, png...) ", justify="left")
        adicionar_anexos_info_label2.pack(padx=10, pady=(0,10))
        adicionar_anexos_info_label3= customtkinter.CTkLabel(adicionar_anexos_coluna1, text="Exemplo (só pode haver espaços no Nome do Ficheiro):", justify="left")
        adicionar_anexos_info_label3.pack(padx=10, pady=0)
        adicionar_anexos_info_label3= customtkinter.CTkLabel(adicionar_anexos_coluna1, text=r"C:\Users\gaspar\Desktop\Ficha de Inscrição.pdf", justify="left")
        adicionar_anexos_info_label3.pack(padx=10, pady=0)
        adicionar_anexos_info_label5= customtkinter.CTkLabel(adicionar_anexos_coluna1, text="Para não enviar um anexo basta colocar um espço e carregar no anexar. \nNOTA: Se tiver 2 espaços ou um caracter vai dar erro no envio. Tal como\nse colocar um caminho inválido ou com extensão incorreta (ter em atenção\nque um documento word pode ter como extensão .doc ou .docx) ", justify="left")
        adicionar_anexos_info_label5.pack(padx=10, pady=10)

        #criar espaço para ter os botões ao lado das entrys
        adicionar_anexos_coluna2 = customtkinter.CTkFrame(janela_adicionar_anexos, width=180, height=430)
        adicionar_anexos_coluna2.place(x=480,y=20)  
        adicionar_anexos_label_coluna2= customtkinter.CTkLabel(adicionar_anexos_coluna2, text=" ")
        adicionar_anexos_label_coluna2.place(x=20,y=20)
        botao_adicionar_anexos_entry1 = customtkinter.CTkButton(adicionar_anexos_coluna2, text = " Anexo 1 ", corner_radius=8, command = clique_anexo_full_path)
        botao_adicionar_anexos_entry1.pack(padx=10, pady=10)
        botao_adicionar_anexos_entry2 = customtkinter.CTkButton(adicionar_anexos_coluna2, text = " Anexo 2 ", corner_radius=8, command = clique_anexo_full_path2)
        botao_adicionar_anexos_entry2.pack(padx=10, pady=10)
        botao_adicionar_anexos_entry3 = customtkinter.CTkButton(adicionar_anexos_coluna2, text = " Anexo 3 ", corner_radius=8, command = clique_anexo_full_path3)
        botao_adicionar_anexos_entry3.pack(padx=10, pady=10)
        botao_adicionar_anexos_entry4 = customtkinter.CTkButton(adicionar_anexos_coluna2, text = " Anexo 4 ", corner_radius=8, command = clique_anexo_full_path4)
        botao_adicionar_anexos_entry4.pack(padx=10, pady=10)
        botao_adicionar_anexos_entry5 = customtkinter.CTkButton(adicionar_anexos_coluna2, text = " Anexo 5 ", corner_radius=8, command = clique_anexo_full_path5)
        botao_adicionar_anexos_entry5.pack(padx=10, pady=10)

        botao_voltar = customtkinter.CTkButton(adicionar_anexos_coluna2, text = " Voltar ", corner_radius=8, command = janela_adicionar_anexos.destroy)
        botao_voltar.pack(padx=10, pady=30)
        pass

    coluna2 = customtkinter.CTkFrame(frame_janela, width=400, height=400)
    coluna2.place(x=20,y=50)
    assunto_entry = customtkinter.CTkEntry(coluna2, width=300, placeholder_text= " Novo assunto ")
    assunto_entry.pack(padx=10, pady=10)
    inserir_corpo_texto = customtkinter.CTkEntry(coluna2, width=300, placeholder_text="Novo corpo de texto")
    inserir_corpo_texto.pack(padx=10, pady=10)
    btn_corpo_texto1 = customtkinter.CTkButton(coluna2, width=250, text = " Construtor do email ", corner_radius=8, command = botao_construtor_HTML)
    btn_corpo_texto1.pack(padx=10, pady=10)
    botao_adicionar_anexos = customtkinter.CTkButton(coluna2, width=250, text = " Adicionar Anexos ", corner_radius=8, command = clique_adicionar_anexos)
    botao_adicionar_anexos.pack(padx=10, pady=10)


    coluna3 = customtkinter.CTkFrame(frame_janela, width=250, height=400)
    coluna3.place(x=360,y=50)
    btn_assunto = customtkinter.CTkButton(coluna3, width=200, text = " Adicionar Assunto ", corner_radius=8, command = botao_guardar_assunto)
    btn_assunto.pack(padx=10, pady=10)
    btn_corpo_texto = customtkinter.CTkButton(coluna3, width=200, text = " Adicionar corpo de texto ", corner_radius=8, command = botao_guardar_corpo_texto)
    btn_corpo_texto.pack(padx=10, pady=10)
    

    botao_voltar = customtkinter.CTkButton(coluna3, text = " Voltar ", corner_radius=8, command = janela_mudar_email.destroy)
    botao_voltar.pack(padx=10, pady=(40,10))

# B O T Ã O   M O S T R A R   E M A I L 
def clique_mostrar_emails():
#criar nova tela
    janela_mostrar_email = customtkinter.CTkToplevel(janela)
    janela_mostrar_email.title("Projeto e-mails")
    janela_mostrar_email.geometry("400x650+500+10")
#caixa de texto    
    muda_assunto_label = customtkinter.CTkLabel(master = janela_mostrar_email, text = "E-mails restantes !", font = customtkinter.CTkFont(size=16, weight = "bold"))
    muda_assunto_label.pack(padx=10, pady=(20,10))
    
#mostrar excel
    importar_mostrar_excel = pd.read_excel("dados_emails.xlsx")
    importar_mostrar_emails = importar_mostrar_excel[['emails']]

    tamanho_lista = len(importar_mostrar_excel['emails'])

    exibir_emails = customtkinter.CTkScrollableFrame(janela_mostrar_email, width=300, height=430)
    exibir_emails.pack(padx=10, pady=10)
    #mostrar só os 60 primeiros porque se pedir 61 vai aparecer umas reticencias no meio
    scrolabletext= customtkinter.CTkLabel(exibir_emails, text = importar_mostrar_emails[['emails']].head(60))
    scrolabletext.pack(padx=10, pady=10)

    tamanho_lista_label1= customtkinter.CTkLabel(janela_mostrar_email, text = "E-mails por enviar:")
    tamanho_lista_label1.pack(padx=0, pady=0)
    tamanho_lista_label2= customtkinter.CTkLabel(janela_mostrar_email, text = tamanho_lista)
    tamanho_lista_label2.pack(padx=0, pady=0)

    # Botão HELP
    def mostrar_emails_informacoes():
        janela_mostrar_emails_informacoes = customtkinter.CTkToplevel(janela)
        janela_mostrar_emails_informacoes.title("Projeto e-mails")
        janela_mostrar_emails_informacoes.geometry("400x650+900+10")

        emails_invalidos_informacoes_label = customtkinter.CTkLabel(janela_mostrar_emails_informacoes, text = "Se na lista aparecer \n\nemails\nxx  -----\n\n ou\n\nEmpty DataFrame \nColumns:[] \nIndex: []\n\n significa que a lista acabou")
        emails_invalidos_informacoes_label.pack(padx=10, pady=10)
        
        botao_voltar = customtkinter.CTkButton(janela_mostrar_emails_informacoes, text = " Voltar ", corner_radius=8, command = janela_mostrar_emails_informacoes.destroy)
        botao_voltar.pack(padx=10, pady=10)

        pass
    botao_mostrar_emails_informacoes = customtkinter.CTkButton(janela_mostrar_email, text = "Help",width=20, corner_radius=20, command = mostrar_emails_informacoes)
    botao_mostrar_emails_informacoes.place(x=10,y=10)

    botao_voltar = customtkinter.CTkButton(janela_mostrar_email, text = " Voltar ", corner_radius=8, command = janela_mostrar_email.destroy)
    botao_voltar.pack(padx=10, pady=10)

# B O T Ã O   O U T R A S   O P Ç Õ E S 
def clique_outras_opcoes():
    #Criação da janela
    janela_outras_opcoes = customtkinter.CTkToplevel(janela)
    janela_outras_opcoes.title("Projeto e-mails - Outras opções")
    janela_outras_opcoes.geometry("400x650+500+10")
   
    outras_opcoes_label = customtkinter.CTkLabel(janela_outras_opcoes, text = "Outras opções:", font = customtkinter.CTkFont(size=16, weight = "bold"))
    outras_opcoes_label.pack(padx=10, pady=(20,10))
    
    outras_opcoes_frame= customtkinter.CTkFrame(janela_outras_opcoes, width=350, height=550)
    outras_opcoes_frame.pack(padx=10, pady=10)

    # NUMERO DE EMAILS
    def clique_alterar_n_emails():
        try:
            #vou guardar no txt
            with open("numero_envios.txt","w", encoding = "utf-8") as n_envios_arq:       
                n_envios_arq.write(n_emails.get())
            #vou buscar
            with open("numero_envios.txt","r", encoding = "utf-8") as n_envios_arq:       
                numero_envios = n_envios_arq.read()
            #vou verificar
            #caso correto vou guardar
            numero_envios1 = int(numero_envios)
            if numero_envios1 > 0 : # o numero de envios deve ser um numero maior do que zero
                #janela de confirmação da alteração da quantidare - alteração feita com sucesso
                janela_alterar_n_emails = customtkinter.CTkToplevel(janela_outras_opcoes)
                janela_alterar_n_emails.title("Projeto e-mails")
                janela_alterar_n_emails.geometry("300x150+900+100")

                alterar_n_emails_label = customtkinter.CTkLabel(janela_alterar_n_emails, text = "Alteração feita com sucesso", font = customtkinter.CTkFont(size=16, weight = "bold"))
                alterar_n_emails_label.pack(padx=10, pady=(20,10))

                botao_voltar = customtkinter.CTkButton(janela_alterar_n_emails, text = " Voltar ", corner_radius=8, command = janela_alterar_n_emails.destroy)
                botao_voltar.pack(padx=10, pady=10)
            pass
        except ValueError:
            #janela de erro quando o valor guardado não é um numero
            janela_lista_original_nao_encontrada = customtkinter.CTkToplevel(janela, fg_color="red")
            janela_lista_original_nao_encontrada.title("Projeto e-mails - Erro")
            janela_lista_original_nao_encontrada.geometry("300x150+900+100") 

            lista_original_nao_encontrada_label = customtkinter.CTkLabel(janela_lista_original_nao_encontrada, text = "O que inseriu não é numero", font = customtkinter.CTkFont(size=16, weight = "bold"))
            lista_original_nao_encontrada_label.pack(padx=10, pady=(20,10))
            botao_voltar = customtkinter.CTkButton(janela_lista_original_nao_encontrada, text = " Voltar ", corner_radius=8, command = janela_lista_original_nao_encontrada.destroy)
            botao_voltar.pack(padx=10, pady=10)
            pass

    # ALTERAR EMAIL DE TESTES       
    def clique_alterar_email():
        #ir buscar o que foi escrito na entry
        with open("email_teste.txt","w", encoding = "utf-8") as email_teste_arq:       
            email_testar = email_teste_arq.write(alterar_email_teste_entry.get())
        #janela de confirmação
        janela_alterar_email_teste = customtkinter.CTkToplevel(janela_outras_opcoes)
        janela_alterar_email_teste.title("Projeto e-mails ")
        janela_alterar_email_teste.geometry("300x150+900+100")

        alterar_n_emails_label = customtkinter.CTkLabel(janela_alterar_email_teste, text = "E-mail teste alterado", font = customtkinter.CTkFont(size=16, weight = "bold"))
        alterar_n_emails_label.pack(padx=10, pady=(20,10))

        botao_voltar = customtkinter.CTkButton(janela_alterar_email_teste, text = " Voltar ", corner_radius=8, command = janela_alterar_email_teste.destroy)
        botao_voltar.pack(padx=10, pady=10)
        pass

    # IMPORTAÇÃO DE NOVA LISTA DE EMAILS
    def clique_nova_lista():
        try:
            #abrir excel
            tabela= pd.read_excel("dados_emails - completo.xlsx")
            emails = tabela[['emails']] #puxar colunas       
            #guardar excel 
            file_name = 'dados_emails.xlsx'
            emails.to_excel(file_name)    
                
            janela_nova_lista = customtkinter.CTkToplevel(janela_outras_opcoes)
            janela_nova_lista.title("Projeto e-mails - E-mail teste alterado")
            janela_nova_lista.geometry("300x150+900+100")

            alterar_n_emails_label = customtkinter.CTkLabel(janela_nova_lista, text = "Nova lista importada", font = customtkinter.CTkFont(size=16, weight = "bold"))
            alterar_n_emails_label.pack(padx=10, pady=(20,10))

            botao_voltar = customtkinter.CTkButton(janela_nova_lista, text = " Voltar ", corner_radius=8, command = janela_nova_lista.destroy)
            botao_voltar.pack(padx=10, pady=10)
        except FileNotFoundError:
            #janela de erro caso não exista o documento na pasta com o nome "dados_emails - completo.xlsx"
            janela_lista_original_nao_encontrada = customtkinter.CTkToplevel(janela, fg_color="red")
            janela_lista_original_nao_encontrada.title("Projeto e-mails - Erro")
            janela_lista_original_nao_encontrada.geometry("300x150+900+100") 

            lista_original_nao_encontrada_label = customtkinter.CTkLabel(janela_lista_original_nao_encontrada, text = "Nova lista não \nencontrada", font = customtkinter.CTkFont(size=16, weight = "bold"))
            lista_original_nao_encontrada_label.pack(padx=10, pady=(20,10))
            botao_voltar = customtkinter.CTkButton(janela_lista_original_nao_encontrada, text = " Voltar ", corner_radius=8, command = janela_lista_original_nao_encontrada.destroy)
            botao_voltar.pack(padx=10, pady=10)
            pass

    # SWITCH PARA LIGAR E DESLIGAR A REPOSIÇÃO AUTOMÁTICA
    def switch_reposicao_automatica():
        #quando clicado é guardado
        #print("valor switch reposição automática ->", switch_var.get())
        with open("reposicao_automatica.txt","w", encoding = "utf-8") as reposicao_automatica_arq:       
            reposicao_automatica = reposicao_automatica_arq.write(switch_var.get())
        pass
    
    # VERIFICAR EMAILS REJEITADOS PELO OUTLOOK
    def clique_rejeitados_outlook():

        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        #pasta onde procurar os emails - a caixa de entrada/inbox foi definida como pasta 6 (pela documentação)
        inbox = outlook.GetDefaultFolder(6)
        mensagens_inbox = inbox.Items
        lista_rejeitados=[]

        #vou procurar determinado texto nas mensagens da inbox
        for msg in mensagens_inbox:
            if 'This message was created automatically by mail delivery software' in str(msg.Body):
                #procurar a expressao que coincida com o email
                procurar_email = re.search(r"[A-Za-z0-9_%+-.]+"r"@[A-Za-z0-9.-]+"r"\.[A-Za-z]{2,5}",msg.Body)
                lista_rejeitados.append(procurar_email.group()) #adicionar à lista

        #passar para dataframe para eliminar duplicados
        email_rejeitado_df = pd.DataFrame({'rejeitados': lista_rejeitados})
        email_rejeitado_df.drop_duplicates

        file_email_rejeitado = 'emails_rejeitados_pelo_outlook.xlsx'
        email_rejeitado_df.to_excel(file_email_rejeitado)
    
        def clique_eliminar_rejeitados_outlook():
            tabela_email_rejeitado_eliminar_original= pd.read_excel("dados_emails - completo.xlsx")
            email_rejeitado_eliminar_original = tabela_email_rejeitado_eliminar_original['emails'] #puxar colunas
            
            tabela_email_rejeitado_eliminar1= pd.read_excel("emails_rejeitados_pelo_outlook.xlsx")
            email_rejeitado_eliminar1 = tabela_email_rejeitado_eliminar1['rejeitados'] 

            lista_email_rejeitado_eliminar_original = email_rejeitado_eliminar_original.tolist()
            dicionario_rejeitados = {'emails':lista_email_rejeitado_eliminar_original,'e-mails':lista_email_rejeitado_eliminar_original}
            email_rejeitado_eliminar_original_df = pd.DataFrame( dicionario_rejeitados,columns=['emails'], index= lista_email_rejeitado_eliminar_original)
            lista_nova=[]
            for rejeitado_eliminar3 in email_rejeitado_eliminar_original_df['emails']:   
                for rejeitado in email_rejeitado_eliminar1:          
                    if rejeitado == rejeitado_eliminar3:
                        lista_nova.append(rejeitado)
                    else: a=1
            #print("listanova:",lista_nova)
            email_rejeitado_eliminar_original_df.drop_duplicates
            for rejeitado in lista_nova: 
                email_rejeitado_eliminar_original_df = email_rejeitado_eliminar_original_df.drop(rejeitado)
            #print("segundo print",email_rejeitado_eliminar_original_df)

            file_name_emails_sem_rejeitados = 'dados_emails - completo.xlsx'
            email_rejeitado_eliminar_original_df.to_excel(file_name_emails_sem_rejeitados)

            #janela de confirmação
            janela_email_rejeitado_confirmacao = customtkinter.CTkToplevel(janela)
            janela_email_rejeitado_confirmacao.title("Projeto e-mails - E-mail teste alterado")
            janela_email_rejeitado_confirmacao.geometry("300x150+400+100")

            alterar_n_emails_label = customtkinter.CTkLabel(janela_email_rejeitado_confirmacao, text = "Emails rejeitados eliminados \n da lista original", font = customtkinter.CTkFont(size=16, weight = "bold"))
            alterar_n_emails_label.pack(padx=10, pady=(20,10))

            botao_voltar = customtkinter.CTkButton(janela_email_rejeitado_confirmacao, text = " Voltar ", corner_radius=8, command = janela_email_rejeitado_confirmacao.destroy)
            botao_voltar.pack(padx=10, pady=10)

            pass
        

        #Janela para visualizar rejeitados
        janela_email_rejeitado = customtkinter.CTkToplevel(janela)
        janela_email_rejeitado.title("Projeto e-mails")
        janela_email_rejeitado.geometry("400x650+900+10")

        email_rejeitado_label = customtkinter.CTkLabel(master = janela_email_rejeitado, text = "E-mails rejeitados pelo\n Outlook", font = customtkinter.CTkFont(size=16, weight = "bold"))
        email_rejeitado_label.pack(padx=10, pady=(20,10))

        tamanho_lista_email_rejeitado = len(email_rejeitado_df)

        exibir_email_rejeitado = customtkinter.CTkScrollableFrame(janela_email_rejeitado, width=300, height=400)
        exibir_email_rejeitado.pack(padx=10, pady=10)

        tabela_email_rejeitado= pd.read_excel("emails_rejeitados_pelo_outlook.xlsx")
        apresentar_email_rejeitado = tabela_email_rejeitado[['rejeitados']] #puxar colunas
        email_rejeitado_scrolabletext1= customtkinter.CTkLabel(exibir_email_rejeitado,width=600, text = apresentar_email_rejeitado.head(60))
        email_rejeitado_scrolabletext1.pack(padx=10, pady=10)

        tamanho_lista_emails_invalidos_label1= customtkinter.CTkLabel(janela_email_rejeitado, text = f"Número de emails rejeitados pelo Outlook: {tamanho_lista_email_rejeitado}")
        tamanho_lista_emails_invalidos_label1.pack(padx=0, pady=0)

        botao_eliminar_email_rejeitado = customtkinter.CTkButton(janela_email_rejeitado, text = " Eliminar emails rejeitados da lista completa ", corner_radius=8, command = clique_eliminar_rejeitados_outlook)
        botao_eliminar_email_rejeitado.pack(padx=10, pady=10)

        botao_voltar = customtkinter.CTkButton(janela_email_rejeitado, text = " Voltar ", corner_radius=8, command = janela_email_rejeitado.destroy)
        botao_voltar.pack(padx=10, pady=10)
        pass

    # VERIFICAR SE OS EMAILS ESTÃO CORRETAMENTE PREENCHIDOS
    def clique_verificar_emails():
        #importar emails
        tabela_emails_a_verificar= pd.read_excel("dados_emails.xlsx")
        emails_a_verificar = tabela_emails_a_verificar[['emails']] #puxar colunas

        #importar excel dos dominios que foram acrescentados
        tabela_dominios_a_verificar= pd.read_excel("dominios.xlsx")
        dominios_a_verificar = tabela_dominios_a_verificar['dominios'] #puxar colunas
        dominios_a_verificar_lista = dominios_a_verificar.tolist()

        #deletar registos duplicados
        emails_a_verificar=emails_a_verificar.drop_duplicates()
        lista_bons_dominios=[]
        emails_invalidos = []
        for email_a_verificar in emails_a_verificar['emails']:
            email_a_verificar=email_a_verificar.lower()
            # verificar dominios mais comuns
            p1 = r'@gmail.com$'
            p2 = r'@hotmail.com$'
            p3 = r'@outlook.com$'
            p4 = r'@outlook.pt$'
            p5 = r'@yahoo.com$'
            p6 = r'@sapo.pt$'
            p7 = r'@topdata.pt$'
            
            p9 = r'\s' #verificar espaços
            p10= r'"' #verificar aspas
            p11= r',' #verificar virgulas
            p12= r'[A-Za-z0-9_.+-]+@[A-Za-z0-9_.+-]+@[A-Za-z0-9_.+-]' #verificar tem 2 @ intercalados
            p13=r'@{2,}'#verificar tem 2 @ seguidos            

            # V E R I F I C A R   S E   T E M :
            if re.findall(p9,email_a_verificar):
                #print("TEM ESPAÇOS: |",email_a_verificar,"|")
                emails_invalidos.append(email_a_verificar)
            elif re.findall(p10,email_a_verificar):
                #print("TEM ASPAS: |",email_a_verificar,"|")
                emails_invalidos.append(email_a_verificar)
            elif re.findall(p11,email_a_verificar):
                #print("TEM vírgula: |",email_a_verificar,"|")
                emails_invalidos.append(email_a_verificar)
            elif re.findall(p12,email_a_verificar):
                #print("dois arrobas: |", email_a_verificar, "|")
                emails_invalidos.append(email_a_verificar) 
            elif re.findall(p13,email_a_verificar):
                #print("dois arrobas: |", email_a_verificar, "|")
                emails_invalidos.append(email_a_verificar) 

            elif re.findall(p1,email_a_verificar): a=1
            elif re.findall(p2,email_a_verificar): a=1
            elif re.findall(p3,email_a_verificar): a=1
            elif re.findall(p4,email_a_verificar): a=1
            elif re.findall(p5,email_a_verificar): a=1
            elif re.findall(p6,email_a_verificar): a=1
            elif re.findall(p7,email_a_verificar): a=1
            
            #se tiver um problema é adicionado aos invalidos
            else: 
                for dominio in dominios_a_verificar_lista: 
                    if re.findall(dominio,email_a_verificar):
                        lista_bons_dominios.append(email_a_verificar)
                        continue
                    else:
                        emails_invalidos.append(email_a_verificar)
                        
        dicionarios_invalidos = {'emails':emails_invalidos,'indice':emails_invalidos }
        emails_invalidos_df = pd.DataFrame( dicionarios_invalidos,columns=['emails'], index= emails_invalidos)
        emails_invalidos_df = emails_invalidos_df.drop_duplicates()
                    
        for dominios in lista_bons_dominios: 
            emails_invalidos_df = emails_invalidos_df.drop(dominios)
        
        #guardar a informação num excel - substiui o excel se ele já existir
        file_name_emails_invalidos= 'emails_invalidos.xlsx'
        emails_invalidos_df.to_excel(file_name_emails_invalidos)

        #janela para mostrar os emails que podem ter erros
        janela_emails_invalidos = customtkinter.CTkToplevel(janela)
        janela_emails_invalidos.title("Projeto e-mails")
        janela_emails_invalidos.geometry("400x650+900+10")
        
        def mostrar_dominios():
            janela_mostrar_dominios = customtkinter.CTkToplevel(janela)
            janela_mostrar_dominios.title("Projeto e-mails")
            janela_mostrar_dominios.geometry("400x650+200+10")
        #caixa de texto    
            mostrar_dominios_label = customtkinter.CTkLabel(master = janela_mostrar_dominios, text = " Dominios adicionados", font = customtkinter.CTkFont(size=16, weight = "bold"))
            mostrar_dominios_label.pack(padx=10, pady=(20,10))
            
        #mostrar excel
            mostrar_dominios_excel = pd.read_excel("dominios.xlsx")
            mostrar_dominios_emails = mostrar_dominios_excel[['dominios']]

            mostrar_dominios_tamanho_lista = len(mostrar_dominios_emails['dominios'])

            mostrar_dominios_frame = customtkinter.CTkScrollableFrame(janela_mostrar_dominios, width=300, height=360)
            mostrar_dominios_frame.pack(padx=10, pady=10)
            mostrar_dominios_scrolabletext= customtkinter.CTkLabel(mostrar_dominios_frame, text = mostrar_dominios_emails[['dominios']].head(60))
            mostrar_dominios_scrolabletext.pack(padx=10, pady=10)

            mostrar_dominios_tamanho_lista_label1= customtkinter.CTkLabel(janela_mostrar_dominios, text = f"Número de domínios adicionados: {mostrar_dominios_tamanho_lista}")
            mostrar_dominios_tamanho_lista_label1.pack(padx=0, pady=0)

            def clique_eliminar_mostrar_dominios():
                tabela_dominios_eliminar= pd.read_excel("dominios.xlsx")
                dominios_eliminar = tabela_dominios_eliminar[['dominios']] #puxar colunas

                indice_eliminar = int(eliminar_mostrar_dominios_entry.get())
                dominios_eliminar = dominios_eliminar.drop([indice_eliminar],axis=0)

                file_name_dominios_eliminar = 'dominios.xlsx'
                dominios_eliminar.to_excel(file_name_dominios_eliminar)

                #janela de confirmação
                janela_dominios_eliminar_confirmacao = customtkinter.CTkToplevel(janela)
                janela_dominios_eliminar_confirmacao.title("Projeto e-mails ")
                janela_dominios_eliminar_confirmacao.geometry("300x150+400+100")

                alterar_n_emails_label = customtkinter.CTkLabel(janela_dominios_eliminar_confirmacao, text = "Dominio eliminados", font = customtkinter.CTkFont(size=16, weight = "bold"))
                alterar_n_emails_label.pack(padx=10, pady=(20,10))

                botao_voltar = customtkinter.CTkButton(janela_dominios_eliminar_confirmacao, text = " Voltar ", corner_radius=8, command = janela_dominios_eliminar_confirmacao.destroy)
                botao_voltar.pack(padx=10, pady=10)
                

            #label informar que é para colocar o numero do indice que se encontra antes do email
            eliminar_mostrar_dominios_label = customtkinter.CTkLabel(janela_mostrar_dominios, text = "Indique o numero do indice do domínio que pretende eliminar:")
            eliminar_mostrar_dominios_label.pack(padx=0, pady=0)
            #entry indice do dominio a eliminar
            eliminar_mostrar_dominios_entry = customtkinter.CTkEntry(janela_mostrar_dominios, width=40, placeholder_text= "Numero")
            eliminar_mostrar_dominios_entry.pack(padx=0, pady=0)
            #botão de fazer evento
            eliminar_mostrar_dominios_botao = customtkinter.CTkButton(janela_mostrar_dominios, width=80, text = " Eliminar ", corner_radius=8, command = clique_eliminar_mostrar_dominios)
            eliminar_mostrar_dominios_botao.pack(padx=10, pady=10)

            botao_voltar = customtkinter.CTkButton(janela_mostrar_dominios, text = " Voltar ", corner_radius=8, command = janela_mostrar_dominios.destroy)
            botao_voltar.pack(padx=10, pady=10)
            pass

        emails_invalidos_label = customtkinter.CTkLabel(master = janela_emails_invalidos, text = "E-mails possivelmente \ninválidos", font = customtkinter.CTkFont(size=16, weight = "bold"))
        emails_invalidos_label.pack(padx=10, pady=(20,10))

        tamanho_lista_emails_invalidos = len(emails_invalidos_df)

        exibir_emails_invalidos = customtkinter.CTkScrollableFrame(janela_emails_invalidos, width=300, height=360)
        exibir_emails_invalidos.pack(padx=10, pady=10)

        tabela_apresentar_invalidos= pd.read_excel("emails_invalidos.xlsx")
        apresentar_invalidos = tabela_apresentar_invalidos[['emails']] #puxar colunas
        scrolabletext1= customtkinter.CTkLabel(exibir_emails_invalidos,width=600, text = apresentar_invalidos.head(60))
        scrolabletext1.pack(padx=10, pady=10)

        tamanho_lista_emails_invalidos_label1= customtkinter.CTkLabel(janela_emails_invalidos, text = f"Número de emails possivelmente inválidos: {tamanho_lista_emails_invalidos}")
        tamanho_lista_emails_invalidos_label1.pack(padx=0, pady=0)
        
        # Botão HELP
        def emails_invalidos_informacoes():
            janela_emails_invalidos_informacoes = customtkinter.CTkToplevel(janela)
            janela_emails_invalidos_informacoes.title("Projeto e-mails")
            janela_emails_invalidos_informacoes.geometry("400x650+200+10")
            
            emails_invalidos_informacoes_label1 = customtkinter.CTkLabel(janela_emails_invalidos_informacoes, text = "Se na lista aparecer \n\nEmpty DataFrame \nColumns:[] \nIndex: []\n\n significa que não foram identificados emails para reparar")
            emails_invalidos_informacoes_label1.pack(padx=10, pady=10)
            separador_de_informação = customtkinter.CTkLabel(janela_emails_invalidos_informacoes, text = " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ")
            separador_de_informação.pack(padx=10, pady=5)
            emails_invalidos_informacoes_label2 = customtkinter.CTkLabel(janela_emails_invalidos_informacoes, text = "Se aparecer: \n\n emails indice \n ----- ----- ----- \n \n também significa que o excel está vazio")
            emails_invalidos_informacoes_label2.pack(padx=10, pady=10)
            separador_de_informação2 = customtkinter.CTkLabel(janela_emails_invalidos_informacoes, text = " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ")
            separador_de_informação2.pack(padx=10, pady=5)
            emails_invalidos_informacoes_label3 = customtkinter.CTkLabel(janela_emails_invalidos_informacoes, text = "A lista de domínios deve ter 3 ou mais domínios \n para não dar problemas")
            emails_invalidos_informacoes_label3.pack(padx=10, pady=10)

            botao_voltar = customtkinter.CTkButton(janela_emails_invalidos_informacoes, text = " Voltar ", corner_radius=8, command = janela_emails_invalidos_informacoes.destroy)
            botao_voltar.pack(padx=10, pady=10)
            pass
        
        #def clique_botao_emails_invalidos():
        #        tabela_apresentar_invalidos_eliminar= pd.read_excel("emails_invalidos.xlsx")
        #        apresentar_invalidos_eliminar = tabela_apresentar_invalidos_eliminar[['emails']] #puxar colunas
        #    
        #        email_indice_eliminar = int(botao_emails_invalidos_entry.get())
        #        print(email_indice_eliminar)

        botao_emails_invalidos_informacoes = customtkinter.CTkButton(janela_emails_invalidos, text = "Help",width=20, corner_radius=8, command = emails_invalidos_informacoes)
        botao_emails_invalidos_informacoes.place(x=10,y=10)

        #label informar que é para colocar o numero do indice que se encontra antes do email
        #botao_emails_invalidos_label = customtkinter.CTkLabel(janela_emails_invalidos, text = "Indique o numero do indice do email que pretende eliminar:")
        #botao_emails_invalidos_label.pack(padx=0, pady=0)
        #entry indice do dominio a eliminar
        #botao_emails_invalidos_entry = customtkinter.CTkEntry(janela_emails_invalidos, width=40, placeholder_text= "Numero")
        #botao_emails_invalidos_entry.place(x=100, y=520)
        #botão de fazer evento
        #botao_emails_invalidos_botao = customtkinter.CTkButton(janela_emails_invalidos, width=80, text = " Eliminar ", corner_radius=8, command = clique_botao_emails_invalidos)
        #botao_emails_invalidos_botao.place(x=160, y=520)

        botao_mostrar_dominios = customtkinter.CTkButton(janela_emails_invalidos, text = " Mostrar domínios adicionados ", corner_radius=8, command = mostrar_dominios)
        botao_mostrar_dominios.pack(padx=10, pady=10)

        botao_voltar = customtkinter.CTkButton(janela_emails_invalidos, text = " Voltar ", corner_radius=8, command = janela_emails_invalidos.destroy)
        botao_voltar.pack(padx=10, pady=10)

        pass

    # ADICIONAR DOMINIOS À LISTA DE VERIFICAÇÃO
    def clique_adicionar_dominio():
        tabela_adicionar_dominio= pd.read_excel("dominios.xlsx")
        adicionar_dominio = tabela_adicionar_dominio['dominios'] #puxar colunas
        adicionar_dominio_list = adicionar_dominio.tolist() #criar lista
        
        dominio = adicionar_dominio_entry.get()

        regex_dominio = re.compile(r'@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
        if re.fullmatch(regex_dominio,dominio):
            adicionar_dominio_list.append(dominio)# adicionado à lista
            print(adicionar_dominio_list)
            dicionario = {'dominios':adicionar_dominio_list}
            adicionar_dominio_serie=pd.DataFrame(dicionario)

            novo_dominio_guardar_file = "dominios.xlsx"
            adicionar_dominio_serie.to_excel(novo_dominio_guardar_file)    
            adicionar_dominio_entry.delete(0,END)

            janela_adicionar_dominio = customtkinter.CTkToplevel(janela)
            janela_adicionar_dominio.title("Projeto e-mails ")
            janela_adicionar_dominio.geometry("300x150+900+100")

            adicionar_dominio_label = customtkinter.CTkLabel(master = janela_adicionar_dominio, text = "Dominio adicionado \ncom sucesso", font = customtkinter.CTkFont(size=16, weight = "bold"))
            adicionar_dominio_label.pack(padx=10, pady=(20,10))
            botao_voltar = customtkinter.CTkButton(janela_adicionar_dominio, text = " Voltar ", corner_radius=8, command = janela_adicionar_dominio.destroy)
            botao_voltar.pack(padx=10, pady=10)
        else:
            janela_erro_adicionar_dominio = customtkinter.CTkToplevel(janela, fg_color="red")
            janela_erro_adicionar_dominio.title("Projeto e-mails - Erro")
            janela_erro_adicionar_dominio.geometry("300x150+900+100")

            erro_adicionar_dominio_label = customtkinter.CTkLabel(master = janela_erro_adicionar_dominio, text = "Ocorreu um erro, \nverifique se começou por @ e \nse o dominio está bem escrito", font = customtkinter.CTkFont(size=16, weight = "bold"))
            erro_adicionar_dominio_label.pack(padx=10, pady=(20,10))
            botao_voltar = customtkinter.CTkButton(janela_erro_adicionar_dominio, text = " Voltar ", corner_radius=8, command = janela_erro_adicionar_dominio.destroy)
            botao_voltar.pack(padx=10, pady=10)
        pass

    # ALTERAR TEMPO ENTRE ENVIO DE EMAILS AUTOMÁTICOS
    def clique_emails_automaticos_tempo():
        try:
            #vou guardar no txt
            with open("emails_automaticos_tempo.txt","w", encoding = "utf-8") as emails_automaticos_tempo_arq:       
                emails_automaticos_tempo_arq.write(emails_automaticos_tempo_entry.get())
            #vou buscar
            with open("emails_automaticos_tempo.txt","r", encoding = "utf-8") as emails_automaticos_tempo_arq:       
                emails_automaticos_tempo = emails_automaticos_tempo_arq.read()
            #vou verificar
            #caso correto vou guardar
            emails_automaticos_tempo1 = int(emails_automaticos_tempo)
            if emails_automaticos_tempo1 > 0 :
                janela_emails_automaticos_tempo = customtkinter.CTkToplevel(janela_outras_opcoes)
                janela_emails_automaticos_tempo.title("Projeto e-mails")
                janela_emails_automaticos_tempo.geometry("300x150+900+100")

                alterar_n_emails_label = customtkinter.CTkLabel(janela_emails_automaticos_tempo, text = "Tempo Alterado", font = customtkinter.CTkFont(size=16, weight = "bold"))
                alterar_n_emails_label.pack(padx=10, pady=(20,10))

                botao_voltar = customtkinter.CTkButton(janela_emails_automaticos_tempo, text = " Voltar ", corner_radius=8, command = janela_emails_automaticos_tempo.destroy)
                botao_voltar.pack(padx=10, pady=10)
            pass
        except ValueError:
            #janela de erro devido ao caracter inserido ser inválido
            janela_emails_automaticos_tempo_erro_valor = customtkinter.CTkToplevel(janela, fg_color="red")
            janela_emails_automaticos_tempo_erro_valor.title("Projeto e-mails - Erro")
            janela_emails_automaticos_tempo_erro_valor.geometry("300x150+900+100") 

            with open("emails_automaticos_tempo.txt","w", encoding = "utf-8") as emails_automaticos_tempo_arq:       
                emails_automaticos_tempo_arq.write("60")
            lista_original_nao_encontrada_label1 = customtkinter.CTkLabel(janela_emails_automaticos_tempo_erro_valor, text = "O que inseriu não é numero", font = customtkinter.CTkFont(size=16, weight = "bold"))
            lista_original_nao_encontrada_label1.pack(padx=10, pady=(20,0))
            lista_original_nao_encontrada_label2 = customtkinter.CTkLabel(janela_emails_automaticos_tempo_erro_valor, text = "Vai ser reposto o tempo original de 1 hora")
            lista_original_nao_encontrada_label2.pack(padx=10, pady=0)
            botao_voltar = customtkinter.CTkButton(janela_emails_automaticos_tempo_erro_valor, text = " Voltar ", corner_radius=8, command = janela_emails_automaticos_tempo_erro_valor.destroy)
            botao_voltar.pack(padx=10, pady=10)
            pass
        
    # SWITCH PARA LIGAR E DESLIGAR OS ENVIOS AUTOMÁTICOS
    def switch_emails_automaticos_tempo():
        #print("valor switch emails automaticos tempo->", switch_var1.get())
        with open("emails_automaticos_tempo_ativo.txt","w", encoding = "utf-8") as emails_automaticos_tempo_ativo_arq:       
            emails_automaticos_tempo_ativo = emails_automaticos_tempo_ativo_arq.write(switch_var1.get())
        pass

    def clique_email_final_lista():
        #ir buscar o que foi escrito na entry
        with open("email_para_avisos.txt","w", encoding = "utf-8") as email_para_avisos_arq:       
            email_para_avisos = email_para_avisos_arq.write(email_para_avisos_entry.get())
        #janela de confirmação
        janela_email_para_avisos = customtkinter.CTkToplevel(janela_outras_opcoes)
        janela_email_para_avisos.title("Projeto e-mails ")
        janela_email_para_avisos.geometry("300x150+900+100")

        email_para_avisos_label = customtkinter.CTkLabel(janela_email_para_avisos, text = "E-mail para avisos alterado", font = customtkinter.CTkFont(size=16, weight = "bold"))
        email_para_avisos_label.pack(padx=10, pady=(20,10))

        botao_voltar = customtkinter.CTkButton(janela_email_para_avisos, text = " Voltar ", corner_radius=8, command = janela_email_para_avisos.destroy)
        botao_voltar.pack(padx=10, pady=10)
        pass

    # C O M P O N E N T E S   D A   J A N E L A   -   O U T R A S   O P Ç Õ E S 
    # NUMERO DE EMAILS
    n_emails_label = customtkinter.CTkLabel(outras_opcoes_frame, text = "Quantidade de e-mails a enviar:")
    n_emails_label.place(x=20,y=10)
        #ir buscar a informação que já está guardada no bloco de notas para se saber o que tem lá guardado
    with open("numero_envios.txt", "r", encoding = "utf-8") as numero_envios_arq:
            n_requerido_envios = int(numero_envios_arq.read())
    n_emails = customtkinter.CTkEntry(outras_opcoes_frame, width=40, placeholder_text= n_requerido_envios)
    n_emails.place(x=210, y=10)
    botao_n_emails = customtkinter.CTkButton(outras_opcoes_frame, width=80, text = " Alterar ", corner_radius=8, command = clique_alterar_n_emails)
    botao_n_emails.place(x=260, y=10)

    # ALTERAR EMAIL DE TESTES 
    alterar_email_teste_label = customtkinter.CTkLabel(outras_opcoes_frame, text = "Alterar e-mail de testes:")
    alterar_email_teste_label.place(x=20,y=50)
    with open("email_teste.txt", "r", encoding = "utf-8") as teste_email_arq:
            teste_email = teste_email_arq.read()
    alterar_email_teste_entry = customtkinter.CTkEntry(outras_opcoes_frame, width=200, placeholder_text= teste_email)
    alterar_email_teste_entry.place(x=20, y=80)
    alterar_email_teste_botao = customtkinter.CTkButton(outras_opcoes_frame, width=80, text = " Alterar ", corner_radius=8, command = clique_alterar_email)
    alterar_email_teste_botao.place(x=260, y=80)   

    # IMPORTAÇÃO DE NOVA LISTA DE EMAILS
    nova_lista_label = customtkinter.CTkLabel(outras_opcoes_frame, text = "Importar lista de e-mails nova:")
    nova_lista_label.place(x=20,y=120)
    botao_nova_lista = customtkinter.CTkButton(outras_opcoes_frame, width=80, text = " Importar ", corner_radius=8, command = clique_nova_lista)
    botao_nova_lista.place(x=260, y=120)

    # SWITCH PARA LIGAR E DESLIGAR A REPOSIÇÃO AUTOMÁTICA
    #primeiro tenho de ler o que esá guardado
    with open("reposicao_automatica.txt","r", encoding = "utf-8") as reposicao_automatica_arq:       
            reposicao_automatica = reposicao_automatica_arq.read()
    switch_var = customtkinter.StringVar(value=reposicao_automatica)
    reposicao_automatica_label = customtkinter.CTkLabel(outras_opcoes_frame, text = "Reposição automatica OFF / ON:")
    reposicao_automatica_label.place(x=20,y=160)
    reposicao_automatica_switch = customtkinter.CTkSwitch(outras_opcoes_frame, text=None, switch_width=50, command = switch_reposicao_automatica, variable=switch_var, onvalue="True", offvalue="False")
    reposicao_automatica_switch.place(x=280, y=160)

    # JANELA COM OS EMAILS REJEITADOS PELO OUTLOOK
    rejeitados_outlook_label = customtkinter.CTkLabel(outras_opcoes_frame, text = "E-mails rejeitados pelo Outlook:")
    rejeitados_outlook_label.place(x=20,y=200)
    rejeitados_outlook_botao = customtkinter.CTkButton(outras_opcoes_frame, width=120, text = " Ver lista ", corner_radius=8, command = clique_rejeitados_outlook)
    rejeitados_outlook_botao.place(x=220, y=200)

    # VERIFICAR SE OS EMAILS ESTÃO CORRETAMENTE PREENCHIDOS
    verificar_emails_label = customtkinter.CTkLabel(outras_opcoes_frame, text = "Verificar erros em e-mails: ")
    verificar_emails_label.place(x=20,y=240)
    verificar_emails_botao = customtkinter.CTkButton(outras_opcoes_frame, width=120, text = " Verificar E-mails ", corner_radius=8, command = clique_verificar_emails)
    verificar_emails_botao.place(x=220, y=240) 

    # ADICIONAR DOMINIO
    adicionar_dominio_entry = customtkinter.CTkEntry(outras_opcoes_frame, width=200, placeholder_text= "@dominio.pt")
    adicionar_dominio_entry.place(x=20, y=280)
    adicionar_dominio_botao = customtkinter.CTkButton(outras_opcoes_frame, width=80, text = " Adicionar ", corner_radius=8, command = clique_adicionar_dominio)
    adicionar_dominio_botao.place(x=260, y=280)  

    # ALTERAR TEMPO ENTRE ENVIO DE EMAILS AUTOMÁTICOS
    emails_automaticos_tempo_label = customtkinter.CTkLabel(outras_opcoes_frame, text = "Tempo entre envios, em minutos")
    emails_automaticos_tempo_label.place(x=20,y=320)
    with open("emails_automaticos_tempo.txt", "r", encoding = "utf-8") as emails_automaticos_tempo_arq:
            emails_automaticos_tempo = int(emails_automaticos_tempo_arq.read())
    emails_automaticos_tempo_entry = customtkinter.CTkEntry(outras_opcoes_frame, width=40, placeholder_text= emails_automaticos_tempo)
    emails_automaticos_tempo_entry.place(x=210, y=320)
    botao_emails_automaticos_tempo = customtkinter.CTkButton(outras_opcoes_frame, width=80, text = " Alterar ", corner_radius=8, command = clique_emails_automaticos_tempo)
    botao_emails_automaticos_tempo.place(x=260, y=320)

    # SWITCH PARA LIGAR E DESLIGAR OS ENVIOS AUTOMÁTICOS
    with open("emails_automaticos_tempo_ativo.txt","r", encoding = "utf-8") as emails_automaticos_tempo_ativo_arq:       
            emails_automaticos_tempo_ativo = emails_automaticos_tempo_ativo_arq.read()
    switch_var1 = customtkinter.StringVar(value=emails_automaticos_tempo_ativo)
    emails_automaticos_tempo_ativo_label = customtkinter.CTkLabel(outras_opcoes_frame, text = "Envio automatico OFF / ON:")
    emails_automaticos_tempo_ativo_label.place(x=20,y=360)
    emails_automaticos_tempo_ativo_switch = customtkinter.CTkSwitch(outras_opcoes_frame, text=None, switch_width=50, command = switch_emails_automaticos_tempo, 
                                                                    variable=switch_var1, onvalue="True", offvalue="False")
    emails_automaticos_tempo_ativo_switch.place(x=280, y=360)

    email_para_avisos_label_ = customtkinter.CTkLabel(outras_opcoes_frame, text = "Alterar e-mail para avisos:")
    email_para_avisos_label_.place(x=20,y=400)
    with open("email_para_avisos.txt", "r", encoding = "utf-8") as email_para_avisos_arq:
            email_para_avisos = email_para_avisos_arq.read()
    email_para_avisos_entry = customtkinter.CTkEntry(outras_opcoes_frame, width=200, placeholder_text= email_para_avisos)
    email_para_avisos_entry.place(x=20, y=430)
    email_para_avisos_botao = customtkinter.CTkButton(outras_opcoes_frame, width=80, text = " Alterar ", corner_radius=8, command = clique_email_final_lista)
    email_para_avisos_botao.place(x=260, y=430)   

    botao_voltar = customtkinter.CTkButton(outras_opcoes_frame, text = " Voltar ", corner_radius=8, command = janela_outras_opcoes.destroy)
    botao_voltar.place(x=110, y=480)

    pass

# C O M P O N E N T E S   D A   J A N E L A   I N I C I A L
titulo = customtkinter.CTkLabel(master = janela, text = " |  Enviar e-mails  | ", font = customtkinter.CTkFont(size=30, weight = "bold"))
titulo.pack(padx=10, pady=(40,20)) 

botao_emails = customtkinter.CTkButton(janela, text = " Enviar e-mails ", corner_radius=8, command = clique_enviar_emails)
botao_emails.pack(padx=10, pady=10)

botao_email_teste = customtkinter.CTkButton(janela, text = " Enviar e-mail teste ", corner_radius=8, command = clique_enviar_email_teste)
botao_email_teste.pack(padx=10, pady=10)

botao_mensagem = customtkinter.CTkButton(janela, text = " Mudar a mensagem ", corner_radius=8, command = clique_mudar_email)
botao_mensagem.pack(padx=10, pady=10)

botao_mostrar_emails = customtkinter.CTkButton(janela, text = " Mostrar e-mails ", corner_radius=8, command = clique_mostrar_emails)
botao_mostrar_emails.pack(padx=10, pady=10)

botao_outros_opcoes = customtkinter.CTkButton(janela, text = " Outras opções ", corner_radius=8, command = clique_outras_opcoes)
botao_outros_opcoes.pack(padx=10, pady=10)


botao_sair = customtkinter.CTkButton(janela, text = " Sair ", corner_radius=8, command = janela.destroy)
botao_sair.pack(padx=10, pady=10)


# L O O P   P A R A   O   P R O G R A M A   E S T A R   S E M P R E   A   C O R R E R 
janela.mainloop()

