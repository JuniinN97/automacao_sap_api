import customtkinter as ctk
from PIL import Image
import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import Calendar
import datetime
import getpass
import win32com.client
import win32com.client as win32
from pathlib import Path
import pyperclip 
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl import load_workbook, Workbook
import time
import subprocess
import shutil
from tkinter import Label, Tk
import csv
from pathlib import Path
import glob
import sys
import pandas as pd
from copy import copy
import json
import numpy as np
import requests
from math import isnan, isinf
import re
import unicodedata
from pathlib import Path
import subprocess

# Variáveis globais
usuario_global = ""
senha_global = ""
executando = False
calendarios_visiveis = False
output_textbox_api = None  



'''class DualLogger:
    def __init__(self, log_path, textbox_widget=None):
        self.terminal = sys.__stdout__
        self.log = open(log_path, "a", encoding="utf-8")
        self.textbox = textbox_widget

        usuario = getpass.getuser()
        agora = datetime.datetime.now()
        cabecalho = f"Usuário: {usuario} | Início: {agora.strftime('%d/%m/%Y %H:%M:%S')}\n"
        self.log.write("=" * 80 + "\n")
        self.log.write(cabecalho)
        self.log.write("=" * 80 + "\n")

    def write(self, message):
        # Verifica se a mensagem tem símbolo de sucesso ou erro
        if any(kw in message for kw in ["✅", "❌"]):
            separator = "\n" + "=" * 80 + "\n"
            message = separator + message + separator

        self.terminal.write(message)
        self.log.write(message)

        if self.textbox:
            self.textbox.configure(state="normal")
            self.textbox.insert("end", message)
            self.textbox.see("end")
            self.textbox.configure(state="disabled")
            
    def flush(self):
        self.terminal.flush()
        self.log.flush()
'''

class RedirectPrint:
    def __init__(self, textbox_widget):
        self.textbox = textbox_widget

    def write(self, text):
        if self.textbox:
            self.textbox.configure(state="normal")
            self.textbox.insert("end", text)
            self.textbox.see("end")
            self.textbox.configure(state="disabled")

    def flush(self):
        pass

def iniciar_upload():
    """Executa upload_externalpo() em uma thread separada."""
    thread_upload = threading.Thread(target=upload_externalpo)
    thread_upload.start()
    
# --- Configuração do caminho da pasta SAP ---
pasta_sap = os.path.expanduser(r"~\OneDrive - Accenture\Documents\SAP\SAP GUI")
os.makedirs(pasta_sap, exist_ok=True)

agora = datetime.datetime.now()
nome_log = f"log_{agora.strftime('%d_%m_%H_%M')}.txt"
caminho_log = os.path.join(pasta_sap, nome_log)

# para empacotar recursos
def recurso_caminho(relativo):
    try:
        return os.path.join(sys._MEIPASS, relativo)
    except AttributeError:
        return os.path.join(os.path.abspath("."), relativo)

def atualizar_status(status):
    status_label.configure(text=status)
    print(status)

def atualizar_barra_progresso(valor):
    progress_bar.set(valor)

def voltar_para_inicial():
    global frame_sap
    frame_sap.destroy()
    frame_inicial.pack(expand=True, fill="both")

import logging
timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")
log_file = os.path.join(pasta_sap, f"log_sap_{timestamp}.txt")
logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.INFO, format='%(message)s')


def print_log_sap(msg):
    print(msg)
    logging.info(msg)

def iniciar_automacao_sap():
    def processo_sap():
        atualizar_status("Iniciando automação SAP...")
        atualizar_barra_progresso(0.1)

        for i in range(10, 101, 10):
            time.sleep(0.5)
            atualizar_barra_progresso(i / 100)
            atualizar_status(f"Executando: {i}%")

#        atualizar_status("Automação concluída!")
#        atualizar_barra_progresso(1.0)

    threading.Thread(target=processo_sap).start()

def exibir_sap():
    global frame_sap, progress_bar, status_label, entry_data_de, entry_data_ate, cal_data_de, cal_data_ate, calendarios_visiveis

    frame_inicial.pack_forget()

    frame_sap = ctk.CTkFrame(app, fg_color="#f3f2f9")
    frame_sap.pack(expand=True, fill="both")
    frame_sap.columnconfigure((0, 1, 2), weight=1)

    usuario = getpass.getuser()
    data_hora_local = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    label_usuario_data = ctk.CTkLabel(
        frame_sap, text=f"Usuário: {usuario}  |  Execução: {data_hora_local}",
        font=ctk.CTkFont("Segoe UI", 14), text_color="#1e1e2f"
    )
    label_usuario_data.grid(row=0, column=0, columnspan=3, pady=(20, 10), sticky="n")

    label_data_de = ctk.CTkLabel(frame_sap, text="De (dd.mm.aaaa):", font=("Segoe UI", 12))
    label_data_de.grid(row=1, column=0, pady=10, padx=20, sticky="w")

    entry_data_de = ctk.CTkEntry(frame_sap, width=150, height=40, border_width=0, corner_radius=20)
    entry_data_de.grid(row=1, column=1, pady=10)
    entry_data_de.bind("<KeyRelease>", lambda event: formatar_data(entry_data_de))

    cal_icon_de = ctk.CTkButton(frame_sap, text="📅", width=30, height=30, command=toggle_calendarios)
    cal_icon_de.grid(row=1, column=2, pady=10, padx=10)

    label_data_ate = ctk.CTkLabel(frame_sap, text="Até (dd.mm.aaaa):", font=("Segoe UI", 12))
    label_data_ate.grid(row=2, column=0, pady=10, padx=20, sticky="w")

    entry_data_ate = ctk.CTkEntry(frame_sap, width=150, height=40, border_width=0, corner_radius=20)
    entry_data_ate.grid(row=2, column=1, pady=10)
    entry_data_ate.bind("<KeyRelease>", lambda event: formatar_data(entry_data_ate))

    cal_icon_ate = ctk.CTkButton(frame_sap, text="📅", width=30, height=30, command=toggle_calendarios)
    cal_icon_ate.grid(row=2, column=2, pady=10, padx=10)

    # Ambos calendários no mesmo row, col 1 e 2
    cal_data_de = Calendar(frame_sap, selectmode="day", date_pattern="dd.mm.yyyy",
                           background="white", foreground="black",
                           headersbackground="gray", normalbackground="#dbe3fa", normalforeground="black")
    cal_data_de.grid(row=4, column=1, pady=10, padx=5, sticky="w")
    cal_data_de.grid_remove()
    cal_data_de.bind("<<CalendarSelected>>", lambda e: entry_data_de.delete(0, tk.END) or entry_data_de.insert(0, cal_data_de.get_date()))

    cal_data_ate = Calendar(frame_sap, selectmode="day", date_pattern="dd.mm.yyyy",
                            background="white", foreground="black",
                            headersbackground="gray", normalbackground="#dbe3fa", normalforeground="black")
    cal_data_ate.grid(row=4, column=2, pady=10, padx=5, sticky="w")
    cal_data_ate.grid_remove()
    cal_data_ate.bind("<<CalendarSelected>>", lambda e: entry_data_ate.delete(0, tk.END) or entry_data_ate.insert(0, cal_data_ate.get_date()))

 # Linha com Iniciar + Atualizações
    botoes_linha_1 = ctk.CTkFrame(frame_sap, fg_color="transparent")
    botoes_linha_1.grid(row=5, column=0, columnspan=3, pady=(20, 10))

# Linha com Iniciar + Atualizações
    ctk.CTkButton(
    botoes_linha_1, text="ExternalPO", width=160, height=40,
    fg_color="#2d225d", hover_color="#3a2f85",
    font=ctk.CTkFont("Segoe UI", 14, "bold"),
    corner_radius=20, command=iniciar
    ).pack(side="left", padx=10)

    ctk.CTkButton(
    botoes_linha_1, text="ExternalPO\nAtualizações", width=160, height=40,
    fg_color="#1e6f5c", hover_color="#238f75",
    font=ctk.CTkFont("Segoe UI", 14, "bold"),
    corner_radius=20, #
    command=atualizar
    ).pack(side="left", padx=10)

    # Linha com Voltar (acima da barra de progresso)
    botao_voltar_frame = ctk.CTkFrame(frame_sap, fg_color="transparent")
    botao_voltar_frame.grid(row=6, column=0, columnspan=3, pady=(10, 20))

    ctk.CTkButton(
        botao_voltar_frame, text="Voltar", width=200, height=40,
        fg_color="#ff5c5c", hover_color="#d03f3f",
        font=ctk.CTkFont("Segoe UI", 14, "bold"),
        corner_radius=20,
        command=voltar_para_inicial
    ).pack(pady=10)

    # Barra de progresso e status
    progress_bar = ctk.CTkProgressBar(frame_sap, width=400)
    progress_bar.grid(row=7, column=0, columnspan=3, pady=(10, 5))
    progress_bar.set(0.01)

    status_label = ctk.CTkLabel(frame_sap, text="Aguardando datas...", font=("Segoe UI", 12))
    status_label.grid(row=8, column=0, columnspan=3, pady=10)

def formatar_data(entry):
    data = entry.get()
    data = ''.join(filter(str.isdigit, data))
    if len(data) > 8:
        data = data[:8]
    if len(data) > 4:
        data = data[:2] + '.' + data[2:4] + '.' + data[4:]
    elif len(data) > 2:
        data = data[:2] + '.' + data[2:]
    entry.delete(0, tk.END)
    entry.insert(0, data)

def toggle_calendarios():
    global calendarios_visiveis
    if calendarios_visiveis:
        cal_data_de.grid_remove()
        cal_data_ate.grid_remove()
        calendarios_visiveis = False
    else:
        cal_data_de.grid()
        cal_data_ate.grid()
        calendarios_visiveis = True

def iniciar():
    global executando
    if executando:
        return

    executando = True
    data_de = entry_data_de.get()
    data_ate = entry_data_ate.get()

    if not data_de or not data_ate or data_de == "dd.mm.aaaa" or data_ate == "dd.mm.aaaa":
        messagebox.showerror("Erro", "Preencha as datas DE e ATÉ.")
        executando = False
        return

    try:
        atualizar_status("⏳ Conectando ao SAP...")
        atualizar_barra_progresso(0.02)
        app.update() 

        # Conexão com SAP
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

    except Exception as e:
        atualizar_status("❌ Erro ao conectar ao SAP.")
        print_log_sap("❌ Erro ao conectar ao SAP.")

    finally:
        print("Conexão SAP realizada.")

# ==== Ekko ====

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "SE16N"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        # Define a tabela EKKO
        session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKKO"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()


            # Carregar variante
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO AUTOMACA"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.5)

        print_log_sap("⚙ Iniciando extrações EKKO")
        atualizar_status("⚙ Iniciando extrações EKKO")
        atualizar_barra_progresso(0.03)
        app.update() 

        # Define os valores nos campos LOW e HIGH
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,8]").text = data_de
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,8]").text = data_ate

        
        # Inicia a exportação
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Abre o menu de variantes
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")

        # Seleciona a opção "Carregar" (Load)
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&LOAD")

        # Obtém o grid corretamente
        # Aguarda o carregamento do grid por segurança
        time.sleep(1.5)  # ajuste se necessário

        grid = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")

        # Tenta encontrar a linha onde a coluna VARIANT tem o valor '/RPA'
        for i in range(0, grid.RowCount):
            try:
                valor = grid.GetCellValue(i, "VARIANT")
                if valor.strip().upper() == "/RPA":
                    grid.currentCellRow = i
                    grid.selectedRows = str(i)
                    grid.clickCurrentCell()
                    break
            except:
                pass

        # clicar no botão de OK 
        #session.findById("wnd[1]/tbar[0]/btn[0]").press()

        time.sleep(2)

        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

        # Percorre os elementos da tela em busca de "Planilha eletrônica"
        for i in range(5):  # Ajuste conforme necessário
            try:
                opcao = session.findById(f"wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[{i},0]")
                if opcao.text == "Planilha eletrônica":
                    opcao.select()
                    opcao.setFocus()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Confirma seleção
                    break
            except:
                pass 
        # Gerar o nome do arquivo com base na data atual
        data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
        nome_arquivo = f"EXPORT_EKKO_{data_atual}.XLS"
        print_log_sap("💾 Salvando tabela EKKO...")
        # Define o nome do arquivo e salva
    try:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception as e:
        atualizar_status("❌ Erro ao salvar o arquivo.")
        executando = False
        return
    try:
    # Voltar duas vezes
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")
    
    print_log_sap("✅ Dados Ekko extaidos e salvos com sucesso")    
    atualizar_status("✅ Dados Ekko extaidos e salvos com sucesso")
    atualizar_barra_progresso(0.04)
    app.update() 
    
    
# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("EXPORT_EKKO_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

    def esperar_carregamento(session, timeout=180):
        """ Aguarda o SAP concluir o processamento antes de continuar """
        tempo_inicial = time.time()
        while time.time() - tempo_inicial < timeout:
            try:
                # Verifica se a barra de status está ativa, indicando processamento
                if session.findById("wnd[0]/sbar").text.strip():
                    time.sleep(2)  # Aguarda 2 segundos e tenta novamente
                else:
                    return True  # O SAP terminou o processamento
            except:
                time.sleep(1)  # Se der erro, aguarda e tenta de novo

        # Se o tempo limite foi atingido, exibe a mensagem na interface
        print_log_sap("❌ Não existe documento para exportar.")
        atualizar_status("❌ Não existe documento para exportar.")
        atualizar_barra_progresso(0.5)
        app.update()
        return False  # Retorna falso indicando falha


# ==== Manipulando dados para LFA1   ====

    def localizar_arquivo_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("EXPORT_EKKO_*.xlsx"))

            if not arquivos:
                print_log_sap("Nenhum arquivo encontrado para manipular para LFA1.")
                return None

            return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_m(arquivo_xlsx):
    # Usando pandas para ler o arquivo Excel
        df = pd.read_excel(arquivo_xlsx)

        # Acessando a coluna M (a 13ª coluna, que tem o índice 12)
        if df.shape[1] > 12:  # Verifica se a coluna M (índice 12) existe
            coluna_m = df.iloc[4:, 2]  # Começando da linha 6 (índice 5) e acessando a coluna M (índice 12)

            # Limpando os dados (removendo valores nulos e espaços extras)
            dados_coluna_m = coluna_m.dropna().apply(lambda x: str(x).strip()).tolist()

            return dados_coluna_m
        else:
            print_log_sap("A coluna M não foi encontrada no arquivo.")
            return None

# Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
            print_log_sap(f"Arquivo encontrado: {arquivo}")
            fornecedores = obter_dados_coluna_m(arquivo)

            if fornecedores:
                # Juntar os valores da coluna I com \r\n (quebra de linha para o SAP)
                texto_para_copiar = '\r\n'.join(fornecedores)

        # Copiar para a área de transferência
                try:
                    pyperclip.copy(texto_para_copiar)
                    print("\nDados extraídos e copiados para a área de transferência:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(fornecedores)} valores copiados.")
                except pyperclip.PyperclipException:
                    print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                    print("Dados extraídos e copiados:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(fornecedores)} valores copiados.")
            else:
                print_log_sap("Nenhum dado encontrado na coluna I.")

    print_log_sap("⚙ Iniciando extração LFA1")
    atualizar_status("⚙ Iniciando extração LFA1")
    atualizar_barra_progresso(0.10)
    app.update() 

# ==== EXTRAÇÃO LFA1   ====
    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "LFA1"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO LFA1"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)

    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abre o menu de contexto da toolbar de resultados (botão "Exportar")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # Seleciona a opção "Planilha..." no menu de contexto
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Seleciona o formato de exportação (por exemplo, planilha Excel no formato interno SAP)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirma a seleção do formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print_log_sap("💾 Salvando tabela LFA1...")
    atualizar_status("💾 Salvando tabela LFA1...")
    atualizar_barra_progresso(0.15)
    app.update()

    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"EXPORT_LFA1_{data_atual}.XLS"
    # Define o nome do arquivo a ser salvo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirma a exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
    # Voltar 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    print_log_sap ("✅ Dados LFA extraidos e salvos com sucesso")
    atualizar_status("✅ Dados LFA extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.16)
    app.update() 

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("EXPORT_LFA1_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

# ==== Manipulando dados para EKPO   ====
    print_log_sap("⏳ Manipulando dados para extração EKPO")
    atualizar_status("⏳ Manipulando dados para extração EKPO")
    atualizar_barra_progresso(0.18)
    app.update()

    def localizar_arquivo_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("EXPORT_EKKO_*.xlsx"))

            if not arquivos:
                print_log_sap("Nenhum arquivo encontrado para manipular dados para EKPO.")
                return None

            return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_b(arquivo_xlsx):
            # Usando pandas para ler o arquivo Excel
            df = pd.read_excel(arquivo_xlsx)

            # Acessando a coluna B (a 2ª coluna, que tem o índice 1)
            if df.shape[1] > 1:  # Verifica se a coluna B (índice 1) existe
                coluna_b = df.iloc[4:, 1]  # Começando da linha 6 (índice 5) e acessando a coluna B (índice 1)

                # Limpando os dados (removendo valores nulos e espaços extras)
                dados_coluna_b = coluna_b.dropna().apply(lambda x: str(x).strip()).tolist()

                return dados_coluna_b
            else:
                print_log_sap("A coluna B não foi encontrada no arquivo.")
                return None

        # Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
            print_log_sap(f"Arquivo encontrado: {arquivo}")
            docompras = obter_dados_coluna_b(arquivo)

            if docompras:
                # Juntar os valores da coluna B com \r\n (quebra de linha para o SAP)
                texto_para_copiar = '\r\n'.join(docompras)

                # Copiar para a área de transferência
                try:
                    pyperclip.copy(texto_para_copiar)
                    print("\nDados extraídos e copiados para a área de transferência:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(docompras)} valores copiados.")
                except pyperclip.PyperclipException:
                    print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                    print("Dados extraídos e copiados:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(docompras)} valores copiados.")
            else:
                print_log_sap("Nenhum dado encontrado na coluna B.")

    print_log_sap("⚙ Inicinado extração EKPO")
    atualizar_status("⚙ Inicinado extração EKPO")
    atualizar_barra_progresso(0.20)
    app.update()

# ==== EXTRAÇÃO EKPO   ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKPO"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO EKPO"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)
    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abre o menu de variantes
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")

        # Seleciona a opção "Carregar" (Load)
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&LOAD")

        # Obtém o grid corretamente
        # Aguarda o carregamento do grid por segurança
    time.sleep(1.5)  # ajuste se necessário

    grid = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")

        # Tenta encontrar a linha onde a coluna VARIANT tem o valor '/RPA'
    for i in range(0, grid.RowCount):
            try:
                valor = grid.GetCellValue(i, "VARIANT")
                if valor.strip().upper() == "/RPA":
                    grid.currentCellRow = i
                    grid.selectedRows = str(i)
                    grid.clickCurrentCell()
                    break
            except:
                pass

    # Abre o menu de contexto da toolbar de resultados (botão "Exportar")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # Seleciona a opção "Planilha..." no menu de contexto
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Seleciona o formato de exportação (por exemplo, planilha Excel no formato interno SAP)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirma a seleção do formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    print_log_sap("💾 Salvando tabela EKPO...")
    atualizar_status("💾 Salvando tabela EKPO...")
    atualizar_barra_progresso(0.25)
    app.update()

    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"EXPORT_EKPO_{data_atual}.XLS"
    # Define o nome do arquivo a ser salvo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirma a exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
    # Voltar 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    print_log_sap("✅ Dados EKPO extraidos e salvos com sucesso")
    atualizar_status("✅ Dados EKPO extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.26)
    app.update() 


# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("EXPORT_EKPO_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

# ====  Manipulando dados para MARA  ====

    atualizar_status("⏳ Manipulando dados para extração Mara")
    atualizar_barra_progresso(0.27)
    app.update()

# Função para localizar o arquivo mais recente na pasta
    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("EXPORT_EKPO_*.xlsx"))

        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para manipulação de dados para Mara.")
            return None

        # Encontrar o arquivo mais recente baseado na data de modificação
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    # Função para obter os dados da coluna E a partir de um arquivo XLSX
    def obter_dados_coluna_e(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        dados_coluna_e = []

        # Acessando a coluna E (5ª coluna)
        for row in sheet.iter_rows(min_row=6, min_col=5, max_col=5):  # Coluna E = índice 5
            for cell in row:
                if cell.value is not None:
                    dados_coluna_e.append(str(cell.value))  # Converte o valor para string

        return dados_coluna_e

    # Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
        print_log_sap(f"Arquivo encontrado: {arquivo}")
        tipmaterial = obter_dados_coluna_e(arquivo)

        if tipmaterial:
            print("Dados da coluna E:", '\n'.join(tipmaterial))
            
            # Copia os valores da coluna E para a área de transferência, um valor por linha
            pyperclip.copy('\r\n'.join(tipmaterial))  # Usando '\r\n' para SAP reconhecer quebra de linha
            print_log_sap("Valores copiados para a área de transferência.")
        else:
            print_log_sap("Nenhum dado encontrado na coluna E.")

    print_log_sap("⚙ Inicinado extração MARA...")
    atualizar_status("⚙ Inicinado extração MARA...")
    atualizar_barra_progresso(0.30)
    app.update() 

# ==== EXTRAÇÃO MARA   ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "MARA"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO MARA"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)
    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abre o menu de contexto da toolbar de resultados (botão "Exportar")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # Seleciona a opção "Planilha..." no menu de contexto
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Seleciona o formato de exportação (por exemplo, planilha Excel no formato interno SAP)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirma a seleção do formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print_log_sap("💾 Salvando tabela MARA...")
    atualizar_status("💾 Salvando tabela MARA...")
    atualizar_barra_progresso(0.35)
    app.update()

    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"EXPORT_MARA_{data_atual}.XLS"
    # Define o nome do arquivo a ser salvo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirma a exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
    # Voltar 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    print_log_sap("✅ Dados MARA extraidos e salvos com sucesso")
    atualizar_status("✅ Dados MARA extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.36)
    app.update() 

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("EXPORT_MARA_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

#=== Manipulando dados para EKET ===
    print_log_sap("⏳ Manipulando dados para extração EKET")
    atualizar_status("⏳ Manipulando dados para extração EKET")
    atualizar_barra_progresso(0.40)
    app.update()

# Função para localizar o arquivo mais recente na pasta
    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("EXPORT_EKPO_*.xlsx"))

        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para Manipular dodos para EKET.")
            return None

        # Encontrar o arquivo mais recente baseado na data de modificação
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    # Função para obter os dados da coluna B a partir de um arquivo XLSX
    def obter_dados_coluna_b(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        dados_coluna_b = []

        # Acessando a coluna B (2ª coluna)
        for row in sheet.iter_rows(min_row=6, min_col=2, max_col=2):  # Coluna B = índice 2
            for cell in row:
                if cell.value is not None:
                    dados_coluna_b.append(str(cell.value))  # Converte o valor para string

        return dados_coluna_b

    # Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
        print_log_sap(f"Arquivo encontrado: {arquivo}")
        doccompras = obter_dados_coluna_b(arquivo)

        if doccompras:
            print("Dados da coluna B (DocCompras):", '\n'.join(doccompras))

            # Copia os valores da coluna B para a área de transferência, um valor por linha
            pyperclip.copy('\r\n'.join(doccompras))  # Usando '\r\n' para SAP reconhecer quebra de linha
            print_log_sap("Valores copiados para a área de transferência.")
        else:
            print_log_sap("Nenhum dado encontrado na coluna B.")

    print_log_sap("⚙ Inicinado extração EKET")
    atualizar_status("⚙ Inicinado extração EKET")
    atualizar_barra_progresso(0.42)
    app.update() 

# ==== EXTRAÇÃO EKET   ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKET"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO EKET"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)
    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abre o menu de contexto da toolbar de resultados (botão "Exportar")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # Seleciona a opção "Planilha..." no menu de contexto
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Seleciona o formato de exportação (por exemplo, planilha Excel no formato interno SAP)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirma a seleção do formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print_log_sap("💾 Salvando tabela EKET...")
    atualizar_status("💾 Salvando tabela EKET...")
    atualizar_barra_progresso(0.45)
    app.update()

    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"EXPORT_EKET_{data_atual}.XLS"
    # Define o nome do arquivo a ser salvo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirma a exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
    # Voltar 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    print_log_sap("✅ Dados EKET extraidos e salvos com sucesso")
    atualizar_status("✅ Dados EKET extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.46)
    app.update()

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("EXPORT_EKET_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")


# ==== Ekko CONTRATO====
    # Acessa a transação SE16N para a tabela EKKO
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKKO"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carrega a variante /EXTPO CONT
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO CONT"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)

    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)

# ==== Manipulando dados para EKKO CONTRATO   ====

# Função para localizar o arquivo mais recente na pasta
    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("EXPORT_EKPO_*.xlsx"))

        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para manipular dados referente EKKO CONTRATO.")
            return None

        # Encontrar o arquivo mais recente baseado na data de modificação
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    # Função para obter os dados da coluna B a partir de um arquivo XLSX
    def obter_dados_coluna_b(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        dados_coluna_b = []

        # Acessando a coluna O (2ª coluna)
        for row in sheet.iter_rows(min_row=6, min_col=15, max_col=15):  # Coluna O = índice 2
            for cell in row:
                if cell.value is not None:
                    dados_coluna_b.append(str(cell.value))  # Converte o valor para string

        return dados_coluna_b

    # Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
        print_log_sap(f"Arquivo encontrado: {arquivo}")
        doccompras = obter_dados_coluna_b(arquivo)

        if doccompras:
            print("Dados da coluna B (DocCompras):", '\n'.join(doccompras))

            # Copia os valores da coluna B para a área de transferência, um valor por linha
            pyperclip.copy('\r\n'.join(doccompras))  # Usando '\r\n' para SAP reconhecer quebra de linha
            print_log_sap("Valores copiados para a área de transferência.")
        else:
            print_log_sap("Nenhum dado encontrado na coluna B.")

    print_log_sap("⚙ Inicinado extração EKET")
    atualizar_status("⚙ Inicinado extração EKET")
    atualizar_barra_progresso(0.50)
    app.update() 
    
    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    atualizar_status("⚙ Iniciando extrações EKKO CONTRATO")
    atualizar_barra_progresso(0.53)
    app.update()

    esperar_carregamento(session)    

    mensagem = session.findById("wnd[0]/sbar").Text
    if "Nenhum valor encontrado" in mensagem:
        atualizar_status("⚠ Nenhum dado encontrado para exportar.")
        atualizar_barra_progresso(0.55)

        # Voltar para tela inicial
        try:
            session.findById("wnd[0]/tbar[0]/btn[12]").press()
        except:
            pass

        # Pular exportação e conversão — segue direto com o restante do fluxo
    else:
        # --- Continua se houver dados ---

        # Abre o menu de variantes
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&LOAD")
        time.sleep(1.5)

        # Seleciona a variante /RPA CONT
        grid = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")

        # Tenta encontrar a linha onde a coluna VARIANT tem o valor '/RPA'
        for i in range(0, grid.RowCount):
            try:
                valor = grid.GetCellValue(i, "VARIANT")
                if valor.strip().upper() == "/RPA_CONT":
                    grid.currentCellRow = i
                    grid.selectedRows = str(i)
                    grid.clickCurrentCell()
                    break
            except:
                pass

        # clicar no botão de OK 
        #session.findById("wnd[1]/tbar[0]/btn[0]").press()

        time.sleep(2)

        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

        # Seleciona "Planilha eletrônica"
        for i in range(5):
            try:
                opcao = session.findById(f"wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[{i},0]")
                if opcao.text == "Planilha eletrônica":
                    opcao.select()
                    opcao.setFocus()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    break
            except:
                pass

        # Gera o nome do arquivo
        data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
        nome_arquivo = f"EXPORT_CONTRATO_{data_atual}.XLS"

        # Salva o arquivo
        try:
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception as e:
            atualizar_status("❌ Erro ao salvar o arquivo.")
            executando = False
        else:
            # Volta para a tela anterior
            try:
                session.findById("wnd[0]/tbar[0]/btn[12]").press()
            except Exception as e:
                atualizar_status(f"❌ Erro inesperado: {e}")

            print_log_sap("✅ Dados EKKO extraídos e salvos com sucesso")
            # Finaliza com sucesso
            atualizar_status("✅ Dados EKKO extraídos e salvos com sucesso")

            # ==== Conversão de XLS para XLSX ====
            def converter_xls_para_xlsx(caminho_arquivo_xls):
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False
                try:
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

            def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("EXPORT_CONTRATO_*.xls"))
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
                return max(arquivos, key=os.path.getmtime)

            # Executa conversão
            arquivo_xls = encontrar_arquivo_export_mais_recente()
            if arquivo_xls:
                print(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

        atualizar_barra_progresso(0.60)
        app.update()

# ==== Manipulando dados para USR21   ====

    def localizar_arquivo_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("EXPORT_EKKO_*.xlsx"))

            if not arquivos:
                print_log_sap("Nenhum arquivo encontrado para manipulação da URS21.")
                return None

            return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_m(arquivo_xlsx):
    # Usando pandas para ler o arquivo Excel
        df = pd.read_excel(arquivo_xlsx)

        # Acessando a coluna M (a 13ª coluna, que tem o índice 12)
        if df.shape[1] > 6:  # Verifica se a coluna M (índice 12) existe
            coluna_m = df.iloc[4:, 6]  # Começando da linha 6 (índice 5) e acessando a coluna M (índice 12)

            # Limpando os dados (removendo valores nulos e espaços extras)
            dados_coluna_m = coluna_m.dropna().apply(lambda x: str(x).strip()).tolist()

            return dados_coluna_m
        else:
            print_log_sap("A coluna G não foi encontrada no arquivo.")
            return None

# Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
            print_log_sap(f"Arquivo encontrado: {arquivo}")
            criadopor = obter_dados_coluna_m(arquivo)

            if criadopor:
                # Juntar os valores da coluna I com \r\n (quebra de linha para o SAP)
                texto_para_copiar = '\r\n'.join(criadopor)

        # Copiar para a área de transferência
                try:
                    pyperclip.copy(texto_para_copiar)
                    print("\nDados extraídos e copiados para a área de transferência:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(criadopor)} valores copiados.")
                except pyperclip.PyperclipException:
                    print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                    print("Dados extraídos e copiados:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(criadopor)} valores copiados.")
            else:
                print_log_sap("Nenhum dado encontrado na coluna I.")

    print_log_sap("⚙ Iniciando extração USR21")
    atualizar_status("⚙ Iniciando extração USR21")
    atualizar_barra_progresso(0.64)
    app.update() 

# ==== EXTRAÇÃO USR21  ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "USR21"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # (Opcional) Carregar variante, se necessário
    # session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    # session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO USR21"
    # session.findById("wnd[1]/tbar[0]/btn[0]").press()
    # time.sleep(0.5)

    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)

    # Colar da área de transferência
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Executar
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abrir menu de exportação
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Selecionar formato de exportação
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirmar formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    print_log_sap("💾 Salvando tabela USR21...")
    # Atualizar status e progresso
    atualizar_status("💾 Salvando tabela USR21...")
    atualizar_barra_progresso(0.68)
    app.update()

    # Definir nome do arquivo com data e hora
    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"EXPORT_USR21_{data_atual}.XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirmar exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Tentar voltar à tela anterior
    try:
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    # Finalizar status e barra
    print_log_sap("✅ Dados USR21 extraídos e salvos com sucesso")
    atualizar_status("✅ Dados USR21 extraídos e salvos com sucesso")
    atualizar_barra_progresso(0.70)
    app.update()

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("EXPORT_USR21_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")


# ==== Manipulando dados para ADR6   ====

    def localizar_arquivo_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("EXPORT_USR21_*.xlsx"))

            if not arquivos:
                print_log_sap("Nenhum arquivo encontrado para manipluar dados para ADR6.")
                return None

            return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_m(arquivo_xlsx):
    # Usando pandas para ler o arquivo Excel
        df = pd.read_excel(arquivo_xlsx)

        # Acessando a coluna M (a 13ª coluna, que tem o índice 12)
        if df.shape[1] > 2:  # Verifica se a coluna M (índice 12) existe
            coluna_m = df.iloc[4:, 2]  # Começando da linha 6 (índice 5) e acessando a coluna M (índice 12)

            # Limpando os dados (removendo valores nulos e espaços extras)
            dados_coluna_m = coluna_m.dropna().apply(lambda x: str(x).strip()).tolist()

            return dados_coluna_m
        else:
            print_log_sap("A coluna G não foi encontrada no arquivo.")
            return None

# Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
            print_log_sap(f"Arquivo encontrado: {arquivo}")
            criadopor = obter_dados_coluna_m(arquivo)

            if criadopor:
                # Juntar os valores da coluna I com \r\n (quebra de linha para o SAP)
                texto_para_copiar = '\r\n'.join(criadopor)

        # Copiar para a área de transferência
                try:
                    pyperclip.copy(texto_para_copiar)
                    print("\nDados extraídos e copiados para a área de transferência:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(criadopor)} valores copiados.")
                except pyperclip.PyperclipException:
                    print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                    print("Dados extraídos e copiados:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(criadopor)} valores copiados.")
            else:
                print_log_sap("Nenhum dado encontrado na coluna I.")

    print_log_sap("⚙ Iniciando extração ADR6")
    atualizar_status("⚙ Iniciando extração ADR6")
    atualizar_barra_progresso(0.75)
    app.update()

# ==== EXTRAÇÃO ADR6 ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "ADR6"
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""  # Sem limite de linhas
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Abrir seleção múltipla (linha 3 da seleção - coluna 5)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").press()
    time.sleep(0.5)

    # Colar dados da área de transferência
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Executar a transação
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Iniciar exportação
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Selecionar formato .XLS (Spreadsheet)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print_log_sap("💾 Salvando tabela ADR6...")
    # Atualizar status e progresso
    atualizar_status("💾 Salvando tabela ADR6...")
    atualizar_barra_progresso(0.80)
    app.update()

    # Gerar nome do arquivo com data/hora
    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"EXPORT_ADR6_{data_atual}.XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Tentar retornar à tela anterior
    try:
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    # Finalizar status
    print_log_sap("✅ Dados ADR6 extraídos e salvos com sucesso")
    atualizar_status("✅ Dados ADR6 extraídos e salvos com sucesso")
    atualizar_barra_progresso(0.83)
    app.update()

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("EXPORT_ADR6_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

# ==== MANIPULANDO DADOS MM03 ====
    print_log_sap("⏳ Manipulando dados MM03")
    atualizar_status("⏳ Manipulando dados MM03")
    atualizar_barra_progresso(0.85)
    app.update()

    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("EXPORT_MARA_*.xlsx"))
        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para manipular dados para MM03.")
            return None
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_docmateriais(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        return [str(c.value) for c in sheet['B'][5:] if c.value is not None]

    def existe(session, id):
        try:
            session.findById(id)
            return True
        except:
            return False

    def select_gui_table_row_by_text(session, field_text, column_index):
        tabela = session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
        for i in range(tabela.RowCount):
            try:
                celula = session.findById(f"wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[{column_index},{i}]")
                texto = celula.text.strip()
                if texto.upper() == field_text.upper():
                    tabela.getAbsoluteRow(i).selected = True
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    return
            except:
                pass
        raise Exception(f"Texto '{field_text}' não encontrado na coluna {column_index}")

    def salvar_resultados(lista_com_texto, lista_sem_texto, pasta_destino):
        # Obtendo data e hora no formato dd_mm_hh_mm_ss
        timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")
        
        # Caminhos dos arquivos
        path_com = pasta_destino / f"mm03_com_texto_{timestamp}.txt"
        path_sem = pasta_destino / f"mm03_sem_texto_{timestamp}.txt"
        
        # Salvando como Excel
        df = pd.DataFrame(lista_com_texto)
        df.to_excel(path_com, index=False, header=False, engine="openpyxl")

        # Salvando como TXT
        with open(path_sem, "w", encoding="utf-8") as file:
            file.write("\n".join(lista_sem_texto))

        print_log_sap(f"Arquivos salvos:\nExcel: {path_com}\nTXT: {path_sem}")

        with open(path_com, "w", encoding="utf-8") as f_com:
            for numero, texto in lista_com_texto:
                texto_linha_unica = texto.replace("\n", " ").replace("\r", "").strip()
                f_com.write(f"{numero} | {texto_linha_unica}\n")

        with open(path_sem, "w", encoding="utf-8") as f_sem:
            for numero in lista_sem_texto:
                f_sem.write(f"{numero}\n")

    # --- Execução principal ---

    arquivo = localizar_arquivo_mais_recente()
    if not arquivo:
        print_log_sap("Arquivo de origem não encontrado.")
        exit()

    docmateriais = obter_docmateriais(arquivo)
    if not docmateriais:
        print_log_sap("Nenhum número encontrado na coluna C.")
        exit()

    print_log_sap("⚙ Iniciando extrações MM03")
    atualizar_status("⚙ Iniciando extrações MM03")
    atualizar_barra_progresso(0.90)
    app.update() 

    com_texto = []
    sem_texto = []
    pasta_destino = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"

    for numero in docmateriais:
        try:
            # Transação MM03
            session.findById("wnd[0]/tbar[0]/okcd").text = "MM03"
            session.findById("wnd[0]/tbar[0]/btn[0]").press()

            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = numero
            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = len(numero)
            session.findById("wnd[0]/tbar[0]/btn[0]").press()

            select_gui_table_row_by_text(session, field_text="Dados básicos 1", column_index=0)
            session.findById("wnd[0]/tbar[1]/btn[30]").press()
            session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05").select()
            time.sleep(1)

            texto_completo = ""

            try:
                control = session.findById(
                    "wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/"
                    "subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell"
                )

                if hasattr(control, "RowCount"):
                    linhas = [control.GetCellValue(i, 0) for i in range(control.RowCount)]
                    texto_completo = "\n".join(linhas)
                else:
                    try:
                        texto_completo = control.Text
                    except Exception:
                        texto_completo = ""

            except:
                try:
                    texto_completo = session.findById(
                        "wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/"
                        "subSUB2:SAPLMGD1:2031/txtRSTXT-TXLINE"
                    ).Text
                except:
                    texto_completo = ""

            if texto_completo.strip():
                com_texto.append((numero, texto_completo))
                print(f"Texto capturado para {numero}")
            else:
                sem_texto.append(numero)
                print(f"Sem texto para {numero}")

            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()

        except Exception as e:
            print_log_sap(f"Erro ao processar {numero}: {e}")
            sem_texto.append(numero)
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except:
                pass

    salvar_resultados(com_texto, sem_texto, pasta_destino)
    print_log_sap("✅ Dados MM03 extraidos e salvos com sucesso")

    atualizar_status("✅ Dados MM03 extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.94)
    app.update()  

# ==== MANIPULANDO DADOS ME23N ====
    print_log_sap("⏳ Manipulando dados ME23N")
    atualizar_status("⏳ Manipulando dados ME23N")
    atualizar_barra_progresso(0.96)
    app.update()

    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("EXPORT_EKKO_*.xlsx"))
        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para manipulação de dados para ME23N.")
            return None
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_docmateriais(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        return [str(c.value) for c in sheet['B'][5:] if c.value is not None]

    def extrair_texto_texto_de_cabecalho(session):
        try:
            time.sleep(1)
            editor_id = (
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/"
                "subSUB1:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/"
                "subSUB1:SAPLMEGUI:1102/"
                "tabsHEADER_DETAIL/tabpTABHDT3/"
                "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/"
                "subTEXTS:SAPLMMTE:0100/"
                "subEDITOR:SAPLMMTE:0101/"
                "cntlTEXT_EDITOR_0101/shellcont/shell"
            )
            editor = session.findById(editor_id)
            try:
                texto = editor.text
            except:
                texto = editor.getProperty("Text")
            return texto.strip()
        except Exception:
            return ""


    def salvar_resultados(lista_com_texto, lista_sem_texto, pasta_destino):
        timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")
        path_com = Path(pasta_destino) / f"me23n_com_texto_{timestamp}.txt"
        path_sem = Path(pasta_destino) / f"me23n_sem_texto_{timestamp}.txt"

        with open(path_com, "w", encoding="utf-8") as f_com:
            for numero, texto in lista_com_texto:
                texto_linha_unica = texto.replace("\n", " ").replace("\r", "").strip()
                f_com.write(f"{numero} | {texto_linha_unica}\n")

        with open(path_sem, "w", encoding="utf-8") as f_sem:
            for numero in lista_sem_texto:
                f_sem.write(f"{numero}\n")

        print(f"\nArquivos salvos:")
        print_log_sap(f"✓ Pedidos COM texto: {path_com}")
        print_log_sap(f"✓ Pedidos SEM texto: {path_sem}")

    print_log_sap("⚙ Iniciando extrações ME23N")
    def executar_automacao_ME23N():
        arquivo = localizar_arquivo_mais_recente()
        if not arquivo:
            return

        docmateriais = obter_docmateriais(arquivo)

        if not docmateriais:
            print_log_sap("Nenhum número de pedido encontrado.")
            return

        pasta_base = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        os.makedirs(pasta_base, exist_ok=True)

        lista_com_texto = []
        lista_sem_texto = []

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME23N"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        time.sleep(2)

        if not extrair_texto_texto_de_cabecalho (session, "Textos"):
            print_log_sap("Aba 'Textos' não encontrada.")
            return
        time.sleep(1)

        for pedido in docmateriais:
            try:
                session.findById("wnd[0]/tbar[1]/btn[17]").press()
                time.sleep(0.5)

                session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = pedido
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                time.sleep(2)

                texto = extrair_texto_texto_de_cabecalho(session)

                if texto:
                    lista_com_texto.append((pedido, texto))
                else:
                    lista_sem_texto.append(pedido)

            except Exception as e:
                print_log_sap(f"Erro com o pedido {pedido}: {str(e)}")
                lista_sem_texto.append(f"{pedido} (erro)")
                continue

        salvar_resultados(lista_com_texto, lista_sem_texto, pasta_base)

    executar_automacao_ME23N()

    # Apaga o xls
    pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"

    # Apaga todos os arquivos .xls da pasta
    for arquivo in pasta.glob("*.xls"):
        try:
            os.remove(arquivo)
            print(f"🗑 Arquivo removido: {arquivo.name}")
        except Exception as e:
            print(f"❌ Erro ao remover {arquivo.name}: {e}")

    print_log_sap("✅ Dados ME23N extraidos e salvos com sucesso")
    atualizar_status("✅ Dados ME23N extraidos e salvos com sucesso")
    atualizar_barra_progresso(1)
    app.update()  

        # PRECO POR

    pasta_base = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
    os.makedirs(pasta_base, exist_ok=True)

    # Procura todos os arquivos CDHDR_EKPO_*.xlsx dentro da pasta (sem subpastas)
    arquivos = list(pasta_base.glob("EXPORT_EKPO_*.xlsx"))

    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo EXPORT_EKPO_*.xlsx encontrado na pasta SAP GUI.")

    for arquivo in arquivos:
        print_log_sap(f"Processando arquivo: {arquivo}")

        wb = load_workbook(arquivo)
        ws = wb.active

        # Extrair dados a partir da linha 6
        dados = [row for row in ws.iter_rows(min_row=6, values_only=True)]
        df = pd.DataFrame(dados)

        # Verifica colunas suficientes (mínimo 19 colunas)
        if df.shape[1] < 19:
            print_log_sap(f"Arquivo {arquivo.name} ignorado: menos de 19 colunas.")
            continue

        # Coluna G = índice 6, coluna S = índice 18 (0-based)
        coluna_preco = pd.to_numeric(df.iloc[:, 6], errors='coerce')
        coluna_divisao = pd.to_numeric(df.iloc[:, 18], errors='coerce')

        # Divisão segura: evita divisão por zero ou valores NaN
        def safe_divide(x, y):
            if pd.isna(x) or pd.isna(y) or y == 0:
                return None
            else:
                return x / y

        df['valor_ajustado'] = [safe_divide(x, y) for x, y in zip(coluna_preco, coluna_divisao)]

        # Escrever cabeçalho na célula T4 (coluna 20)
        ws.cell(row=4, column=20).value = "valor_ajustado"

        # Escrever valores ajustados na coluna T, a partir da linha 6
        for i, valor in enumerate(df['valor_ajustado'], start=6):
            ws.cell(row=i, column=20).value = valor

        # Preparar nome do arquivo novo
        timestamp = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
        novo_nome = f"EXPORT_EKPO_{timestamp}.xlsx"
        novo_caminho = arquivo.parent / novo_nome
        wb.save(novo_caminho)

        print_log_sap(f"Arquivo salvo como {novo_caminho}")

    print_log_sap("✅ Processamento concluído para todos os arquivos.")

def exibir_api():

    global frame_api, output_textbox_api

    #sys.stdout = DualLogger(caminho_log, output_textbox_api)
    sys.stderr = sys.stdout

    frame_inicial.pack_forget()
    frame_api = ctk.CTkFrame(app, fg_color="#f3f2f9")
    frame_api.pack(expand=True, fill="both")
    frame_api.columnconfigure((0, 1, 2), weight=1)

    # Usuário e Data
    usuario = getpass.getuser()
    data_hora_local = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    label_usuario_data = ctk.CTkLabel(
        frame_api,
        text=f"Usuário: {usuario}  |  Execução: {data_hora_local}",
        font=ctk.CTkFont("Segoe UI", 14),
        text_color="#1e1e2f"
    )
    label_usuario_data.grid(row=0, column=0, columnspan=3, pady=(20, 10), sticky="n")

    # Título
    label_titulo = ctk.CTkLabel(
        frame_api,
        text="Upload via API",
        font=ctk.CTkFont("Segoe UI", 18, "bold"),
        text_color="#1e1e2f"
    )
    label_titulo.grid(row=1, column=0, columnspan=3, pady=(10, 20))

    # Status label
    status_label = ctk.CTkLabel(frame_api, text="", font=("Segoe UI", 12), text_color="green")
    status_label.grid(row=2, column=0, columnspan=3, pady=(10, 5))

    # Botões Upload
    botoes = ctk.CTkFrame(frame_api, fg_color="transparent")
    botoes.grid(row=3, column=0, columnspan=3, pady=(10, 10))
    botoes.grid_columnconfigure((0, 1), weight=1)

    ctk.CTkButton(
        botoes,
        text="Upload ExternalPO",
        width=160,
        height=40,
        fg_color="#2d225d",
        hover_color="#3a2f85",
        font=ctk.CTkFont("Segoe UI", 14, "bold"),
        corner_radius=20,
        command=iniciar_upload
    ).grid(row=0, column=0, padx=10)

    ctk.CTkButton(
        botoes,
        text="Upload ExternalPO\nAtualizações",
        width=160,
        height=40,
        fg_color="#1e6f5c",
        hover_color="#238f75",
        font=ctk.CTkFont("Segoe UI", 14, "bold"),
        corner_radius=20,
        command=upload_atualizacoes
    ).grid(row=0, column=1, padx=10)

    # Caixa de log
    output_textbox_api = ctk.CTkTextbox(frame_api, width=580, height=200, corner_radius=10)
    output_textbox_api.grid(row=4, column=0, columnspan=3, pady=(10, 10), padx=20)
    output_textbox_api.configure(state="disabled")

    print("🚀 Interface da API carregada!")  # Deve aparecer na interface gráfica

    # Botão Voltar
    ctk.CTkButton(
        frame_api,
        text="Voltar",
        width=200,
        height=40,
        fg_color="#ff5c5c",
        hover_color="#d03f3f",
        font=ctk.CTkFont("Segoe UI", 14, "bold"),
        corner_radius=20,
        command=lambda: (frame_api.destroy(), frame_inicial.pack(fill="both", expand=True))
    ).grid(row=5, column=0, columnspan=3, pady=(10, 30))


def upload_externalpo():
    global output_textbox_api
    #sys.stdout = DualLogger(caminho_log, output_textbox_api)
    sys.stdout = RedirectPrint(output_textbox_api)
    sys.stderr = RedirectPrint(output_textbox_api)

# ---- CIF ----
    print("📤 Adicionando IncTm e Incotm.2\n")

    # Função para verificar se uma célula está realmente vazia
    def is_realmente_vazio(valor):
        if valor is None:
            return True
        valor_str = str(valor)
        valor_str = unicodedata.normalize('NFKC', valor_str).strip()
        return valor_str == ""

    # Obter nome do usuário e montar o caminho
    usuario = getpass.getuser()
    pasta = Path(f"C:/Users/{usuario}/OneDrive - Accenture/Documents/SAP/SAP GUI")

    # Localizar o arquivo EXPORT_EKKO_
    arquivos_ekko = list(pasta.glob("EXPORT_EKKO_*.xlsx"))
    if not arquivos_ekko:
        print("Arquivo EXPORT_EKKO_ não encontrado.")
    else:
        caminho_arquivo = arquivos_ekko[0]
        wb = openpyxl.load_workbook(caminho_arquivo)
        ws = wb.active

        # Preencher coluna F (coluna 6) a partir da linha 6 com "CIF" se estiver vazia
        for row in ws.iter_rows(min_row=6, min_col=6, max_col=6):
            cell = row[0]
            if is_realmente_vazio(cell.value):
                cell.value = "CIF"

        # Preencher coluna M (coluna 13) a partir da linha 6 com "Custo, seguro & frete" se estiver vazia
        for row in ws.iter_rows(min_row=6, min_col=13, max_col=13):
            cell = row[0]
            if is_realmente_vazio(cell.value):
                cell.value = "Custo, seguro & frete"

        # Salvar alterações
        wb.save(caminho_arquivo)
        print(f"Arquivo atualizado com sucesso: {caminho_arquivo}")
    

    # ---- GRUPO MERCADORIA -----
    print("📤 Convertendo Grupo Mercadoria\n")

    # --- Caminho da pasta SAP ---
    pasta_sap = os.path.expanduser(r"~\OneDrive - Accenture\Documents\SAP\SAP GUI")
    if not os.path.exists(pasta_sap):
        os.makedirs(pasta_sap)

    # --- Localiza o arquivo EKPO mais recente ---
    arquivos_ekpo = [f for f in os.listdir(pasta_sap) if f.startswith("EXPORT_EKPO_") and f.endswith(".xlsx")]
    arquivos_ekpo.sort(key=lambda x: os.path.getmtime(os.path.join(pasta_sap, x)), reverse=True)
    arquivo_ekpo = os.path.join(pasta_sap, arquivos_ekpo[0]) if arquivos_ekpo else None

    # --- Caminho do arquivo GrupoMercadoria ---
    arquivo_grupo = os.path.join(pasta_sap, "GrupoMercadoria.xlsx")

    # --- Verificação de existência dos arquivos ---
    if not arquivo_ekpo or not os.path.exists(arquivo_grupo):
        print("Arquivo EKPO ou GrupoMercadoria não encontrado.")
        exit()

    # --- Carrega os dados do GrupoMercadoria ---
    wb_grupo = openpyxl.load_workbook(arquivo_grupo)
    ws_grupo = wb_grupo.active

    # Cria dicionário: {codigo_grupo: nome_grupo}
    mapa_grupo = {}
    for row in ws_grupo.iter_rows(min_row=2, values_only=True):  # assumindo cabeçalho
        if row[0] is not None and row[1] is not None and row[2] is not None:
            mapa_grupo[str(row[0]).strip()] = str(row[2]).strip()

    # --- Carrega o arquivo EKPO ---
    wb_ekpo = openpyxl.load_workbook(arquivo_ekpo)
    ws_ekpo = wb_ekpo.active

    # --- Substitui valores da coluna Q (17ª) a partir da linha 6 ---
    for row in ws_ekpo.iter_rows(min_row=6, min_col=17, max_col=17):
        cell = row[0]
        valor_original = str(cell.value).strip() if cell.value is not None else ""
        if valor_original in mapa_grupo:
            cell.value = mapa_grupo[valor_original]

    # --- Salva com nome dinâmico baseado na data e hora ---
    agora = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")
    nome_arquivo = f"EXPORT_EKPO_{agora}.xlsx"
    caminho_saida = os.path.join(pasta_sap, nome_arquivo)
    wb_ekpo.save(caminho_saida)

    print(f"[{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Arquivo salvo como: {caminho_saida}")

    start_time = datetime.datetime.now()
    start_format = start_time.strftime("%d_%m às %H_%M_%S")
        
    pasta_base = os.path.join(os.path.expanduser("~"), r"OneDrive - Accenture\Documents\SAP\SAP GUI")

    # Timestamp para nomear o arquivo
    timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")

    log_file = os.path.join(pasta_base, f"log_api_{timestamp}.txt")

    import logging

    log_file = os.path.join(pasta_base, f"log_api_{timestamp}.txt")
    logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.INFO, format='%(message)s')

    def print_log(msg):
        print(msg)
        logging.info(msg)


    # Registrar o tempo de início
    start_time = datetime.datetime.now()
    start_format = start_time.strftime("%d_%m às %H_%M_%S")

    def extract_error_details(response):
        try:
            error_data = response.json()
            detailed_errors = []

            if "errors" in error_data:
                for key, messages in error_data["errors"].items():
                    if messages:
                        for msg in messages:
                            detailed_errors.append(f"[{key}] {msg}")

            return "\n".join(detailed_errors) if detailed_errors else response.text
        except Exception:
            return response.text

    def format_error(status_code, message):
        return {"error": {"code": status_code, "message": message}}

    def limpar_nans(obj):
        if isinstance(obj, float):
            if isnan(obj) or isinf(obj):
                return None
            return obj
        elif isinstance(obj, dict):
            return {k: limpar_nans(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [limpar_nans(v) for v in obj]
        else:
            return obj


    AUTH_URL = 'xxxx'
    API_URL = 'xxxx'
    client = 'xxxx'
    client_secret = 'xxxx'
    scope = 'xxxx'
    grant_type = 'xxxx'
    token = None
    

    def get_jwt_token() -> str:
        global token
        global headers
        headers = {}

        print_log(f'Iniciando a requisição do token de acesso ao COUPA!')

        param = {
            'client_id': client,
            'client_secret': client_secret,
            'scope': scope,
            'grant_type': grant_type
        }

        response = requests.post(AUTH_URL, params=param)

        if response.status_code == 200:
            token = response.json().get("access_token")
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
            print_log(f'Token gerado com sucesso!')
            return headers
        else:
            raise Exception(
                f"Failed to get token: {response.status_code} - {response.text}"
            )

    import json

    import re

    def post_data(data, num_max_tentativas=3, status_token=False):
        num_tentativas = 0
        error_mensage = ""
        code_error_status = None

        while num_tentativas < num_max_tentativas:
            num_tentativas += 1
            try:
                if not status_token:
                    headers_api = get_jwt_token()
                    status_token = True

                response = requests.post(API_URL, json=data, headers=headers_api)

                if response.status_code in [200, 201]:
                    return response.json()

                elif response.status_code == 400:
                    try:
                        status_error = response.json()
                        order_errors = status_error.get("errors", {}).get("order-header", [])
                        status_error_str = str(order_errors).replace("'", "")
                        pattern = r'\[Unable to find valid User record for ship_to_user with keys \{"login"=>"[^"]+"\}'

                        if re.match(pattern, status_error_str):
                            data["ship-to-user"]["login"] = "dummy@teste.com"
                            response = requests.post(API_URL, json=data, headers=headers_api)
                            if response.status_code in [200, 201]:
                                return response.json()
                            else:
                                code_error_status = response.status_code
                                detailed_message = extract_error_details(response)
                                raise Exception(f"API Error:\n{detailed_message}")
                        else:
                            code_error_status = response.status_code
                            detailed_message = extract_error_details(response)
                            raise Exception(f"API Error:\n{detailed_message}")
                    except Exception as e:
                        raise Exception(str(e))

                else:
                    code_error_status = response.status_code
                    detailed_message = extract_error_details(response)
                    raise Exception(f"API Error:\n{detailed_message}")

            except Exception as e:
                status_token = False
                error_mensage = e
                print_log(f'Erro no POST: {e}')
        return format_error(code_error_status, f"{error_mensage}")


    def put_data(data, id_coupa):
        API_URL_PUT = f'xxxx/{id_coupa}/issue_without_send?return_object=limited'

        if token is None:
            get_jwt_token()

        response = requests.put(API_URL_PUT, json=data, headers=headers)

        if response.status_code in [200, 201]:
            return response.json()
        elif response.status_code == 401:
            get_jwt_token()
            return format_error(response.status_code, f"API Token Access Error: {response.text}")
        else:
            return format_error(response.status_code, f"API Error: {response.text}")


    def get_data(id_po):
        API_URL_GET = f'xxxx/?po_number={id_po}'

        if token is None:
            get_jwt_token()

        response = requests.get(API_URL_GET, headers=headers)

        if response.status_code in [200, 201]:
            dict_data = response.json()
            return dict_data[0]['id']
        
        elif response.status_code == 401:
            get_jwt_token()
            return format_error(response.status_code, f"API Token Access Error: {response.text}")
        else:
            return format_error(response.status_code, f"API Error: {response.text}")
            
    usuario = os.getlogin()
    pasta_base = f"C:\\Users\\{usuario}\\OneDrive - Accenture\\Documents\\SAP\\SAP GUI\\"

    def encontrar_arquivo_recente(padrao):
        arquivos = glob.glob(padrao)
        if not arquivos:
            return None
        return max(arquivos, key=os.path.getmtime)

    arquivo_ekko = encontrar_arquivo_recente(f"{pasta_base}EXPORT_EKKO_*.xlsx")
    arquivo_ekpo = encontrar_arquivo_recente(f"{pasta_base}EXPORT_EKPO_*.xlsx")
    arquivo_usr21 = encontrar_arquivo_recente(f"{pasta_base}EXPORT_USR21_*.xlsx")
    arquivo_adr6 = encontrar_arquivo_recente(f"{pasta_base}EXPORT_ADR6_*.xlsx")
    arquivo_lfa1 = encontrar_arquivo_recente(f"{pasta_base}EXPORT_LFA1_*.xlsx")
    arquivo_me23n = encontrar_arquivo_recente(f"{pasta_base}me23n_com_texto_*.txt")
    arquivo_mm03 = encontrar_arquivo_recente(f"{pasta_base}mm03_com_texto_*.txt")
    arquivo_contrato = encontrar_arquivo_recente(f"{pasta_base}EXPORT_CONTRATO_*.xlsx")
    arquivo_eket = encontrar_arquivo_recente(f"{pasta_base}EXPORT_EKET_*.xlsx")
    arquivo_mara = encontrar_arquivo_recente(f"{pasta_base}EXPORT_MARA_*.xlsx")

    print_log("\n✔️ Arquivos encontrados com sucesso.")

    df_ekko = pd.read_excel(arquivo_ekko, header=None, skiprows=5) if arquivo_ekko else None
    df_ekpo = pd.read_excel(arquivo_ekpo, header=None, skiprows=5) if arquivo_ekpo else None
    df_usr21 = pd.read_excel(arquivo_usr21, header=None, skiprows=5) if arquivo_usr21 else None
    df_adr6 = pd.read_excel(arquivo_adr6, header=None, skiprows=5) if arquivo_adr6 else None
    df_lfa1 = pd.read_excel(arquivo_lfa1, header=None, skiprows=5) if arquivo_lfa1 else None
    df_contrato = pd.read_excel(arquivo_contrato, header=None, skiprows=5) if arquivo_contrato else None
    df_eket = pd.read_excel(arquivo_eket, header=None, skiprows=5) if arquivo_eket else None
    df_mara = pd.read_excel(arquivo_mara, header=None, skiprows=5) if arquivo_mara else None

    def converter_tipo(valor):
        if isinstance(valor, (np.int64, np.float64)):
            return int(valor) if isinstance(valor, np.int64) else float(valor)
        if pd.isna(valor):
            return ""
        return str(valor)
    

    #Carregar descrições longas de mm03 em um dicionário com padding de zeros
    descricao_longa_dict = {}
    if arquivo_mm03:
        with open(arquivo_mm03, "r", encoding="utf-8") as file:
            for linha in file:
                partes = linha.strip().split("|")
                if len(partes) >= 2:
                    codigo_material = partes[0].strip().zfill(18)
                    descricao = partes[1].strip()
                    descricao_longa_dict[codigo_material] = descricao

    po_numeros_unicos = df_ekko.iloc[:, 1].dropna().unique() if df_ekko is not None else []
    pos_com_sucesso = []
    pos_com_erro = []

    df_ekpo.iloc[:, 1] = df_ekpo.iloc[:, 1].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )

    #Padroniza a coluna 1 da EKKO também
    df_ekko.iloc[:, 1] = df_ekko.iloc[:, 1].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )

    for po_number in po_numeros_unicos:
        try:
            po_number = (
                str(int(po_number)) if isinstance(po_number, float) and po_number.is_integer()
                else str(po_number).strip()
            )
            #print(po_number)            

            # Filtrar a linha da EKKO correspondente à PO
            linha_ekko = df_ekko[df_ekko.iloc[:, 1] == po_number]
            if linha_ekko.empty:
                continue
            linha = linha_ekko.iloc[0]

            # 👉 NOVA VALIDAÇÃO: verificar se há itens na EKPO para essa PO
            linhas_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]
            if linhas_ekpo.empty:
                print_log(f"⚠️ PO {po_number} ignorada: sem itens na EKPO.")
                continue
            valor_total_bruto = converter_tipo(linha[10])
            incoterms_parte_2 = "" if pd.isna(linha[12]) else converter_tipo(linha[12])
            unidade_organizacional = converter_tipo(linha[13])
            grupo_comprador_raw = converter_tipo(linha[14])
            grupo_comprador = (
                str(grupo_comprador_raw).zfill(3)
                if str(grupo_comprador_raw).strip().isdigit()
                else str(grupo_comprador_raw).strip()
            )
            sua_referencia = converter_tipo(linha[8]) or "N/A"
            emissor_fatura_distinto = converter_tipo(linha[11])
            currency_code = converter_tipo(linha[3])

            #ship_to_user = "dummy@teste.com"

            ship_to_user = ""
            if df_usr21 is not None and df_adr6 is not None:
                usuario_ekko = linha[6]
                usr_match = df_usr21[df_usr21.iloc[:, 1] == usuario_ekko]
                if not usr_match.empty:
                    adr_match = df_adr6[df_adr6.iloc[:, 2] == usr_match.iloc[0, 2]]
                    ship_to_user = converter_tipo(adr_match.iloc[0, 8]) if not adr_match.empty else ""
                        
            location_code = ""
            if df_ekpo is not None:
                linhas_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]

                if not linhas_ekpo.empty:
                    for _, row in linhas_ekpo.iterrows():
                        line_num = converter_tipo(row[2])
                        centro_logistico = converter_tipo(row[8])  # Extraindo o valor correto
                        
                    location_code = centro_logistico 


            supplier_number = ""
            if df_lfa1 is not None:
                filtrado = df_lfa1[df_lfa1.iloc[:, 1] == linha[2]]
                
                if not filtrado.empty:
                    supplier_number = converter_tipo(filtrado.iloc[0, 2])
                    supplier_number_str = str(supplier_number).strip()

                    # Remove ".0" se for número vindo como float do Excel
                    if supplier_number_str.endswith(".0"):
                        supplier_number_str = supplier_number_str[:-2]

                    # Adiciona zeros se for só números
                    if supplier_number_str.isdigit():
                        supplier_number = supplier_number_str.zfill(14)
                    else:
                        supplier_number = supplier_number_str  # Mantém o original

            texto_cabecalho = ""
            if arquivo_me23n:
                with open(arquivo_me23n, "r", encoding="utf-8") as file:
                    for linha_txt in file:
                        if str(po_number) in linha_txt:
                            texto_cabecalho = linha_txt.split("|")[1].strip()
                            break

            # Inicializa a variável deposito
            deposito = ""

            if df_ekpo is not None:
                filtrado_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]  # Filtrando pelo PO

                if not filtrado_ekpo.empty and filtrado_ekpo.shape[1] > 17:  
                    deposito = converter_tipo(filtrado_ekpo.iloc[0, 17]) if pd.notna(filtrado_ekpo.iloc[0, 17]) else ""
                    print_log(f"Depósito extraído para PO {po_number}: {deposito}")  # Confirmação do valor
                else:
                    print_log(f"⚠ Atenção: Coluna 17 não encontrada na tabela EKPO para PO {po_number}.")



            json_data = {
                "type": "ExternalOrderHeader",
                "po-number": po_number,
                "version": 1,
                "payment-method": "invoice",
                "ship-to-attention": "",
                "ship-to-address": {"location-code": location_code},
                "ship-to-user": {
                    "login": ship_to_user
                }, 
                "supplier": {
                    "number": supplier_number
                },  
                "payment-term": {
                    "code": (
                        str(converter_tipo(linha[4])).zfill(4)
                            if pd.notna(linha[4]) and str(linha[4]).strip() != ""
                            else "Z034"
                        )
                }
                ,
                "shipping-term": {
                "code": (
                    "" if pd.isna(linha[5]) or str(linha[5]).strip() == "" 
                    else converter_tipo(linha[5])
                )
            },
                "custom-fields": {
                    "texto-de-cabecalho": converter_tipo(texto_cabecalho),
                "tipo-de-pedido": {
                        "external-ref-num": (
                            "" if pd.isna(linha[7]) or str(linha[7]).strip() == "" 
                            else converter_tipo(linha[7])
                        )
                    },
                    "valor-icms": "",
                    "valor-ipi": "",
                    "valor-icmsst": "",
                    "valor-total-bruto": valor_total_bruto,
                    "incoterms-parte-2": incoterms_parte_2,
                    "unidade-organizacional": unidade_organizacional,
                    "pedido-sap": po_number,
                    "grupo-comprador": {
                        "external-ref-num": grupo_comprador, 
                    },
                    "numero-rc": "",
                    "sua-referencia": converter_tipo(sua_referencia),
                    "empresa": "Energia",
                    "aceite-tacito": False,
                    "emissor-de-fatura-distinto": {
                        "external-ref-num": "N/A" if pd.isna(linha[11]) or linha[11] in [None, ""] else str(int(linha[11])) if isinstance(linha[11], float) and linha[11].is_integer() else str(linha[11]).strip()
                    }
                },
                "currency": {"code": currency_code},
                "order-lines": []
            }

            if df_ekpo is not None:
                linhas_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]
                for _, row in linhas_ekpo.iterrows():
                    line_num = converter_tipo(row[2])
                    description = f"{converter_tipo(row[3])} | {converter_tipo(row[4])}"
                    valor_f_line_br = row[19]
                    if isinstance(valor_f_line_br, str):
                        valor_f_line_amer = valor_f_line_br.replace('.', '').replace(',', '.')
                    else:
                        valor_f_line_amer = f"{valor_f_line_br:.2f}"
                    quantity = converter_tipo(row[5])
                    need_by_date = ""

                    if df_eket is not None:
                        eket_match = df_eket[df_eket.iloc[:, 1] == po_number]
                        if not eket_match.empty:
                            raw_date = eket_match.iloc[0, 2]
                            try:
                                parsed_date = datetime.datetime.strptime(str(raw_date), "%d.%m.%Y")
                                need_by_date = parsed_date.strftime("%Y/%m/%d")
                            except ValueError:
                                pass

                    service_type = "non_service"
                    if df_mara is not None:
                        mara_match = df_mara[df_mara.iloc[:, 1] == row[4]]
                        if mara_match.empty:
                            mara_type = "quantity_deliverable"
                        else:
                            mara_type = converter_tipo(mara_match.iloc[0, 2])
                        if mara_type in ["DIEN", "ZIEN", "ZSER", "ZSGS", "quantity_deliverable"]:
                            service_type = "quantity_deliverable"
                    
                    tipo_delinha = "Material"  
                    if df_mara is not None:
                        mara_match = df_mara[df_mara.iloc[:, 1] == row[4]]
                        if mara_match.empty:
                            mara_type = "Serviço"
                        else:
                            mara_type = converter_tipo(mara_match.iloc[0, 2]) 
                        if mara_type in ["DIEN", "ZIEN", "ZSER", "ZSGS", "Serviço"]:
                            tipo_delinha = "Serviço"

                    codigo_material = converter_tipo(row[4]).strip().zfill(18)
                    descricao_longa = descricao_longa_dict.get(codigo_material, "")

                    data_contrato, id_contrato_coupa = "", ""
                    numero_contrato = str(int(row[14])) if pd.notna(row[14]) else ""
                    if numero_contrato and df_contrato is not None:
                        contrato_match = df_contrato[df_contrato.iloc[:, 1].astype(str).str.strip() == numero_contrato]
                        if not contrato_match.empty:
                            if pd.notna(contrato_match.iloc[0, 3]):
                                try:
                                    data_contrato = datetime.datetime.strptime(str(contrato_match.iloc[0, 3]), "%d.%m.%Y").strftime("%Y/%m/%d")
                                except ValueError:
                                    pass
                            id_contrato_coupa = converter_tipo(contrato_match.iloc[0, 5]) if pd.notna(contrato_match.iloc[0, 5]) else ""

                    centro_logistico = converter_tipo(row[8])
                    utilizacao_material = str(int(float(row[9]))) if pd.notna(row[9]) else "N/A"
                    origem_material_valor = converter_tipo(row[10]) if pd.notna(row[10]) else "N/A"
                    try:
                        if origem_material_valor.replace('.', '', 1).isdigit():
                            origem_material = int(float(origem_material_valor)) if float(origem_material_valor).is_integer() else float(origem_material_valor)
                        else:
                            origem_material = origem_material_valor
                    except:
                        origem_material = origem_material_valor

                    deposito_str = "N/A"  # Define N/A por padrão

                    if df_ekpo is not None:
                        match_ekpo = df_ekpo[
                            (df_ekpo.iloc[:, 1].astype(str).str.strip() == str(po_number)) &
                            (df_ekpo.iloc[:, 2].astype(str).str.strip() == str(line_num))
                        ]
                        
                        if not match_ekpo.empty and match_ekpo.shape[1] > 17:
                            deposito_raw = match_ekpo.iloc[0, 17]
                            deposito_str = str(deposito_raw).strip() if not pd.isna(deposito_raw) and deposito_raw not in ["", None] else "N/A"
                            
                        print_log(f"Depósito extraído para PO {po_number}, linha {line_num}: {deposito_str}")


                    json_data["order-lines"].append({
                        "line-num": line_num,
                        "description": description,
                        "price": valor_f_line_amer,
                        "quantity": quantity,
                        "need-by-date": need_by_date,
                        "type": "OrderQuantityLine",
                        "custom-fields": {
                            "descricao-longa": descricao_longa,
                            "centro-logistico": {
                            "external-ref-num": centro_logistico,  
                            },
                            "tipo-da-linha": tipo_delinha,
                            "utilizacao-do-material": {
                            "external-ref-num": utilizacao_material,  
                            },
                            "origem-do-material": {
                            "external-ref-num": origem_material,  
                            },
                            "codigo-ncm": str(int(row[11])) if isinstance(row[11], float) and row[11].is_integer() else str(row[11]).strip(),
                            "codigo-do-imposto": converter_tipo(row[12]),
                            "preco-por": (
                                converter_tipo(row[18]) if pd.notna(row[18]) and str(row[18]).strip() != "" else ""
                            ),
                            "deposito": {
                                "external-ref-num": deposito_str,  
                            },
                            "data-contrato": data_contrato,
                            "numero-do-contrato": numero_contrato,
                            "item-do-contrato": str(int(row[15])) if pd.notna(row[15]) else "",
                            "id-contrato-coupa": id_contrato_coupa
                        },
                        "uom": {"code": converter_tipo(row[13])},
                        "account": {
                            "code": "Dummy",
                            "segment-1": "Dummy",
                            "account-type": {
                                "name": "COA - zzzz"
                            },
                        },
                        "currency": {"code": currency_code},
                        "commodity":  {
                            "name": converter_tipo(row[16]),  # <- AJUSTADO
                        },
                        "service-type": service_type
                    })
            json_data = limpar_nans(json_data)
            print_log(f"\n🚚 Enviando PO {po_number}...")

            # 👇 Adicione esta linha para inspecionar o payload
            print_log(json.dumps(json_data, indent=2, ensure_ascii=False))  # ensure_ascii=False para exibir acentos corretamente


            response_post = post_data(json_data)
            if isinstance(response_post, dict) and "error" in response_post:
                raise Exception(f"Erro POST: {response_post['error']['message']}")
            
            idcoupa = get_data(json_data["po-number"])

            if isinstance(idcoupa, dict) and "error" in idcoupa:
                raise Exception(f"Erro GET: {idcoupa['error']['message']}")

            response_put = put_data(json_data, idcoupa)
            if isinstance(response_put, dict) and "error" in response_put:
                raise Exception(f"Erro PUT: {response_put['error']['message']}")

            json_data["api_responses"] = {
                "post_response": response_post,
                "get_response": idcoupa,
                "put_response": response_put
            }

            pos_com_sucesso.append(json_data)
            print_log(f"✅ PO {po_number}enviada com sucesso.)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              .")

        except Exception as e:
            erro_info = {
                "po_number": po_number,
                "erro": str(e)
            }
            pos_com_erro.append(erro_info)
            print_log(f"❌ Erro ao enviar PO {po_number}: {str(e)}")

    # Salvar os resultados em arquivos JSON separados
    timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")

    # Salvar os resultados em arquivos JSON separados com timestamp
    with open(os.path.join(pasta_base, f"sucesso_{timestamp}.json"), "w", encoding="utf-8") as f:
        json.dump(pos_com_sucesso, f, ensure_ascii=False, indent=4)

    with open(os.path.join(pasta_base, f"erros_{timestamp}.json"), "w", encoding="utf-8") as f:
        json.dump(pos_com_erro, f, ensure_ascii=False, indent=4)

    print_log(f"\n📦 Total de POs enviadas com sucesso: {len(pos_com_sucesso)}")
    print_log(f"⚠️ Total de erros: {len(pos_com_erro)}")
    print_log(f"📝 Arquivo de sucesso salvo em: {os.path.join(pasta_base, 'sucesso.json')}")
    print_log(f"📝 Arquivo de erros salvo em: {os.path.join(pasta_base, 'erros.json')}")

    import time
    time.sleep(3)  # Simulação de processamento

    # Registrar o tempo de término
    end_time = datetime.datetime.now()
    end_format = end_time.strftime("%d_%m às %H_%M_%S")

    # Calcular o tempo total
    execution_time = end_time - start_time

    # Converter para formato legível (hh:mm:ss)
    execution_seconds = execution_time.total_seconds()
    hours, remainder = divmod(execution_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    print_log(f"🟢 Início: {start_format}")
    print_log(f"🔴 Finalizou: {end_format}")
    print_log(f"⏱️ Tempo total de execução: {int(hours)}h {int(minutes)}m {int(seconds)}s")

    # --- Extraindo os erros ---
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font
    from openpyxl.chart import BarChart, Reference

    # Caminho da pasta SAP GUI
    user_profile = os.environ["USERPROFILE"]
    sap_folder = os.path.join(
        user_profile,
        "OneDrive - Accenture",
        "Documents",
        "SAP",
        "SAP GUI"
    )

    # Subpasta de destino
    output_folder = os.path.join(sap_folder, "Analise")
    os.makedirs(output_folder, exist_ok=True)

    # Arquivo JSON mais recente
    json_files = glob.glob(os.path.join(sap_folder, "erros_*.json"))
    if not json_files:
        print_log("Nenhum arquivo de erro encontrado.")
        exit()

    latest_json = max(json_files, key=os.path.getctime)

    # Formatar o nome com dd_mm_HH_MM
    now = datetime.datetime.now()
    data_hora_str = now.strftime("%d_%m_%H_%M")
    excel_name = f"analise_erro_{data_hora_str}.xlsx"
    excel_path = os.path.join(output_folder, excel_name)

    # Carregar JSON
    with open(latest_json, "r", encoding="utf-8") as f:
        dados = json.load(f)

    # Lista consolidada de linhas
    linhas = []
    po_processadas = set()

    for item in dados:
        po = item.get("po_number", "")
        msg = item.get("erro", "")
        msg_lower = msg.lower()
        erro_identificado = False
        valor_extraido = ""

        if "ssl" in msg_lower or "certificate verify failed" in msg_lower:
            linhas.append({
                "Po_Number": po,
                "Erro": "SSL: CERTIFICATE_VERIFY_FAILED",
                "Valor": "",
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if 'Unable to find valid Supplier' in msg:
            msg_normalizada = msg.replace('\\u003e', '=>').replace('\u003e', '=>')
            msg_normalizada = msg_normalizada.replace('\\"', '"').replace('\\\\', '')

            match = re.search(r'number["=\s]*=>["=\s]*"([\w\d]+)', msg_normalizada)
            valor_extraido = match.group(1) if match else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "SUPPLIER",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "LookupValue record for emissor_de_fatura_distinto" in msg:
            matches = re.findall(r'external_ref_num["=\s]*=>["=\s]*"([\w\d]+)', msg)
            valor_extraido = ", ".join(matches) if matches else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "EMISSOR DE FATURA",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "LookupValue record for deposito/custom_field_4" in msg:
            matches = re.findall(r'external_ref_num["=\s]*=>["=\s]*"([\w\d]+)', msg)
            valor_extraido = ", ".join(matches) if matches else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "DEPOSITO",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "has already been taken" in msg:
            linhas.append({
                "Po_Number": po,
                "Erro": "PO DUPLICADO",
                "Valor": po,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if not erro_identificado:
            if po not in po_processadas:
                valor = ""
                if "Unable to find valid Address record for ship_to_address" in msg:
                    valor = "Verificar Status da PO"

                linhas.append({
                    "Po_Number": po,
                    "Erro": "ERRO DESCONHECIDO",
                    "Valor": valor,
                    "Mensagem_erro": msg
                })

        if "Unable to find valid Uom record for uom" in msg:
            match = re.search(r'code["=\s]*=>["=\s]*"?(\w+)"?', msg)
            valor_extraido = match.group(1) if match else ""
            linhas.append({
                "Po_Number": po,
                "Erro": "UOM INVALIDA",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })

            po_processadas.add(po)

    # Gerar DataFrame
    df_erros = pd.DataFrame(linhas)
    df_erros = df_erros[["Po_Number", "Erro", "Valor", "Mensagem_erro"]]

    # Gerar levantamento de erros com valores
    df_levantamento = df_erros.groupby(["Erro"])["Valor"].apply(lambda x: ", ".join(map(str, x.dropna().unique()))).reset_index()
    df_levantamento["Quantidade"] = df_erros["Erro"].value_counts().values

    # Salvar Excel com múltiplas abas
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_erros.to_excel(writer, sheet_name="Erros Detalhados", index=False)
        df_levantamento.to_excel(writer, sheet_name="Levantamento", index=False)

    # Abrir arquivo e adicionar gráfico
    wb = load_workbook(excel_path)
    ws = wb["Levantamento"]

    # Estilizar cabeçalhos
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Criar gráfico de barras
    chart = BarChart()
    chart.title = "Distribuição de Erros"
    chart.x_axis.title = "Tipo de Erro"
    chart.y_axis.title = "Quantidade"
    chart.style = 10

    data = Reference(ws, min_col=3, min_row=1, max_row=len(df_levantamento)+1)
    categories = Reference(ws, min_col=1, min_row=2, max_row=len(df_levantamento)+1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    ws.add_chart(chart, "E5")

    wb.save(excel_path)
    print_log(f"Arquivo salvo com gráfico e levantamento: {excel_path}")

    # --- COM SUCESSO
    # Caminho da pasta principal SAP GUI
    pasta_sap = os.path.expanduser(r"~\OneDrive - Accenture\Documents\SAP\SAP GUI")

    # Subpasta 'Analise'
    pasta_analise = os.path.join(pasta_sap, "Analise")
    os.makedirs(pasta_analise, exist_ok=True)

    # Localiza o arquivo JSON mais recente com sucesso
    arquivos_json = glob.glob(os.path.join(pasta_sap, "sucesso_*.json"))
    if not arquivos_json:
        print_log("❌ Nenhum arquivo de sucesso encontrado.")
        exit()

    arquivo_mais_recente = max(arquivos_json, key=os.path.getmtime)

    # Carrega o conteúdo do JSON
    with open(arquivo_mais_recente, encoding='utf-8') as f:
        dados = json.load(f)

    # Extrai os dados desejados corretamente
    linhas = []
    for item in dados:
        po = item.get("po-number")
        
        api_responses = item.get("api_responses", {})
        get_response = api_responses.get("get_response", None)
        
        put_response = api_responses.get("put_response", {})
        put_id = put_response.get("id") if isinstance(put_response, dict) else None

        linhas.append({
            "po-number": po,
            "get_response": get_response,
            "put_response": json.dumps(put_response, ensure_ascii=False),
            "id": put_id
        })

    # Cria o DataFrame
    df = pd.DataFrame(linhas)

    # Nome dinâmico com data/hora
    agora = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"PO_SUCESSO_{agora}.xlsx"
    caminho_saida = os.path.join(pasta_analise, nome_arquivo)

    # Salva no Excel
    df.to_excel(caminho_saida, index=False)

    print_log(f"✅ Arquivo gerado com sucesso: {caminho_saida}")

    print_log(f"✅ Processo de envio a Api finalizado")

# ======== SUBIR ATUALIZAÇÕES DA PO =============

def upload_atualizacoes():
    # Caminho do script (mesma pasta que este arquivo)
    caminho_script = os.path.join(os.path.dirname(__file__), "encontrarpo.py")

    # Executar o script
    subprocess.run(["python", caminho_script], check=True)

#================== CANCELA LINHA ========================
    AUTH_URL = 'xxxx'
    API_URL = 'xxxx'
    API_URL_BASE = 'xxxx'
    client = 'xxxx'
    client_secret = 'xxxx'
    scope = 'xxxx'
    grant_type = 'xxxx'

    token = None
    headers = {}


    # --- CONFIGURAÇÕES ---
    usuario = getpass.getuser()
    pasta = Path(f"C:/Users/{usuario}/OneDrive - Accenture/Documents/SAP/SAP GUI")

    # --- GERAR TOKEN ---
    def get_jwt_token():
        global token, headers
        print_log_sap("🔐 Solicitando token de acesso ao Coupa...")
        param = {
            'client_id': client,
            'client_secret': client_secret,
            'scope': scope,
            'grant_type': grant_type
        }
        response = requests.post(AUTH_URL, params=param)
        if response.status_code == 200:
            token = response.json().get("access_token")
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
            print_log_sap("✅ Token obtido com sucesso.")
        else:
            raise Exception(f"❌ Erro ao obter token: {response.status_code} - {response.text}")

    # --- LOCALIZAR ARQUIVO COMPARATIVO MAIS RECENTE ---
    def localizar_arquivo_comparativo():
        arquivos = list(pasta.glob("Comparativo_*.xlsx"))
        if not arquivos:
            print_log_sap("❌ Nenhum arquivo Comparativo encontrado.")
            return None
        return max(arquivos, key=os.path.getctime)

    # --- CANCELAR LINHA INDIVIDUAL ---
    def cancelar_linha(po_id, line_id):
        cancel_url = f"{API_URL_BASE}/{po_id}?return_object=limited"
        payload = {
            "order-lines": [
                {
                    "id": int(line_id),
                    "_delete": "true"
                }
            ]
        }

        response = requests.put(cancel_url, headers=headers, json=payload)
        if response.status_code == 200:
            print_log_sap(f"✅ Linha {line_id} da PO {po_id} cancelada com sucesso.")
        else:
            print_log_sap(f"❌ Falha ao cancelar linha {line_id} da PO {po_id}.")
            print_log_sap(f"Status: {response.status_code} - {response.text}")

    # --- EXECUÇÃO PRINCIPAL ---
    def main():
        try:
            get_jwt_token()

            arquivo = localizar_arquivo_comparativo()
            if not arquivo:
                return

            df = pd.read_excel(arquivo)
            print_log_sap(f"📄 Comparativo carregado: {arquivo.name}")

            print_log_sap("📋 Conteúdo do arquivo lido:")
            print_log_sap(df.to_string(index=False))

            # Salvar conteúdo em arquivo de log
            log_path = pasta / "log_conteudo_excel.txt"
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(f"Conteúdo do arquivo: {arquivo.name}\n\n")
                f.write(df.to_string(index=False))

            print_log_sap(f"📝 Conteúdo salvo em: {log_path}")

            # Filtrar linhas com cod_elim = L ou S e com Line_id válido
            df_filtrado = df[
                df["cod_elim"].isin(["L", "S"]) &
                (df["Line_id"] != "não consta no Coupa") &
                (df["Line_id"].notna())
            ]

            if df_filtrado.empty:
                print_log_sap("⚠️ Nenhuma linha elegível para cancelamento.")
                return

            # --- COMENTADO: ENVIO PARA A API ---
            for _, row in df_filtrado.iterrows():
                po_id = row["PO_id"]
                line_id = row["Line_id"]
                cancelar_linha(po_id, line_id)

            
        except Exception as e:
            print_log_sap(f"❌ Erro geral: {e}")

    if __name__ == "__main__":
        main()    
#================== ATUALIZA PO ========================
    start_time = datetime.datetime.now()
    start_format = start_time.strftime("%d_%m às %H_%M_%S")

    pasta_base = os.path.join(os.path.expanduser("~"), r"OneDrive - Accenture\Documents\SAP\SAP GUI")
    print_log_sap("Arquivos encontrados para Comparativo*.xlsx:")
    print_log_sap(glob.glob(os.path.join(pasta_base, "Comparativo*.xlsx")))

    # Timestamp para nomear o arquivo
    timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")

    log_file = os.path.join(pasta_base, f"log_api_{timestamp}.txt")

    import logging

    log_file = os.path.join(pasta_base, f"log_api_{timestamp}.txt")
    logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.INFO, format='%(message)s')

    def print_log(msg):
        print(msg)
        logging.info(msg)


    # Registrar o tempo de início
    start_time = datetime.datetime.now()
    start_format = start_time.strftime("%d_%m às %H_%M_%S")

    def extract_error_details(response):
        try:
            error_data = response.json()
            detailed_errors = []

            if "errors" in error_data:
                for key, messages in error_data["errors"].items():
                    if messages:
                        for msg in messages:
                            detailed_errors.append(f"[{key}] {msg}")

            return "\n".join(detailed_errors) if detailed_errors else response.text
        except Exception:
            return response.text

    def format_error(status_code, message):
        return {"error": {"code": status_code, "message": message}}

    def limpar_nans(obj):
        if isinstance(obj, float):
            if isnan(obj) or isinf(obj):
                return None
            return obj
        elif isinstance(obj, dict):
            return {k: limpar_nans(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [limpar_nans(v) for v in obj]
        else:
            return obj


    AUTH_URL = 'xxxx'
    API_URL = 'xxxx'
    client = 'xxxx'
    client_secret = 'xxxx'
    scope = 'xxxx'
    grant_type = 'xxxx'
    token = None


    def get_jwt_token() -> str:
        global token
        global headers
        headers = {}

        print_log(f'Iniciando a requisição do token de acesso ao COUPA!')

        param = {
            'client_id': client,
            'client_secret': client_secret,
            'scope': scope,
            'grant_type': grant_type
        }

        response = requests.post(AUTH_URL, params=param)

        if response.status_code == 200:
            token = response.json().get("access_token")
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
            print_log(f'Token gerado com sucesso!')
            return headers
        else:
            raise Exception(
                f"Failed to get token: {response.status_code} - {response.text}"
            )

    import json

    import re

    def put_data(data, id_coupa):
        API_URL_PUT = f'xxxx/{id_coupa}/issue_without_send?return_object=limited'

        if token is None:
            get_jwt_token()

        response = requests.put(API_URL_PUT, json=data, headers=headers)

        if response.status_code in [200, 201]:
            return response.json()
        elif response.status_code == 401:
            get_jwt_token()
            return format_error(response.status_code, f"API Token Access Error: {response.text}")
        else:
            return format_error(response.status_code, f"API Error: {response.text}")


    def encontrar_arquivo_recente(padrao):
        arquivos = glob.glob(padrao)
        print_log(f"🔍 Arquivos encontrados: {arquivos}")
        if not arquivos:
            return None
        return max(arquivos, key=os.path.getmtime)

    # Caminho base
    usuario = os.getlogin()
    pasta_base = f"C:\\Users\\{usuario}\\OneDrive - Accenture\\Documents\\SAP\\SAP GUI\\"
    padrao_comparativo = os.path.join(pasta_base, "Comparativo_*.xlsx")
    arquivo_comparativo = encontrar_arquivo_recente(padrao_comparativo)

    # Verificação final
    print_log(f"🧪 Caminho final do Comparativo: {arquivo_comparativo}")
    print_log(f"📚 Tipo de dado recebido: {type(arquivo_comparativo)}")

    if arquivo_comparativo and isinstance(arquivo_comparativo, str):
        df_comparativo = pd.read_excel(arquivo_comparativo)
        print_log("✅ Comparativo carregado com sucesso.")
    else:
        print_log("❌ Falha ao localizar o arquivo de Comparativo.")
        exit()


    arquivo_ekko = encontrar_arquivo_recente(f"{pasta_base}CDHDR_EKKO_*.xlsx")
    arquivo_ekpo = encontrar_arquivo_recente(f"{pasta_base}CDHDR_EKPO_*.xlsx")
    arquivo_usr21 = encontrar_arquivo_recente(f"{pasta_base}CDHDR_USR21_*.xlsx")
    arquivo_adr6 = encontrar_arquivo_recente(f"{pasta_base}CDHDR_ADR6_*.xlsx")
    arquivo_lfa1 = encontrar_arquivo_recente(f"{pasta_base}CDHDR_LFA1_*.xlsx")
    arquivo_me23n = encontrar_arquivo_recente(f"{pasta_base}CDHDR_me23n_com_texto_*.txt")
    arquivo_mm03 = encontrar_arquivo_recente(f"{pasta_base}CDHDR_mm03_com_texto_*.txt")
    arquivo_contrato = encontrar_arquivo_recente(f"{pasta_base}CDHDR_CONTRATO_*.xlsx")
    arquivo_eket = encontrar_arquivo_recente(f"{pasta_base}CDHDR_EKET_*.xlsx")
    arquivo_mara = encontrar_arquivo_recente(f"{pasta_base}CDHDR_MARA_*.xlsx")

    print_log("\n✔️ Arquivos encontrados com sucesso.")

    df_ekko = pd.read_excel(arquivo_ekko, header=None, skiprows=5) if arquivo_ekko else None
    df_ekpo = pd.read_excel(arquivo_ekpo, header=None, skiprows=5) if arquivo_ekpo else None
    df_usr21 = pd.read_excel(arquivo_usr21, header=None, skiprows=5) if arquivo_usr21 else None
    df_adr6 = pd.read_excel(arquivo_adr6, header=None, skiprows=5) if arquivo_adr6 else None
    df_lfa1 = pd.read_excel(arquivo_lfa1, header=None, skiprows=5) if arquivo_lfa1 else None
    df_contrato = pd.read_excel(arquivo_contrato, header=None, skiprows=5) if arquivo_contrato else None
    df_eket = pd.read_excel(arquivo_eket, header=None, skiprows=5) if arquivo_eket else None
    df_mara = pd.read_excel(arquivo_mara, header=None, skiprows=5) if arquivo_mara else None

    # Normalizar as colunas do Comparativo
    df_comparativo["PO_number"] = df_comparativo["PO_number"].astype(str).str.strip()
    df_comparativo["PO_id"] = df_comparativo["PO_id"].astype(str).str.strip()
    df_comparativo["cod_elim"] = df_comparativo["cod_elim"].astype(str).str.strip().str.upper()

    # Obter POs da EKKO (coluna B, índice 1), a partir da linha 6 (índice 5)
    po_ekko = df_ekko.iloc[5:, 1].dropna().astype(str).str.strip().unique()

    # Aplicar as regras do filtro
    df_filtrado = df_comparativo[
        df_comparativo["PO_id"].notna() &
        (df_comparativo["PO_id"].astype(str).str.strip() != "") &
        (df_comparativo["PO_id"].str.lower() != "não consta no coupa") &
        df_comparativo["Line_id"].notna() &
        (df_comparativo["Line_id"].astype(str).str.strip() != "") &
        df_comparativo["cod_elim"].notna() &
        (df_comparativo["cod_elim"].astype(str).str.strip() != "") &
        (~df_comparativo["cod_elim"].astype(str).str.upper().isin(["L", "S"]))
    ]


    # Obter a lista final de POs válidas
    po_ids_unicos = df_filtrado["PO_id"].astype(str).str.strip().unique()


    def converter_tipo(valor):
        if isinstance(valor, (np.int64, np.float64)):
            return int(valor) if isinstance(valor, np.int64) else float(valor)
        if pd.isna(valor):
            return ""
        return str(valor)


    #Carregar descrições longas de mm03 em um dicionário com padding de zeros
    descricao_longa_dict = {}
    if arquivo_mm03:
        with open(arquivo_mm03, "r", encoding="utf-8") as file:
            for linha in file:
                partes = linha.strip().split("|")
                if len(partes) >= 2:
                    codigo_material = partes[0].strip().zfill(18)
                    descricao = partes[1].strip()
                    descricao_longa_dict[codigo_material] = descricao

    po_numeros_unicos = df_ekko.iloc[:, 1].dropna().unique() if df_ekko is not None else []
    pos_com_sucesso = []
    pos_com_erro = []

    df_ekpo.iloc[:, 1] = df_ekpo.iloc[:, 1].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )

    #Padroniza a coluna 1 da EKKO também
    df_ekko.iloc[:, 1] = df_ekko.iloc[:, 1].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )

    po_ids_unicos = df_filtrado["PO_id"].astype(str).str.strip().unique()

    for po_id in po_ids_unicos:  
        try:
            # Filtrar as linhas do df_filtrado para esse po_id
            linhas_po = df_filtrado[df_filtrado["PO_id"].astype(str).str.strip() == po_id]
            idcoupa = str(linhas_po.iloc[0]["PO_id"]).strip()

            # Pega o po_number da primeira linha (assumo que seja o mesmo para todas linhas do mesmo po_id)
            po_number = linhas_po.iloc[0]["PO_number"]
            po_number = (
                str(int(po_number)) if isinstance(po_number, float) and po_number.is_integer()
                else str(po_number).strip()
            )

            # Buscar a linha correspondente na EKKO pelo po_number
            linha_ekko = df_ekko[df_ekko.iloc[:, 1] == po_number]
            if linha_ekko.empty:
                print_log(f"⚠️ PO {po_number} (PO_ID {po_id}) ignorada: não encontrada na EKKO.")
                continue
            linha = linha_ekko.iloc[0]
            valor_total_bruto = converter_tipo(linha[10])
            incoterms_parte_2 = "" if pd.isna(linha[12]) else converter_tipo(linha[12])
            unidade_organizacional = converter_tipo(linha[13])
            grupo_comprador_raw = converter_tipo(linha[14])
            grupo_comprador = (
                str(grupo_comprador_raw).zfill(3)
                if str(grupo_comprador_raw).strip().isdigit()
                else str(grupo_comprador_raw).strip()
            )
            sua_referencia = converter_tipo(linha[8]) or "N/A"
            emissor_fatura_distinto = converter_tipo(linha[11])
            currency_code = converter_tipo(linha[3])

            #ship_to_user = "dummy@teste.com"

            ship_to_user = ""
            if df_usr21 is not None and df_adr6 is not None:
                usuario_ekko = linha[6]
                usr_match = df_usr21[df_usr21.iloc[:, 1] == usuario_ekko]
                if not usr_match.empty:
                    adr_match = df_adr6[df_adr6.iloc[:, 2] == usr_match.iloc[0, 2]]
                    ship_to_user = converter_tipo(adr_match.iloc[0, 8]) if not adr_match.empty else ""
                        
            location_code = ""
            if df_ekpo is not None:
                linhas_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]

                if not linhas_ekpo.empty:
                    for _, row in linhas_ekpo.iterrows():
                        line_num = converter_tipo(row[2])
                        centro_logistico = converter_tipo(row[8])  # Extraindo o valor correto
                        
                    location_code = centro_logistico 


            supplier_number = ""
            if df_lfa1 is not None:
                filtrado = df_lfa1[df_lfa1.iloc[:, 1] == linha[2]]
                
                if not filtrado.empty:
                    supplier_number = converter_tipo(filtrado.iloc[0, 2])
                    supplier_number_str = str(supplier_number).strip()

                    # Remove ".0" se for número vindo como float do Excel
                    if supplier_number_str.endswith(".0"):
                        supplier_number_str = supplier_number_str[:-2]

                    # Adiciona zeros se for só números
                    if supplier_number_str.isdigit():
                        supplier_number = supplier_number_str.zfill(14)
                    else:
                        supplier_number = supplier_number_str  # Mantém o original

            texto_cabecalho = ""
            if arquivo_me23n:
                with open(arquivo_me23n, "r", encoding="utf-8") as file:
                    for linha_txt in file:
                        if str(po_number) in linha_txt:
                            texto_cabecalho = linha_txt.split("|")[1].strip()
                            break

            # Inicializa a variável deposito
            deposito = ""

            if df_ekpo is not None:
                filtrado_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]  # Filtrando pelo PO

                if not filtrado_ekpo.empty and filtrado_ekpo.shape[1] > 17:  
                    deposito = converter_tipo(filtrado_ekpo.iloc[0, 17]) if pd.notna(filtrado_ekpo.iloc[0, 17]) else ""
                    print_log(f"Depósito extraído para PO {po_number}: {deposito}")  # Confirmação do valor
                else:
                    print_log(f"⚠ Atenção: Coluna 17 não encontrada na tabela EKPO para PO {po_number}.")



            json_data = {
                "type": "ExternalOrderHeader",
                "po-number": po_number,
                "version": 1,
                "payment-method": "invoice",
                "ship-to-attention": "",
                "ship-to-address": {"location-code": location_code},
                "ship-to-user": {
                    "login": ship_to_user
                }, 
                "supplier": {
                    "number": supplier_number
                },  
                "payment-term": {
                    "code": (
                        str(converter_tipo(linha[4])).zfill(4)
                            if pd.notna(linha[4]) and str(linha[4]).strip() != ""
                            else "Z034"
                        )
                }
                ,
                "shipping-term": {
                "code": (
                    "" if pd.isna(linha[5]) or str(linha[5]).strip() == "" 
                    else converter_tipo(linha[5])
                )
            },
                "custom-fields": {
                    "texto-de-cabecalho": converter_tipo(texto_cabecalho),
                "tipo-de-pedido": {
                        "external-ref-num": (
                            "" if pd.isna(linha[7]) or str(linha[7]).strip() == "" 
                            else converter_tipo(linha[7])
                        )
                    },
                    "valor-icms": "",
                    "valor-ipi": "",
                    "valor-icmsst": "",
                    "valor-total-bruto": valor_total_bruto,
                    "incoterms-parte-2": incoterms_parte_2,
                    "unidade-organizacional": unidade_organizacional,
                    "pedido-sap": po_number,
                    "grupo-comprador": {
                        "external-ref-num": grupo_comprador, 
                    },
                    "numero-rc": "",
                    "sua-referencia": converter_tipo(sua_referencia),
                    "empresa": "Energia",
                    "aceite-tacito": False,
                    "emissor-de-fatura-distinto": {
                        "external-ref-num": "N/A" if pd.isna(linha[11]) or linha[11] in [None, ""] else str(int(linha[11])) if isinstance(linha[11], float) and linha[11].is_integer() else str(linha[11]).strip()
                    }
                },
                "currency": {"code": currency_code},
                "order-lines": []
            }

            if df_ekpo is not None:
                linhas_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]
                for _, row in linhas_ekpo.iterrows():
                    line_num = converter_tipo(row[2])
                    description = f"{converter_tipo(row[3])} | {converter_tipo(row[4])}"
                    valor_f_line_br = row[21]
                    if isinstance(valor_f_line_br, str):
                        valor_f_line_amer = valor_f_line_br.replace('.', '').replace(',', '.')
                    else:
                        valor_f_line_amer = f"{valor_f_line_br:.2f}"
                    quantity = converter_tipo(row[5])
                    need_by_date = ""

                    if df_eket is not None:
                        eket_match = df_eket[df_eket.iloc[:, 1] == po_number]
                        if not eket_match.empty:
                            raw_date = eket_match.iloc[0, 2]
                            try:
                                parsed_date = datetime.datetime.strptime(str(raw_date), "%d.%m.%Y")
                                need_by_date = parsed_date.strftime("%Y/%m/%d")
                            except ValueError:
                                pass

                    service_type = "non_service"
                    if df_mara is not None:
                        mara_match = df_mara[df_mara.iloc[:, 1] == row[4]]
                        if mara_match.empty:
                            mara_type = "quantity_deliverable"
                        else:
                            mara_type = converter_tipo(mara_match.iloc[0, 2])
                        if mara_type in ["DIEN", "ZIEN", "ZSER", "ZSGS", "quantity_deliverable"]:
                            service_type = "quantity_deliverable"
                    
                    tipo_delinha = "Material"  
                    if df_mara is not None:
                        mara_match = df_mara[df_mara.iloc[:, 1] == row[4]]
                        if mara_match.empty:
                            mara_type = "Serviço"
                        else:
                            mara_type = converter_tipo(mara_match.iloc[0, 2]) 
                        if mara_type in ["DIEN", "ZIEN", "ZSER", "ZSGS", "Serviço"]:
                            tipo_delinha = "Serviço"

                    codigo_material = converter_tipo(row[4]).strip().zfill(18)
                    descricao_longa = descricao_longa_dict.get(codigo_material, "")

                    data_contrato, id_contrato_coupa = "", ""
                    numero_contrato = str(int(row[14])) if pd.notna(row[14]) else ""
                    if numero_contrato and df_contrato is not None:
                        contrato_match = df_contrato[df_contrato.iloc[:, 1].astype(str).str.strip() == numero_contrato]
                        if not contrato_match.empty:
                            if pd.notna(contrato_match.iloc[0, 3]):
                                try:
                                    data_contrato = datetime.datetime.strptime(str(contrato_match.iloc[0, 3]), "%d.%m.%Y").strftime("%Y/%m/%d")
                                except ValueError:
                                    pass
                            id_contrato_coupa = converter_tipo(contrato_match.iloc[0, 5]) if pd.notna(contrato_match.iloc[0, 5]) else ""

                    centro_logistico = converter_tipo(row[8])
                    utilizacao_material = str(int(float(row[9]))) if pd.notna(row[9]) else "N/A"
                    origem_material_valor = converter_tipo(row[10]) if pd.notna(row[10]) else "N/A"
                    try:
                        if origem_material_valor.replace('.', '', 1).isdigit():
                            origem_material = int(float(origem_material_valor)) if float(origem_material_valor).is_integer() else float(origem_material_valor)
                        else:
                            origem_material = origem_material_valor
                    except:
                        origem_material = origem_material_valor

                    deposito_str = "N/A"  # Define N/A por padrão

                    if df_ekpo is not None:
                        match_ekpo = df_ekpo[
                            (df_ekpo.iloc[:, 1].astype(str).str.strip() == str(po_number)) &
                            (df_ekpo.iloc[:, 2].astype(str).str.strip() == str(line_num))
                        ]
                        
                        if not match_ekpo.empty and match_ekpo.shape[1] > 17:
                            deposito_raw = match_ekpo.iloc[0, 17]
                            deposito_str = str(deposito_raw).strip() if not pd.isna(deposito_raw) and deposito_raw not in ["", None] else "N/A"
                            
                        print_log(f"Depósito extraído para PO {po_number}, linha {line_num}: {deposito_str}")


                    json_data["order-lines"].append({
                        "line-num": line_num,
                        "description": description,
                        "price": valor_f_line_amer,
                        "quantity": quantity,
                        "need-by-date": need_by_date,
                        "type": "OrderQuantityLine",
                        "custom-fields": {
                            "descricao-longa": descricao_longa,
                            "centro-logistico": {
                            "external-ref-num": centro_logistico,  
                            },
                            "tipo-da-linha": tipo_delinha,
                            "utilizacao-do-material": {
                            "external-ref-num": utilizacao_material,  
                            },
                            "origem-do-material": {
                            "external-ref-num": origem_material,  
                            },
                            "codigo-ncm": str(int(row[11])) if isinstance(row[11], float) and row[11].is_integer() else str(row[11]).strip(),
                            "codigo-do-imposto": converter_tipo(row[12]),
                            "preco-por": (
                                converter_tipo(row[18]) if pd.notna(row[18]) and str(row[18]).strip() != "" else ""
                            ),
                            "deposito": {
                                "external-ref-num": deposito_str,  
                            },
                            "data-contrato": data_contrato,
                            "numero-do-contrato": numero_contrato,
                            "item-do-contrato": str(int(row[15])) if pd.notna(row[15]) else "",
                            "id-contrato-coupa": id_contrato_coupa
                        },
                        "uom": {"code": converter_tipo(row[13])},
                        "account": {
                            "code": "Dummy",
                            "segment-1": "Dummy",
                            "account-type": {
                                "name": "COA - zzzz"
                            },
                        },
                        "currency": {"code": currency_code},
                        "commodity":  {
                            "name": converter_tipo(row[16]),  # <- AJUSTADO
                        },
                        "service-type": service_type
                    })
            json_data = limpar_nans(json_data)
            print_log(f"\n🚚 Enviando PO {po_number}...")

            # 👇 Adicione esta linha para inspecionar o payload
            print_log(json.dumps(json_data, indent=2, ensure_ascii=False))  # ensure_ascii=False para exibir acentos corretamente


            #response_post = post_data(json_data)
            #if isinstance(response_post, dict) and "error" in response_post:
            #   raise Exception(f"Erro POST: {response_post['error']['message']}")
            
            '''idcoupa = get_data(json_data["po-number"])

            if isinstance(idcoupa, dict) and "error" in idcoupa:
                raise Exception(f"Erro GET: {idcoupa['error']['message']}")'''

            response_put = put_data(json_data, idcoupa)
            if isinstance(response_put, dict) and "error" in response_put:
                raise Exception(f"Erro PUT: {response_put['error']['message']}")

            json_data["api_responses"] = {
                #"post_response": response_post,
                #"get_response": idcoupa,
                "put_response": response_put
            }

            pos_com_sucesso.append(json_data)
            print_log(f"✅ PO {po_number} enviada com sucesso.)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              .")

        except Exception as e:
            erro_info = {
                "po_number": po_number,
                "erro": str(e)
            }
            pos_com_erro.append(erro_info)
            print_log(f"❌ Erro ao enviar PO {po_number}: {str(e)}")

    # Salvar os resultados em arquivos JSON separados
    timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")

    # Salvar os resultados em arquivos JSON separados com timestamp
    with open(os.path.join(pasta_base, f"sucesso_{timestamp}.json"), "w", encoding="utf-8") as f:
        json.dump(pos_com_sucesso, f, ensure_ascii=False, indent=4)

    with open(os.path.join(pasta_base, f"erros_{timestamp}.json"), "w", encoding="utf-8") as f:
        json.dump(pos_com_erro, f, ensure_ascii=False, indent=4)

    print_log(f"\n📦 Total de POs enviadas com sucesso: {len(pos_com_sucesso)}")
    print_log(f"⚠️ Total de erros: {len(pos_com_erro)}")
    print_log(f"📝 Arquivo de sucesso salvo em: {os.path.join(pasta_base, 'sucesso.json')}")
    print_log(f"📝 Arquivo de erros salvo em: {os.path.join(pasta_base, 'erros.json')}")

    import time
    time.sleep(3)  # Simulação de processamento

    # Registrar o tempo de término
    end_time = datetime.datetime.now()
    end_format = end_time.strftime("%d_%m às %H_%M_%S")

    # Calcular o tempo total
    execution_time = end_time - start_time

    # Converter para formato legível (hh:mm:ss)
    execution_seconds = execution_time.total_seconds()
    hours, remainder = divmod(execution_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    print_log(f"🟢 Início: {start_format}")
    print_log(f"🔴 Finalizou: {end_format}")
    print_log(f"⏱️ Tempo total de execução: {int(hours)}h {int(minutes)}m {int(seconds)}s")

    # --- Extraindo os erros ---
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font
    from openpyxl.chart import BarChart, Reference

    # Caminho da pasta SAP GUI
    user_profile = os.environ["USERPROFILE"]
    sap_folder = os.path.join(
        user_profile,
        "OneDrive - Accenture",
        "Documents",
        "SAP",
        "SAP GUI"
    )

    # Subpasta de destino
    output_folder = os.path.join(sap_folder, "Analise")
    os.makedirs(output_folder, exist_ok=True)

    # Arquivo JSON mais recente
    json_files = glob.glob(os.path.join(sap_folder, "erros_*.json"))
    if not json_files:
        print_log("Nenhum arquivo de erro encontrado.")
        exit()

    latest_json = max(json_files, key=os.path.getctime)

    # Formatar o nome com dd_mm_HH_MM
    now = datetime.datetime.now()
    data_hora_str = now.strftime("%d_%m_%H_%M")
    excel_name = f"analise_erro_{data_hora_str}.xlsx"
    excel_path = os.path.join(output_folder, excel_name)

    # Carregar JSON
    with open(latest_json, "r", encoding="utf-8") as f:
        dados = json.load(f)

    # Lista consolidada de linhas
    linhas = []
    po_processadas = set()

    for item in dados:
        po = item.get("po_number", "")
        msg = item.get("erro", "")
        msg_lower = msg.lower()
        erro_identificado = False
        valor_extraido = ""

        if "ssl" in msg_lower or "certificate verify failed" in msg_lower:
            linhas.append({
                "Po_Number": po,
                "Erro": "SSL: CERTIFICATE_VERIFY_FAILED",
                "Valor": "",
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if 'Unable to find valid Supplier' in msg:
            msg_normalizada = msg.replace('\\u003e', '=>').replace('\u003e', '=>')
            msg_normalizada = msg_normalizada.replace('\\"', '"').replace('\\\\', '')

            match = re.search(r'number["=\s]*=>["=\s]*"([\w\d]+)', msg_normalizada)
            valor_extraido = match.group(1) if match else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "SUPPLIER",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "LookupValue record for emissor_de_fatura_distinto" in msg:
            matches = re.findall(r'external_ref_num["=\s]*=>["=\s]*"([\w\d]+)', msg)
            valor_extraido = ", ".join(matches) if matches else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "EMISSOR DE FATURA",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "LookupValue record for deposito/custom_field_4" in msg:
            matches = re.findall(r'external_ref_num["=\s]*=>["=\s]*"([\w\d]+)', msg)
            valor_extraido = ", ".join(matches) if matches else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "DEPOSITO",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "has already been taken" in msg:
            linhas.append({
                "Po_Number": po,
                "Erro": "PO DUPLICADO",
                "Valor": po,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if not erro_identificado:
            if po not in po_processadas:
                valor = ""
                if "Unable to find valid Address record for ship_to_address" in msg:
                    valor = "Verificar Status da PO"

                linhas.append({
                    "Po_Number": po,
                    "Erro": "ERRO DESCONHECIDO",
                    "Valor": valor,
                    "Mensagem_erro": msg
                })

        if "Unable to find valid Uom record for uom" in msg:
            match = re.search(r'code["=\s]*=>["=\s]*"?(\w+)"?', msg)
            valor_extraido = match.group(1) if match else ""
            linhas.append({
                "Po_Number": po,
                "Erro": "UOM INVALIDA",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })

            po_processadas.add(po)

    # Gerar DataFrame
    df_erros = pd.DataFrame(linhas)
    df_erros = df_erros[["Po_Number", "Erro", "Valor", "Mensagem_erro"]]

    # Gerar levantamento de erros com valores
    df_levantamento = df_erros.groupby(["Erro"])["Valor"].apply(lambda x: ", ".join(map(str, x.dropna().unique()))).reset_index()
    df_levantamento["Quantidade"] = df_erros["Erro"].value_counts().values

    # Salvar Excel com múltiplas abas
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_erros.to_excel(writer, sheet_name="Erros Detalhados", index=False)
        df_levantamento.to_excel(writer, sheet_name="Levantamento", index=False)

    # Abrir arquivo e adicionar gráfico
    wb = load_workbook(excel_path)
    ws = wb["Levantamento"]

    # Estilizar cabeçalhos
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Criar gráfico de barras
    chart = BarChart()
    chart.title = "Distribuição de Erros"
    chart.x_axis.title = "Tipo de Erro"
    chart.y_axis.title = "Quantidade"
    chart.style = 10

    data = Reference(ws, min_col=3, min_row=1, max_row=len(df_levantamento)+1)
    categories = Reference(ws, min_col=1, min_row=2, max_row=len(df_levantamento)+1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    ws.add_chart(chart, "E5")

    wb.save(excel_path)
    print_log(f"Arquivo salvo com gráfico e levantamento: {excel_path}")

    # --- COM SUCESSO
    # Caminho da pasta principal SAP GUI
    pasta_sap = os.path.expanduser(r"~\OneDrive - Accenture\Documents\SAP\SAP GUI")

    # Subpasta 'Analise'
    pasta_analise = os.path.join(pasta_sap, "Analise")
    os.makedirs(pasta_analise, exist_ok=True)

    # Localiza o arquivo JSON mais recente com sucesso
    arquivos_json = glob.glob(os.path.join(pasta_sap, "sucesso_*.json"))
    if not arquivos_json:
        print_log("❌ Nenhum arquivo de sucesso encontrado.")
        exit()

    arquivo_mais_recente = max(arquivos_json, key=os.path.getmtime)

    # Carrega o conteúdo do JSON
    with open(arquivo_mais_recente, encoding='utf-8') as f:
        dados = json.load(f)

    # Extrai os dados desejados corretamente
    linhas = []
    for item in dados:
        po = item.get("po-number")
        
        api_responses = item.get("api_responses", {})
        get_response = api_responses.get("get_response", None)
        
        put_response = api_responses.get("put_response", {})
        put_id = put_response.get("id") if isinstance(put_response, dict) else None

        linhas.append({
            "po-number": po,
            "get_response": get_response,
            "put_response": json.dumps(put_response, ensure_ascii=False),
            "id": put_id
        })

    # Cria o DataFrame
    df = pd.DataFrame(linhas)

    # Nome dinâmico com data/hora
    agora = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"PO_SUCESSO_{agora}.xlsx"
    caminho_saida = os.path.join(pasta_analise, nome_arquivo)

    # Salva no Excel
    df.to_excel(caminho_saida, index=False)

    print_log(f"✅ Arquivo gerado com sucesso: {caminho_saida}")

    print_log(f"✅ Processo de envio a Api finalizado")

# ======== CRIA  NOVA PO =============

    # --- Caminho da pasta SAP ---
    pasta_sap = os.path.expanduser(r"~\OneDrive - Accenture\Documents\SAP\SAP GUI")
    if not os.path.exists(pasta_sap):
        os.makedirs(pasta_sap)

    # --- Localiza o arquivo EKPO mais recente ---
    arquivos_ekpo = [f for f in os.listdir(pasta_sap) if f.startswith("CDHDR_EKPO_") and f.endswith(".xlsx")]
    arquivos_ekpo.sort(key=lambda x: os.path.getmtime(os.path.join(pasta_sap, x)), reverse=True)
    arquivo_ekpo = os.path.join(pasta_sap, arquivos_ekpo[0]) if arquivos_ekpo else None

    # --- Caminho do arquivo GrupoMercadoria ---
    arquivo_grupo = os.path.join(pasta_sap, "GrupoMercadoria.xlsx")

    # --- Verificação de existência dos arquivos ---
    if not arquivo_ekpo or not os.path.exists(arquivo_grupo):
        print_log("Arquivo EKPO ou GrupoMercadoria não encontrado.")
        exit()

    # --- Carrega os dados do GrupoMercadoria ---
    wb_grupo = openpyxl.load_workbook(arquivo_grupo)
    ws_grupo = wb_grupo.active

    # Cria dicionário: {codigo_grupo: nome_grupo}
    mapa_grupo = {}
    for row in ws_grupo.iter_rows(min_row=2, values_only=True):  # assumindo cabeçalho
        if row[0] is not None and row[1] is not None and row[2] is not None:
            mapa_grupo[str(row[0]).strip()] = str(row[2]).strip()

    # --- Carrega o arquivo EKPO ---
    wb_ekpo = openpyxl.load_workbook(arquivo_ekpo)
    ws_ekpo = wb_ekpo.active

    # --- Substitui valores da coluna Q (17ª) a partir da linha 6 ---
    for row in ws_ekpo.iter_rows(min_row=6, min_col=17, max_col=17):
        cell = row[0]
        valor_original = str(cell.value).strip() if cell.value is not None else ""
        if valor_original in mapa_grupo:
            cell.value = mapa_grupo[valor_original]

    # --- Salva com nome dinâmico baseado na data e hora ---
    agora = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")
    nome_arquivo = f"CDHDR_EKPO_{agora}.xlsx"
    caminho_saida = os.path.join(pasta_sap, nome_arquivo)
    wb_ekpo.save(caminho_saida)

    print_log(f"[{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Arquivo salvo como: {caminho_saida}")

    start_time = datetime.datetime.now()
    start_format = start_time.strftime("%d_%m às %H_%M_%S")

    pasta_base = os.path.join(os.path.expanduser("~"), r"OneDrive - Accenture\Documents\SAP\SAP GUI")
    print_log("Arquivos encontrados para Comparativo*.xlsx:")
    print_log(glob.glob(os.path.join(pasta_base, "Comparativo*.xlsx")))

    # Timestamp para nomear o arquivo
    timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")

    log_file = os.path.join(pasta_base, f"log_api_{timestamp}.txt")

    import logging

    log_file = os.path.join(pasta_base, f"log_api_{timestamp}.txt")
    logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.INFO, format='%(message)s')

    def print_log(msg):
        print(msg)
        logging.info(msg)


    # Registrar o tempo de início
    start_time = datetime.datetime.now()
    start_format = start_time.strftime("%d_%m às %H_%M_%S")

    def extract_error_details(response):
        try:
            error_data = response.json()
            detailed_errors = []

            if "errors" in error_data:
                for key, messages in error_data["errors"].items():
                    if messages:
                        for msg in messages:
                            detailed_errors.append(f"[{key}] {msg}")

            return "\n".join(detailed_errors) if detailed_errors else response.text
        except Exception:
            return response.text

    def format_error(status_code, message):
        return {"error": {"code": status_code, "message": message}}

    def limpar_nans(obj):
        if isinstance(obj, float):
            if isnan(obj) or isinf(obj):
                return None
            return obj
        elif isinstance(obj, dict):
            return {k: limpar_nans(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [limpar_nans(v) for v in obj]
        else:
            return obj


    AUTH_URL = 'xxxx'
    API_URL = 'xxxx'
    client = 'xxxx'
    client_secret = 'xxxx'
    scope = 'xxxx'
    grant_type = 'xxxx'
    token = None


    def get_jwt_token() -> str:
        global token
        global headers
        headers = {}

        print_log(f'Iniciando a requisição do token de acesso ao COUPA!')

        param = {
            'client_id': client,
            'client_secret': client_secret,
            'scope': scope,
            'grant_type': grant_type
        }

        response = requests.post(AUTH_URL, params=param)

        if response.status_code == 200:
            token = response.json().get("access_token")
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
            print_log(f'Token gerado com sucesso!')
            return headers
        else:
            raise Exception(
                f"Failed to get token: {response.status_code} - {response.text}"
            )

    import json

    import re

    def post_data(data, num_max_tentativas=3, status_token=False):
        num_tentativas = 0
        error_mensage = ""
        code_error_status = None

        while num_tentativas < num_max_tentativas:
            num_tentativas += 1
            try:
                if not status_token:
                    headers_api = get_jwt_token()
                    status_token = True

                response = requests.post(API_URL, json=data, headers=headers_api)

                if response.status_code in [200, 201]:
                    return response.json()

                elif response.status_code == 400:
                    try:
                        status_error = response.json()
                        order_errors = status_error.get("errors", {}).get("order-header", [])
                        status_error_str = str(order_errors).replace("'", "")
                        pattern = r'\\[Unable to find valid User record for ship_to_user with keys \\{"login"=>"[^"]+"\}'

                        if re.match(pattern, status_error_str):
                            data["ship-to-user"]["login"] = "dummy@teste.com"
                            response = requests.post(API_URL, json=data, headers=headers_api)
                            if response.status_code in [200, 201]:
                                return response.json()
                            else:
                                code_error_status = response.status_code
                                detailed_message = extract_error_details(response)
                                raise Exception(f"API Error:\n{detailed_message}")
                        else:
                            code_error_status = response.status_code
                            detailed_message = extract_error_details(response)
                            raise Exception(f"API Error:\n{detailed_message}")
                    except Exception as e:
                        raise Exception(str(e))

                else:
                    code_error_status = response.status_code
                    detailed_message = extract_error_details(response)
                    raise Exception(f"API Error:\n{detailed_message}")

            except Exception as e:
                status_token = False
                error_mensage = e
                print_log(f'Erro no POST: {e}')
        return format_error(code_error_status, f"{error_mensage}")


    def put_data(data, id_coupa):
        API_URL_PUT = f'xxxx/{id_coupa}/issue_without_send?return_object=limited'

        if token is None:
            get_jwt_token()

        response = requests.put(API_URL_PUT, json=data, headers=headers)

        if response.status_code in [200, 201]:
            return response.json()
        elif response.status_code == 401:
            get_jwt_token()
            return format_error(response.status_code, f"API Token Access Error: {response.text}")
        else:
            return format_error(response.status_code, f"API Error: {response.text}")


    def get_data(id_po):
        API_URL_GET = f'xxxx/?po_number={id_po}'

        if token is None:
            get_jwt_token()

        response = requests.get(API_URL_GET, headers=headers)

        if response.status_code in [200, 201]:
            dict_data = response.json()
            return dict_data[0]['id']
        
        elif response.status_code == 401:
            get_jwt_token()
            return format_error(response.status_code, f"API Token Access Error: {response.text}")
        else:
            return format_error(response.status_code, f"API Error: {response.text}")
            
    usuario = os.getlogin()
    pasta_base = f"C:\\Users\\{usuario}\\OneDrive - Accenture\\Documents\\SAP\\SAP GUI\\"

    def encontrar_arquivo_recente(padrao):
        arquivos = glob.glob(padrao)
        print(f"🔍 Arquivos encontrados: {arquivos}")
        if not arquivos:
            return None
        return max(arquivos, key=os.path.getmtime)

    # Caminho base
    usuario = os.getlogin()
    pasta_base = f"C:\\Users\\{usuario}\\OneDrive - Accenture\\Documents\\SAP\\SAP GUI\\"
    padrao_comparativo = os.path.join(pasta_base, "Comparativo_*.xlsx")
    arquivo_comparativo = encontrar_arquivo_recente(padrao_comparativo)

    # Verificação final
    print(f"🧪 Caminho final do Comparativo: {arquivo_comparativo}")
    print(f"📚 Tipo de dado recebido: {type(arquivo_comparativo)}")

    if arquivo_comparativo and isinstance(arquivo_comparativo, str):
        df_comparativo = pd.read_excel(arquivo_comparativo)
        print_log("✅ Comparativo carregado com sucesso.")
    else:
        print_log("❌ Falha ao localizar o arquivo de Comparativo.")
        exit()


    arquivo_ekko = encontrar_arquivo_recente(f"{pasta_base}CDHDR_EKKO_*.xlsx")
    arquivo_ekpo = encontrar_arquivo_recente(f"{pasta_base}CDHDR_EKPO_*.xlsx")
    arquivo_usr21 = encontrar_arquivo_recente(f"{pasta_base}CDHDR_USR21_*.xlsx")
    arquivo_adr6 = encontrar_arquivo_recente(f"{pasta_base}CDHDR_ADR6_*.xlsx")
    arquivo_lfa1 = encontrar_arquivo_recente(f"{pasta_base}CDHDR_LFA1_*.xlsx")
    arquivo_me23n = encontrar_arquivo_recente(f"{pasta_base}CDHDR_me23n_com_texto_*.txt")
    arquivo_mm03 = encontrar_arquivo_recente(f"{pasta_base}CDHDR_mm03_com_texto_*.txt")
    arquivo_contrato = encontrar_arquivo_recente(f"{pasta_base}CDHDR_CONTRATO_*.xlsx")
    arquivo_eket = encontrar_arquivo_recente(f"{pasta_base}CDHDR_EKET_*.xlsx")
    arquivo_mara = encontrar_arquivo_recente(f"{pasta_base}CDHDR_MARA_*.xlsx")

    print_log("\n✔️ Arquivos encontrados com sucesso.")

    df_ekko = pd.read_excel(arquivo_ekko, header=None, skiprows=5) if arquivo_ekko else None
    df_ekpo = pd.read_excel(arquivo_ekpo, header=None, skiprows=5) if arquivo_ekpo else None
    df_usr21 = pd.read_excel(arquivo_usr21, header=None, skiprows=5) if arquivo_usr21 else None
    df_adr6 = pd.read_excel(arquivo_adr6, header=None, skiprows=5) if arquivo_adr6 else None
    df_lfa1 = pd.read_excel(arquivo_lfa1, header=None, skiprows=5) if arquivo_lfa1 else None
    df_contrato = pd.read_excel(arquivo_contrato, header=None, skiprows=5) if arquivo_contrato else None
    df_eket = pd.read_excel(arquivo_eket, header=None, skiprows=5) if arquivo_eket else None
    df_mara = pd.read_excel(arquivo_mara, header=None, skiprows=5) if arquivo_mara else None

    # Normalizar as colunas do Comparativo
    df_comparativo["PO_number"] = df_comparativo["PO_number"].astype(str).str.strip()
    df_comparativo["PO_id"] = df_comparativo["PO_id"].astype(str).str.strip()
    df_comparativo["cod_elim"] = df_comparativo["cod_elim"].astype(str).str.strip().str.upper()

    # Obter POs da EKKO (coluna B, índice 1), a partir da linha 6 (índice 5)
    po_ekko = df_ekko.iloc[5:, 1].dropna().astype(str).str.strip().unique()

    # Aplicar as regras do filtro
    df_filtrado = df_comparativo[
        df_comparativo["PO_id"].notna() &
        (df_comparativo["PO_id"].astype(str).str.strip() != "") &
        (df_comparativo["PO_id"].str.lower() != "não consta no coupa") &
        df_comparativo["Line_id"].notna() &
        (df_comparativo["Line_id"].astype(str).str.strip() != "") &
        df_comparativo["cod_elim"].notna() &
        (df_comparativo["cod_elim"].astype(str).str.strip() != "") &
        (~df_comparativo["cod_elim"].astype(str).str.upper().isin(["L", "S"]))
    ]


    # Obter a lista final de POs válidas
    po_ids_unicos = df_filtrado["PO_id"].astype(str).str.strip().unique()


    def converter_tipo(valor):
        if isinstance(valor, (np.int64, np.float64)):
            return int(valor) if isinstance(valor, np.int64) else float(valor)
        if pd.isna(valor):
            return ""
        return str(valor)


    #Carregar descrições longas de mm03 em um dicionário com padding de zeros
    descricao_longa_dict = {}
    if arquivo_mm03:
        with open(arquivo_mm03, "r", encoding="utf-8") as file:
            for linha in file:
                partes = linha.strip().split("|")
                if len(partes) >= 2:
                    codigo_material = partes[0].strip().zfill(18)
                    descricao = partes[1].strip()
                    descricao_longa_dict[codigo_material] = descricao

    po_numeros_unicos = df_ekko.iloc[:, 1].dropna().unique() if df_ekko is not None else []
    pos_com_sucesso = []
    pos_com_erro = []

    df_ekpo.iloc[:, 1] = df_ekpo.iloc[:, 1].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )

    #Padroniza a coluna 1 da EKKO também
    df_ekko.iloc[:, 1] = df_ekko.iloc[:, 1].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )

    po_ids_unicos = df_filtrado["PO_id"].astype(str).str.strip().unique()

    for po_id in po_ids_unicos:  
        try:
            # Filtrar as linhas do df_filtrado para esse po_id
            linhas_po = df_filtrado[df_filtrado["PO_id"].astype(str).str.strip() == po_id]
            idcoupa = str(linhas_po.iloc[0]["PO_id"]).strip()

            # Pega o po_number da primeira linha (assumo que seja o mesmo para todas linhas do mesmo po_id)
            po_number = linhas_po.iloc[0]["PO_number"]
            po_number = (
                str(int(po_number)) if isinstance(po_number, float) and po_number.is_integer()
                else str(po_number).strip()
            )

            # Buscar a linha correspondente na EKKO pelo po_number
            linha_ekko = df_ekko[df_ekko.iloc[:, 1] == po_number]
            if linha_ekko.empty:
                print_log(f"⚠️ PO {po_number} (PO_ID {po_id}) ignorada: não encontrada na EKKO.")
                continue
            linha = linha_ekko.iloc[0]
            valor_total_bruto = converter_tipo(linha[10])
            incoterms_parte_2 = "" if pd.isna(linha[12]) else converter_tipo(linha[12])
            unidade_organizacional = converter_tipo(linha[13])
            grupo_comprador_raw = converter_tipo(linha[14])
            grupo_comprador = (
                str(grupo_comprador_raw).zfill(3)
                if str(grupo_comprador_raw).strip().isdigit()
                else str(grupo_comprador_raw).strip()
            )
            sua_referencia = converter_tipo(linha[8]) or "N/A"
            emissor_fatura_distinto = converter_tipo(linha[11])
            currency_code = converter_tipo(linha[3])

            #ship_to_user = "dummy@teste.com"

            ship_to_user = ""
            if df_usr21 is not None and df_adr6 is not None:
                usuario_ekko = linha[6]
                usr_match = df_usr21[df_usr21.iloc[:, 1] == usuario_ekko]
                if not usr_match.empty:
                    adr_match = df_adr6[df_adr6.iloc[:, 2] == usr_match.iloc[0, 2]]
                    ship_to_user = converter_tipo(adr_match.iloc[0, 8]) if not adr_match.empty else ""
                        
            location_code = ""
            if df_ekpo is not None:
                linhas_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]

                if not linhas_ekpo.empty:
                    for _, row in linhas_ekpo.iterrows():
                        line_num = converter_tipo(row[2])
                        centro_logistico = converter_tipo(row[8])  # Extraindo o valor correto
                        
                    location_code = centro_logistico 


            supplier_number = ""
            if df_lfa1 is not None:
                filtrado = df_lfa1[df_lfa1.iloc[:, 1] == linha[2]]
                
                if not filtrado.empty:
                    supplier_number = converter_tipo(filtrado.iloc[0, 2])
                    supplier_number_str = str(supplier_number).strip()

                    # Remove ".0" se for número vindo como float do Excel
                    if supplier_number_str.endswith(".0"):
                        supplier_number_str = supplier_number_str[:-2]

                    # Adiciona zeros se for só números
                    if supplier_number_str.isdigit():
                        supplier_number = supplier_number_str.zfill(14)
                    else:
                        supplier_number = supplier_number_str  # Mantém o original

            texto_cabecalho = ""
            if arquivo_me23n:
                with open(arquivo_me23n, "r", encoding="utf-8") as file:
                    for linha_txt in file:
                        if str(po_number) in linha_txt:
                            texto_cabecalho = linha_txt.split("|")[1].strip()
                            break

            # Inicializa a variável deposito
            deposito = ""

            if df_ekpo is not None:
                filtrado_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]  # Filtrando pelo PO

                if not filtrado_ekpo.empty and filtrado_ekpo.shape[1] > 17:  
                    deposito = converter_tipo(filtrado_ekpo.iloc[0, 17]) if pd.notna(filtrado_ekpo.iloc[0, 17]) else ""
                    print_log(f"Depósito extraído para PO {po_number}: {deposito}")  # Confirmação do valor
                else:
                    print_log(f"⚠ Atenção: Coluna 17 não encontrada na tabela EKPO para PO {po_number}.")



            json_data = {
                "type": "ExternalOrderHeader",
                "po-number": po_number,
                "version": 1,
                "payment-method": "invoice",
                "ship-to-attention": "",
                "ship-to-address": {"location-code": location_code},
                "ship-to-user": {
                    "login": ship_to_user
                }, 
                "supplier": {
                    "number": supplier_number
                },  
                "payment-term": {
                    "code": (
                        str(converter_tipo(linha[4])).zfill(4)
                            if pd.notna(linha[4]) and str(linha[4]).strip() != ""
                            else "Z034"
                        )
                }
                ,
                "shipping-term": {
                "code": (
                    "" if pd.isna(linha[5]) or str(linha[5]).strip() == "" 
                    else converter_tipo(linha[5])
                )
            },
                "custom-fields": {
                    "texto-de-cabecalho": converter_tipo(texto_cabecalho),
                "tipo-de-pedido": {
                        "external-ref-num": (
                            "" if pd.isna(linha[7]) or str(linha[7]).strip() == "" 
                            else converter_tipo(linha[7])
                        )
                    },
                    "valor-icms": "",
                    "valor-ipi": "",
                    "valor-icmsst": "",
                    "valor-total-bruto": valor_total_bruto,
                    "incoterms-parte-2": incoterms_parte_2,
                    "unidade-organizacional": unidade_organizacional,
                    "pedido-sap": po_number,
                    "grupo-comprador": {
                        "external-ref-num": grupo_comprador, 
                    },
                    "numero-rc": "",
                    "sua-referencia": converter_tipo(sua_referencia),
                    "empresa": "Energia",
                    "aceite-tacito": False,
                    "emissor-de-fatura-distinto": {
                        "external-ref-num": "N/A" if pd.isna(linha[11]) or linha[11] in [None, ""] else str(int(linha[11])) if isinstance(linha[11], float) and linha[11].is_integer() else str(linha[11]).strip()
                    }
                },
                "currency": {"code": currency_code},
                "order-lines": []
            }

            if df_ekpo is not None:
                linhas_ekpo = df_ekpo[df_ekpo.iloc[:, 1] == po_number]
                for _, row in linhas_ekpo.iterrows():
                    line_num = converter_tipo(row[2])
                    description = f"{converter_tipo(row[3])} | {converter_tipo(row[4])}"
                    valor_f_line_br = row[21]
                    if isinstance(valor_f_line_br, str):
                        valor_f_line_amer = valor_f_line_br.replace('.', '').replace(',', '.')
                    else:
                        valor_f_line_amer = f"{valor_f_line_br:.2f}"
                    quantity = converter_tipo(row[5])
                    need_by_date = ""

                    if df_eket is not None:
                        eket_match = df_eket[df_eket.iloc[:, 1] == po_number]
                        if not eket_match.empty:
                            raw_date = eket_match.iloc[0, 2]
                            try:
                                parsed_date = datetime.datetime.strptime(str(raw_date), "%d.%m.%Y")
                                need_by_date = parsed_date.strftime("%Y/%m/%d")
                            except ValueError:
                                pass

                    service_type = "non_service"
                    if df_mara is not None:
                        mara_match = df_mara[df_mara.iloc[:, 1] == row[4]]
                        if mara_match.empty:
                            mara_type = "quantity_deliverable"
                        else:
                            mara_type = converter_tipo(mara_match.iloc[0, 2])
                        if mara_type in ["DIEN", "ZIEN", "ZSER", "ZSGS", "quantity_deliverable"]:
                            service_type = "quantity_deliverable"
                    
                    tipo_delinha = "Material"  
                    if df_mara is not None:
                        mara_match = df_mara[df_mara.iloc[:, 1] == row[4]]
                        if mara_match.empty:
                            mara_type = "Serviço"
                        else:
                            mara_type = converter_tipo(mara_match.iloc[0, 2]) 
                        if mara_type in ["DIEN", "ZIEN", "ZSER", "ZSGS", "Serviço"]:
                            tipo_delinha = "Serviço"

                    codigo_material = converter_tipo(row[4]).strip().zfill(18)
                    descricao_longa = descricao_longa_dict.get(codigo_material, "")

                    data_contrato, id_contrato_coupa = "", ""
                    numero_contrato = str(int(row[14])) if pd.notna(row[14]) else ""
                    if numero_contrato and df_contrato is not None:
                        contrato_match = df_contrato[df_contrato.iloc[:, 1].astype(str).str.strip() == numero_contrato]
                        if not contrato_match.empty:
                            if pd.notna(contrato_match.iloc[0, 3]):
                                try:
                                    data_contrato = datetime.datetime.strptime(str(contrato_match.iloc[0, 3]), "%d.%m.%Y").strftime("%Y/%m/%d")
                                except ValueError:
                                    pass
                            id_contrato_coupa = converter_tipo(contrato_match.iloc[0, 5]) if pd.notna(contrato_match.iloc[0, 5]) else ""

                    centro_logistico = converter_tipo(row[8])
                    utilizacao_material = str(int(float(row[9]))) if pd.notna(row[9]) else "N/A"
                    origem_material_valor = converter_tipo(row[10]) if pd.notna(row[10]) else "N/A"
                    try:
                        if origem_material_valor.replace('.', '', 1).isdigit():
                            origem_material = int(float(origem_material_valor)) if float(origem_material_valor).is_integer() else float(origem_material_valor)
                        else:
                            origem_material = origem_material_valor
                    except:
                        origem_material = origem_material_valor

                    deposito_str = "N/A"  # Define N/A por padrão

                    if df_ekpo is not None:
                        match_ekpo = df_ekpo[
                            (df_ekpo.iloc[:, 1].astype(str).str.strip() == str(po_number)) &
                            (df_ekpo.iloc[:, 2].astype(str).str.strip() == str(line_num))
                        ]
                        
                        if not match_ekpo.empty and match_ekpo.shape[1] > 17:
                            deposito_raw = match_ekpo.iloc[0, 17]
                            deposito_str = str(deposito_raw).strip() if not pd.isna(deposito_raw) and deposito_raw not in ["", None] else "N/A"
                            
                        print_log(f"Depósito extraído para PO {po_number}, linha {line_num}: {deposito_str}")


                    json_data["order-lines"].append({
                        "line-num": line_num,
                        "description": description,
                        "price": valor_f_line_amer,
                        "quantity": quantity,
                        "need-by-date": need_by_date,
                        "type": "OrderQuantityLine",
                        "custom-fields": {
                            "descricao-longa": descricao_longa,
                            "centro-logistico": {
                            "external-ref-num": centro_logistico,  
                            },
                            "tipo-da-linha": tipo_delinha,
                            "utilizacao-do-material": {
                            "external-ref-num": utilizacao_material,  
                            },
                            "origem-do-material": {
                            "external-ref-num": origem_material,  
                            },
                            "codigo-ncm": str(int(row[11])) if isinstance(row[11], float) and row[11].is_integer() else str(row[11]).strip(),
                            "codigo-do-imposto": converter_tipo(row[12]),
                            "preco-por": (
                                converter_tipo(row[18]) if pd.notna(row[18]) and str(row[18]).strip() != "" else ""
                            ),
                            "deposito": {
                                "external-ref-num": deposito_str,  
                            },
                            "data-contrato": data_contrato,
                            "numero-do-contrato": numero_contrato,
                            "item-do-contrato": str(int(row[15])) if pd.notna(row[15]) else "",
                            "id-contrato-coupa": id_contrato_coupa
                        },
                        "uom": {"code": converter_tipo(row[13])},
                        "account": {
                            "code": "Dummy",
                            "segment-1": "Dummy",
                            "account-type": {
                                "name": "COA - zzzz"
                            },
                        },
                        "currency": {"code": currency_code},
                        "commodity":  {
                            "name": converter_tipo(row[16]),  # <- AJUSTADO
                        },
                        "service-type": service_type
                    })
            json_data = limpar_nans(json_data)
            print_log(f"\n🚚 Enviando PO {po_number}...")

            # 👇 Adicione esta linha para inspecionar o payload
            print_log(json.dumps(json_data, indent=2, ensure_ascii=False))  # ensure_ascii=False para exibir acentos corretamente


            response_post = post_data(json_data)
            if isinstance(response_post, dict) and "error" in response_post:
               raise Exception(f"Erro POST: {response_post['error']['message']}")
            
            idcoupa = get_data(json_data["po-number"])

            if isinstance(idcoupa, dict) and "error" in idcoupa:
                raise Exception(f"Erro GET: {idcoupa['error']['message']}")

            response_put = put_data(json_data, idcoupa)
            if isinstance(response_put, dict) and "error" in response_put:
                raise Exception(f"Erro PUT: {response_put['error']['message']}")

            json_data["api_responses"] = {
                "post_response": response_post,
                "get_response": idcoupa,
                "put_response": response_put
            }

            pos_com_sucesso.append(json_data)
            print_log(f"✅ PO {po_number} enviada com sucesso.)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              .")

        except Exception as e:
            erro_info = {
                "po_number": po_number,
                "erro": str(e)
            }
            pos_com_erro.append(erro_info)
            print_log(f"❌ Erro ao enviar PO {po_number}: {str(e)}")

    # Salvar os resultados em arquivos JSON separados
    timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")

    # Salvar os resultados em arquivos JSON separados com timestamp
    with open(os.path.join(pasta_base, f"sucesso_{timestamp}.json"), "w", encoding="utf-8") as f:
        json.dump(pos_com_sucesso, f, ensure_ascii=False, indent=4)

    with open(os.path.join(pasta_base, f"erros_{timestamp}.json"), "w", encoding="utf-8") as f:
        json.dump(pos_com_erro, f, ensure_ascii=False, indent=4)

    print_log(f"\n📦 Total de POs enviadas com sucesso: {len(pos_com_sucesso)}")
    print_log(f"⚠️ Total de erros: {len(pos_com_erro)}")
    print_log(f"📝 Arquivo de sucesso salvo em: {os.path.join(pasta_base, 'sucesso.json')}")
    print_log(f"📝 Arquivo de erros salvo em: {os.path.join(pasta_base, 'erros.json')}")

    import time
    time.sleep(3)  # Simulação de processamento

    # Registrar o tempo de término
    end_time = datetime.datetime.now()
    end_format = end_time.strftime("%d_%m às %H_%M_%S")

    # Calcular o tempo total
    execution_time = end_time - start_time

    # Converter para formato legível (hh:mm:ss)
    execution_seconds = execution_time.total_seconds()
    hours, remainder = divmod(execution_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    print_log(f"🟢 Início: {start_format}")
    print_log(f"🔴 Finalizou: {end_format}")
    print_log(f"⏱️ Tempo total de execução: {int(hours)}h {int(minutes)}m {int(seconds)}s")

    # --- Extraindo os erros ---
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font
    from openpyxl.chart import BarChart, Reference

    # Caminho da pasta SAP GUI
    user_profile = os.environ["USERPROFILE"]
    sap_folder = os.path.join(
        user_profile,
        "OneDrive - Accenture",
        "Documents",
        "SAP",
        "SAP GUI"
    )

    # Subpasta de destino
    output_folder = os.path.join(sap_folder, "Analise")
    os.makedirs(output_folder, exist_ok=True)

    # Arquivo JSON mais recente
    json_files = glob.glob(os.path.join(sap_folder, "erros_*.json"))
    if not json_files:
        print_log("Nenhum arquivo de erro encontrado.")
        exit()

    latest_json = max(json_files, key=os.path.getctime)

    # Formatar o nome com dd_mm_HH_MM
    now = datetime.datetime.now()
    data_hora_str = now.strftime("%d_%m_%H_%M")
    excel_name = f"analise_erro_{data_hora_str}.xlsx"
    excel_path = os.path.join(output_folder, excel_name)

    # Carregar JSON
    with open(latest_json, "r", encoding="utf-8") as f:
        dados = json.load(f)

    # Lista consolidada de linhas
    linhas = []
    po_processadas = set()

    for item in dados:
        po = item.get("po_number", "")
        msg = item.get("erro", "")
        msg_lower = msg.lower()
        erro_identificado = False
        valor_extraido = ""

        if "ssl" in msg_lower or "certificate verify failed" in msg_lower:
            linhas.append({
                "Po_Number": po,
                "Erro": "SSL: CERTIFICATE_VERIFY_FAILED",
                "Valor": "",
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if 'Unable to find valid Supplier' in msg:
            msg_normalizada = msg.replace('\\u003e', '=>').replace('\u003e', '=>')
            msg_normalizada = msg_normalizada.replace('\\"', '"').replace('\\\\', '')

            match = re.search(r'number["=\s]*=>["=\s]*"([\w\d]+)', msg_normalizada)
            valor_extraido = match.group(1) if match else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "SUPPLIER",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "LookupValue record for emissor_de_fatura_distinto" in msg:
            matches = re.findall(r'external_ref_num["=\s]*=>["=\s]*"([\w\d]+)', msg)
            valor_extraido = ", ".join(matches) if matches else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "EMISSOR DE FATURA",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "LookupValue record for deposito/custom_field_4" in msg:
            matches = re.findall(r'external_ref_num["=\s]*=>["=\s]*"([\w\d]+)', msg)
            valor_extraido = ", ".join(matches) if matches else ""

            linhas.append({
                "Po_Number": po,
                "Erro": "DEPOSITO",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if "has already been taken" in msg:
            linhas.append({
                "Po_Number": po,
                "Erro": "PO DUPLICADO",
                "Valor": po,
                "Mensagem_erro": msg
            })
            erro_identificado = True

        if not erro_identificado:
            if po not in po_processadas:
                valor = ""
                if "Unable to find valid Address record for ship_to_address" in msg:
                    valor = "Verificar Status da PO"

                linhas.append({
                    "Po_Number": po,
                    "Erro": "ERRO DESCONHECIDO",
                    "Valor": valor,
                    "Mensagem_erro": msg
                })

        if "Unable to find valid Uom record for uom" in msg:
            match = re.search(r'code["=\s]*=>["=\s]*"?(\w+)"?', msg)
            valor_extraido = match.group(1) if match else ""
            linhas.append({
                "Po_Number": po,
                "Erro": "UOM INVALIDA",
                "Valor": valor_extraido,
                "Mensagem_erro": msg
            })

            po_processadas.add(po)

    # Gerar DataFrame
    df_erros = pd.DataFrame(linhas)
    df_erros = df_erros[["Po_Number", "Erro", "Valor", "Mensagem_erro"]]

    # Gerar levantamento de erros com valores
    df_levantamento = df_erros.groupby(["Erro"])["Valor"].apply(lambda x: ", ".join(map(str, x.dropna().unique()))).reset_index()
    df_levantamento["Quantidade"] = df_erros["Erro"].value_counts().values

    # Salvar Excel com múltiplas abas
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_erros.to_excel(writer, sheet_name="Erros Detalhados", index=False)
        df_levantamento.to_excel(writer, sheet_name="Levantamento", index=False)

    # Abrir arquivo e adicionar gráfico
    wb = load_workbook(excel_path)
    ws = wb["Levantamento"]

    # Estilizar cabeçalhos
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Criar gráfico de barras
    chart = BarChart()
    chart.title = "Distribuição de Erros"
    chart.x_axis.title = "Tipo de Erro"
    chart.y_axis.title = "Quantidade"
    chart.style = 10

    data = Reference(ws, min_col=3, min_row=1, max_row=len(df_levantamento)+1)
    categories = Reference(ws, min_col=1, min_row=2, max_row=len(df_levantamento)+1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    ws.add_chart(chart, "E5")

    wb.save(excel_path)
    print_log(f"Arquivo salvo com gráfico e levantamento: {excel_path}")

    # --- COM SUCESSO
    # Caminho da pasta principal SAP GUI
    pasta_sap = os.path.expanduser(r"~\OneDrive - Accenture\Documents\SAP\SAP GUI")

    # Subpasta 'Analise'
    pasta_analise = os.path.join(pasta_sap, "Analise")
    os.makedirs(pasta_analise, exist_ok=True)

    # Localiza o arquivo JSON mais recente com sucesso
    arquivos_json = glob.glob(os.path.join(pasta_sap, "sucesso_*.json"))
    if not arquivos_json:
        print_log("❌ Nenhum arquivo de sucesso encontrado.")
        exit()

    arquivo_mais_recente = max(arquivos_json, key=os.path.getmtime)

    # Carrega o conteúdo do JSON
    with open(arquivo_mais_recente, encoding='utf-8') as f:
        dados = json.load(f)

    # Extrai os dados desejados corretamente
    linhas = []
    for item in dados:
        po = item.get("po-number")
        
        api_responses = item.get("api_responses", {})
        get_response = api_responses.get("get_response", None)
        
        put_response = api_responses.get("put_response", {})
        put_id = put_response.get("id") if isinstance(put_response, dict) else None

        linhas.append({
            "po-number": po,
            "get_response": get_response,
            "put_response": json.dumps(put_response, ensure_ascii=False),
            "id": put_id
        })

    # Cria o DataFrame
    df = pd.DataFrame(linhas)

    # Nome dinâmico com data/hora
    agora = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"PO_SUCESSO_{agora}.xlsx"
    caminho_saida = os.path.join(pasta_analise, nome_arquivo)

    # Salva no Excel
    df.to_excel(caminho_saida, index=False)

    print_log(f"✅ Arquivo gerado com sucesso: {caminho_saida}")

    print_log(f"✅ Processo de envio a Api finalizado")
# ======== ATUALIZAÇÕES DA PO =============


def esperar_carregamento(session, timeout=300):
    """ Aguarda o SAP concluir o processamento antes de continuar """
    tempo_inicial = time.time()
    while time.time() - tempo_inicial < timeout:
        try:
            # Verifica se a barra de status está ativa, indicando processamento
            if session.findById("wnd[0]/sbar").text.strip():
                time.sleep(2)  # Aguarda 2 segundos e tenta novamente
            else:
                return True  # O SAP terminou o processamento
        except:
            time.sleep(1)  
    
    raise Exception("Tempo limite atingido ao aguardar o processamento do relatório no SAP.")

#CDHDR
def atualizar():
    atualizar_status("🔄 Buscando atualizações...")
    atualizar_barra_progresso(0.1)
    app.update()
    
    data_de = entry_data_de.get()
    data_ate = entry_data_ate.get()

    if not data_de or not data_ate or data_de == "dd.mm.aaaa" or data_ate == "dd.mm.aaaa":
        atualizar_status("❌ Erro: Preencha as datas DE e ATÉ corretamente.")
        atualizar_barra_progresso(0)
        return

    try:
        # Conectar ao SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        print_log_sap("⚙ Iniciando extrações CDHDR..")
        atualizar_status("⚙ Iniciando extrações CDHDR..")
        atualizar_barra_progresso(0.03)
        app.update() 
        time.sleep(2)

        # Acessa a transação SE16N
        session.findById("wnd[0]/tbar[0]/okcd").text = "SE16N"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()

        # Define a tabela como CDHDR
        session.findById("wnd[0]/usr/ctxtGD-TAB").text = "CDHDR"

        # Carregar variante
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
        time.sleep(0.5)

        session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "CDHDR_AUTOMACAO"
        time.sleep(0.5)

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.5)

        # Define os filtros de data
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,5]").text = data_de
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,5]").text = data_ate

        # Ajusta o foco e a posição do cursor
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,5]").setFocus()
        session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,5]").caretPosition = 10

        # Executa a consulta
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        '''# Usa as datas digitadas na interface
        session.findById("wnd[0]/usr/ctxtI5-LOW").text = data_de
        session.findById("wnd[0]/usr/ctxtI5-HIGH").text = data_ate
        session.findById("wnd[0]/usr/txtMAX_SEL").text = ""
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()'''

        esperar_carregamento(session)

        # Verificar se há mensagem de "nenhuma entrada encontrada"
        mensagem = session.findById("wnd[0]/sbar").Text
        if "Não foi encontrada nenhuma entrada" in mensagem:
             return
        print_log_sap("⚠ Nenhum dado encontrado para exportar.")
        atualizar_status("⚠ Nenhum dado encontrado para exportar.")
        atualizar_barra_progresso(1)
            
        # Gerar o nome do arquivo com base na data atual
        #data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
        #nome_arquivo = f"EXPORT_CDHDR_{data_atual}.XLS"

        print_log_sap("💾 Salvando tabela CDHDR...")
        atualizar_status("💾 Salvando tabela CDHDR...")
        atualizar_barra_progresso(0.05)
        app.update()

        # Abre o menu de exportação
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")
        time.sleep(0.5)

        # Seleciona "Spreadsheet"
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.5)

        # Gera nome dinâmico e define no campo
        # Define o nome do arquivo e confirma o salvamento
        nome_arquivo = f"CDHDR_{datetime.datetime.now():%d_%m_%H_%M_%S}.XLS"

        # Define o nome do arquivo no SAP
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nome_arquivo)

        # Confirma o salvamento
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        print_log_sap("✅ Dados CDHDR salvos com sucesso.")
        atualizar_status("✅ Dados CDHDR salvos com sucesso.")
        atualizar_barra_progresso(0.10)
        app.update()

    except Exception as e:
        atualizar_status(f"❌ Erro ao executar atualização: {str(e)}")
        atualizar_barra_progresso(0)
        print_log_sap(f"❌ Erro ao executar atualização: {str(e)}") 

# Função para converter .xls para .xlsx    
    def converter_xls_para_xlsx(caminho_arquivo_xls):
    # Verificar se o arquivo existe
            if not os.path.exists(caminho_arquivo_xls):
                print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                return None

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False  # Não mostrar a interface do Excel
            try:
        # Abrir o arquivo diretamente no modo de fundo sem exibir
                wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                wb.Close()
                print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                return novo_caminho
            except Exception as e:
                print_log_sap("Erro ao converter:", e)
                return None
            finally:
                excel.Quit()

# Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("CDHDR_*.xls"))
    
            if not arquivos:
                print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                return None
    
    # Ordenar os arquivos por data de modificação (mais recente primeiro)
            arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
            return arquivo_mais_recente

# Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
            print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
            arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
            if arquivo_convertido:
                print_log_sap(f"Arquivo convertido: {arquivo_convertido}")


### Manipulando dados para EKKO da CDHDR

    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("CDHDR_*.xlsx"))

        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado CDHDR_.")
            return None

        return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_c(arquivo_xlsx):
    # Lê o arquivo Excel
        df = pd.read_excel(arquivo_xlsx)

        # Verifica se a coluna C (índice 2) existe
        if df.shape[1] > 2:
            coluna_c = df.iloc[5:, 2]  # A partir da linha 6 (índice 5), coluna C (índice 2)

            # Limpa os dados: remove nulos e espaços extras
            dados_coluna_c = coluna_c.dropna().apply(lambda x: str(x).strip()).tolist()

            return dados_coluna_c
        else:
            print_log_sap("A coluna C não foi encontrada no arquivo.")
            return None

# Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
        print_log_sap(f"Arquivo encontrado: {arquivo}")
        valorobjeto = obter_dados_coluna_c(arquivo)

        if valorobjeto:
        # Juntar os valores da coluna H com \r\n (quebra de linha para o SAP)
            texto_para_copiar = '\r\n'.join(valorobjeto)

        # Copiar para a área de transferência
            try:
                pyperclip.copy(texto_para_copiar)
                print_log_sap("\nDados extraídos e copiados para a área de transferência:")
                print_log_sap(texto_para_copiar)
                print_log_sap(f"\nTotal de {len(valorobjeto)} valores copiados.")
            except pyperclip.PyperclipException:
                print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                print_log_sap("Dados extraídos e copiados:")
                print_log_sap(texto_para_copiar)
                print_log_sap(f"\nTotal de {len(valorobjeto)} valores copiados.")
        else:
            print_log_sap("Nenhum dado encontrado na coluna C.")

# Voltar duas vezes
    session.findById("wnd[0]/tbar[0]/btn[12]").press()
    print_log_sap("⚙ Iniciando extração EKKO referente a CDHDR")
    atualizar_status("⚙ Iniciando extração EKKO referente a CDHDR")
    atualizar_barra_progresso(0.15)
    app.update()

# Acessar a transação da tabela EKKO
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKKO"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()


        # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO CDHDR"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)

    print_log_sap("⚙ Iniciando extrações EKKO referente a CDHDR")
    atualizar_status("⚙ Iniciando extrações EKKO referente a CDHDR")
    atualizar_barra_progresso(0.03)
    app.update() 

    # Abrir seleção múltipla (botão lupa)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()

    # Colar os valores (copiados anteriormente com `\r\n`)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()  # Colar da área de transferência
    session.findById("wnd[1]/tbar[0]/btn[8]").press()   # Confirmar

        # Executar
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Aguarda o carregamento do grid por segurança
    time.sleep(1.5)  # ajuste se necessário
 # Abre o menu de variantes
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")

    # Seleciona a opção "Carregar" (Load)
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&LOAD")

    # Obtém o grid corretamente
    # Aguarda o carregamento do grid por segurança
    time.sleep(1.5)  # ajuste se necessário

    grid = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")

    # Tenta encontrar a linha onde a coluna VARIANT tem o valor '/RPA'
    for i in range(0, grid.RowCount):
        try:
            valor = grid.GetCellValue(i, "VARIANT")
            if valor.strip().upper() == "/RPA":
                grid.currentCellRow = i
                grid.selectedRows = str(i)
                grid.clickCurrentCell()
                break
        except:
            pass

    # clicar no botão de OK 
    #session.findById("wnd[1]/tbar[0]/btn[0]").press()

    time.sleep(2)

    # Exportar os dados
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Selecionar formato "Planilha de cálculo"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    nome_arquivo = f"CDHDR_EKKO_{datetime.datetime.now():%d_%m_%H_%M_%S}.XLS"
    # Nomear e salvar o arquivo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nome_arquivo)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Voltar
    session.findById("wnd[0]/tbar[0]/btn[12]").press()

    print_log_sap("✅ Dados EKKO referente a CDHDR extraidoss e salvos com sucesso")
    atualizar_status("✅ Dados EKKO referente a CDHDR extraidoss e salvos com sucesso")
    atualizar_barra_progresso(0.20)
    app.update()

# Função para converter .xls para .xlsx
    def converter_xls_para_xlsx(caminho_arquivo_xls):
    # Verificar se o arquivo existe
            if not os.path.exists(caminho_arquivo_xls):
                print(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                return None

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False  # Não mostrar a interface do Excel
            try:
        # Abrir o arquivo diretamente no modo de fundo sem exibir
                wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                wb.Close()
                print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                return novo_caminho
            except Exception as e:
                print_log_sap("Erro ao converter:", e)
                return None
            finally:
                excel.Quit()

# Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("CDHDR_EKKO_*.xls"))
    
            if not arquivos:
                print_log_sap("Nenhum arquivo .xls encontrado na pasta CDHDR_EKKO_.")
                return None
    
    # Ordenar os arquivos por data de modificação (mais recente primeiro)
            arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
            return arquivo_mais_recente

# Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
            print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
            arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
    if arquivo_convertido:
                print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

### Manipulando dados para LFA1

    def localizar_arquivo_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("CDHDR_EKKO_*.xlsx"))

            if not arquivos:
                print_log_sap("Nenhum arquivo encontrado CDHDR_EKKO_.")
                return None

            return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_m(arquivo_xlsx):
        # Usando pandas para ler o arquivo Excel
            df = pd.read_excel(arquivo_xlsx)

            # Acessando a coluna M (a 13ª coluna, que tem o índice 12)
            if df.shape[1] > 12:  # Verifica se a coluna M (índice 12) existe
                coluna_m = df.iloc[4:, 2]  # Começando da linha 6 (índice 5) e acessando a coluna M (índice 12)

                # Limpando os dados (removendo valores nulos e espaços extras)
                dados_coluna_m = coluna_m.dropna().apply(lambda x: str(x).strip()).tolist()

                return dados_coluna_m
            else:
                print_log_sap("A coluna M não foi encontrada no arquivo.")
                return None

# Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
            print(f"Arquivo encontrado: {arquivo}")
            fornecedores = obter_dados_coluna_m(arquivo)

            if fornecedores:
                # Juntar os valores da coluna I com \r\n (quebra de linha para o SAP)
                texto_para_copiar = '\r\n'.join(fornecedores)

        # Copiar para a área de transferência
                try:
                    pyperclip.copy(texto_para_copiar)
                    print_log_sap("\nDados extraídos e copiados para a área de transferência:")
                    print_log_sap(texto_para_copiar)
                    print_log_sap(f"\nTotal de {len(fornecedores)} valores copiados.")
                except pyperclip.PyperclipException:
                    print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                    print_log_sap("Dados extraídos e copiados:")
                    print_log_sap(texto_para_copiar)
                    print_log_sap(f"\nTotal de {len(fornecedores)} valores copiados.")
            else:
                print_log_sap("Nenhum dado encontrado na coluna I.")

    print_log_sap("⚙ Iniciando extração LFA1")
    atualizar_status("⚙ Iniciando extração LFA1")
    atualizar_barra_progresso(0.10)
    app.update() 

# ==== EXTRAÇÃO LFA1   ====
    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "LFA1"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO LFA1"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)

    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abre o menu de contexto da toolbar de resultados (botão "Exportar")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # Seleciona a opção "Planilha..." no menu de contexto
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Seleciona o formato de exportação (por exemplo, planilha Excel no formato interno SAP)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirma a seleção do formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print_log_sap("💾 Salvando tabela LFA1...")
    atualizar_status("💾 Salvando tabela LFA1...")
    atualizar_barra_progresso(0.15)
    app.update()

    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"CDHDR_LFA1_{data_atual}.XLS"
    # Define o nome do arquivo a ser salvo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirma a exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
    # Voltar 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    print_log_sap ("✅ Dados LFA extraidos e salvos com sucesso")
    atualizar_status("✅ Dados LFA extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.16)
    app.update() 

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("CDHDR_LFA1_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

# ==== Manipulando dados para EKPO   ====
    print_log_sap("⏳ Manipulando dados para extração EKPO")
    atualizar_status("⏳ Manipulando dados para extração EKPO")
    atualizar_barra_progresso(0.18)
    app.update()

    def localizar_arquivo_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("CDHDR_EKKO_*.xlsx"))

            if not arquivos:
                print_log_sap("Nenhum arquivo encontrado para manipular dados para EKPO.")
                return None

            return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_b(arquivo_xlsx):
            # Usando pandas para ler o arquivo Excel
            df = pd.read_excel(arquivo_xlsx)

            # Acessando a coluna B (a 2ª coluna, que tem o índice 1)
            if df.shape[1] > 1:  # Verifica se a coluna B (índice 1) existe
                coluna_b = df.iloc[4:, 1]  # Começando da linha 6 (índice 5) e acessando a coluna B (índice 1)

                # Limpando os dados (removendo valores nulos e espaços extras)
                dados_coluna_b = coluna_b.dropna().apply(lambda x: str(x).strip()).tolist()

                return dados_coluna_b
            else:
                print_log_sap("A coluna B não foi encontrada no arquivo.")
                return None

        # Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
            print_log_sap(f"Arquivo encontrado: {arquivo}")
            docompras = obter_dados_coluna_b(arquivo)

            if docompras:
                # Juntar os valores da coluna B com \r\n (quebra de linha para o SAP)
                texto_para_copiar = '\r\n'.join(docompras)

                # Copiar para a área de transferência
                try:
                    pyperclip.copy(texto_para_copiar)
                    print("\nDados extraídos e copiados para a área de transferência:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(docompras)} valores copiados.")
                except pyperclip.PyperclipException:
                    print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                    print("Dados extraídos e copiados:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(docompras)} valores copiados.")
            else:
                print_log_sap("Nenhum dado encontrado na coluna B.")

    print_log_sap("⚙ Inicinado extração EKPO")
    atualizar_status("⚙ Inicinado extração EKPO")
    atualizar_barra_progresso(0.20)
    app.update()

# ==== EXTRAÇÃO EKPO   ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKPO"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO CDHDR"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)
    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abre o menu de variantes
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")

        # Seleciona a opção "Carregar" (Load)
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&LOAD")

        # Obtém o grid corretamente
        # Aguarda o carregamento do grid por segurança
    time.sleep(1.5)  # ajuste se necessário

    grid = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")

        # Tenta encontrar a linha onde a coluna VARIANT tem o valor '/CDHDR'
    for i in range(0, grid.RowCount):
            try:
                valor = grid.GetCellValue(i, "VARIANT")
                if valor.strip().upper() == "/CDHDR":
                    grid.currentCellRow = i
                    grid.selectedRows = str(i)
                    grid.clickCurrentCell()
                    break
            except:
                pass

    # Abre o menu de contexto da toolbar de resultados (botão "Exportar")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # Seleciona a opção "Planilha..." no menu de contexto
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Seleciona o formato de exportação (por exemplo, planilha Excel no formato interno SAP)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirma a seleção do formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    print_log_sap("💾 Salvando tabela EKPO...")
    atualizar_status("💾 Salvando tabela EKPO...")
    atualizar_barra_progresso(0.25)
    app.update()

    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"CDHDR_EKPO_{data_atual}.XLS"
    # Define o nome do arquivo a ser salvo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirma a exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
    # Voltar 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    print_log_sap("✅ Dados CDHDR EKPO extraidos e salvos com sucesso")
    atualizar_status("✅ Dados CDHDR EKPO extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.26)
    app.update() 


# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("CDHDR_EKPO_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

# ====  Manipulando dados para MARA  ====
    print_log_sap("⏳ Manipulando dados para extração Mara referente CDHDR")
    atualizar_status("⏳ Manipulando dados para extração Mara referente CDHDR")
    atualizar_barra_progresso(0.27)
    app.update()

# Função para localizar o arquivo mais recente na pasta
    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("CDHDR_EKPO_*.xlsx"))

        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para manipulação de dados para Mara.")
            return None

        # Encontrar o arquivo mais recente baseado na data de modificação
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    # Função para obter os dados da coluna E a partir de um arquivo XLSX
    def obter_dados_coluna_e(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        dados_coluna_e = []

        # Acessando a coluna E (5ª coluna)
        for row in sheet.iter_rows(min_row=6, min_col=5, max_col=5):  # Coluna E = índice 5
            for cell in row:
                if cell.value is not None:
                    dados_coluna_e.append(str(cell.value))  # Converte o valor para string

        return dados_coluna_e

    # Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
        print_log_sap(f"Arquivo encontrado: {arquivo}")
        tipmaterial = obter_dados_coluna_e(arquivo)

        if tipmaterial:
            print("Dados da coluna E:", '\n'.join(tipmaterial))
            
            # Copia os valores da coluna E para a área de transferência, um valor por linha
            pyperclip.copy('\r\n'.join(tipmaterial))  # Usando '\r\n' para SAP reconhecer quebra de linha
            print_log_sap("Valores copiados para a área de transferência.")
        else:
            print_log_sap("Nenhum dado encontrado na coluna E.")

    print_log_sap("⚙ Inicinado extração MARA referente CDHDR...")
    atualizar_status("⚙ Inicinado extração MARA...")
    atualizar_barra_progresso(0.30)
    app.update() 

# ==== EXTRAÇÃO MARA   ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "MARA"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO MARA"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)
    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abre o menu de contexto da toolbar de resultados (botão "Exportar")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # Seleciona a opção "Planilha..." no menu de contexto
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Seleciona o formato de exportação (por exemplo, planilha Excel no formato interno SAP)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirma a seleção do formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print_log_sap("💾 Salvando tabela MARA referente CDHDR...")
    atualizar_status("💾 Salvando tabela MARA...")
    atualizar_barra_progresso(0.35)
    app.update()

    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"CDHDR_MARA_{data_atual}.XLS"
    # Define o nome do arquivo a ser salvo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirma a exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
    # Voltar 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    print_log_sap("✅ Dados MARA refernete CDHDR extraidos e salvos com sucesso")
    atualizar_status("✅ Dados MARA referente CDHDR extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.36)
    app.update() 

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("CDHDR_MARA_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

#=== Manipulando dados para EKET ===
    print_log_sap("⏳ Manipulando dados para extração EKET referente CDHDR")
    atualizar_status("⏳ Manipulando dados para extração EKET referente CDHDR")
    atualizar_barra_progresso(0.40)
    app.update()

# Função para localizar o arquivo mais recente na pasta
    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("CDHDR_EKPO_*.xlsx"))

        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para Manipular dodos para EKET.")
            return None

        # Encontrar o arquivo mais recente baseado na data de modificação
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    # Função para obter os dados da coluna B a partir de um arquivo XLSX
    def obter_dados_coluna_b(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        dados_coluna_b = []

        # Acessando a coluna B (2ª coluna)
        for row in sheet.iter_rows(min_row=6, min_col=2, max_col=2):  # Coluna B = índice 2
            for cell in row:
                if cell.value is not None:
                    dados_coluna_b.append(str(cell.value))  # Converte o valor para string

        return dados_coluna_b

    # Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
        print_log_sap(f"Arquivo encontrado: {arquivo}")
        doccompras = obter_dados_coluna_b(arquivo)

        if doccompras:
            print("Dados da coluna B (DocCompras):", '\n'.join(doccompras))

            # Copia os valores da coluna B para a área de transferência, um valor por linha
            pyperclip.copy('\r\n'.join(doccompras))  # Usando '\r\n' para SAP reconhecer quebra de linha
            print_log_sap("Valores copiados para a área de transferência.")
        else:
            print_log_sap("Nenhum dado encontrado na coluna B.")

    print_log_sap("⚙ Inicinado extração EKET referente CDHDR")
    atualizar_status("⚙ Inicinado extração EKET referente CDHDR")
    atualizar_barra_progresso(0.42)
    app.update() 

# ==== EXTRAÇÃO EKET   ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKET"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carregar variante
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO EKET"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)
    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abre o menu de contexto da toolbar de resultados (botão "Exportar")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

    # Seleciona a opção "Planilha..." no menu de contexto
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Seleciona o formato de exportação (por exemplo, planilha Excel no formato interno SAP)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirma a seleção do formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print_log_sap("💾 Salvando tabela EKET...")
    atualizar_status("💾 Salvando tabela EKET...")
    atualizar_barra_progresso(0.45)
    app.update()

    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"CDHDR_EKET_{data_atual}.XLS"
    # Define o nome do arquivo a ser salvo
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirma a exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
    # Voltar 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    print_log_sap("✅ Dados EKET referente CDHDR extraidos e salvos com sucesso")
    atualizar_status("✅ Dados EKET referente CDHDR extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.46)
    app.update()

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("CDHDR_EKET_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")


# ==== Ekko CONTRATO CDHDR ====
    # Acessa a transação SE16N para a tabela EKKO
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKKO"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Carrega a variante /EXTPO CONT
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO CONT"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.5)

    time.sleep(0.5)
    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)

# ==== Manipulando dados para EKKO CONTRATO   ====

# Função para localizar o arquivo mais recente na pasta
    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("CDHDR_EKPO_*.xlsx"))

        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para manipular dados referente EKKO CONTRATO.")
            return None

        # Encontrar o arquivo mais recente baseado na data de modificação
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    # Função para obter os dados da coluna B a partir de um arquivo XLSX
    def obter_dados_coluna_b(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        dados_coluna_b = []

        # Acessando a coluna O (2ª coluna)
        for row in sheet.iter_rows(min_row=6, min_col=15, max_col=15):  # Coluna O = índice 2
            for cell in row:
                if cell.value is not None:
                    dados_coluna_b.append(str(cell.value))  # Converte o valor para string

        return dados_coluna_b

    # Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
        print_log_sap(f"Arquivo encontrado: {arquivo}")
        doccompras = obter_dados_coluna_b(arquivo)

        if doccompras:
            print("Dados da coluna B (DocCompras):", '\n'.join(doccompras))

            # Copia os valores da coluna B para a área de transferência, um valor por linha
            pyperclip.copy('\r\n'.join(doccompras))  # Usando '\r\n' para SAP reconhecer quebra de linha
            print_log_sap("Valores copiados para a área de transferência.")
        else:
            print_log_sap("Nenhum dado encontrado na coluna B.")
    
    # Pressionar botão "Colar da área de transferência" (ícone de prancheta)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    atualizar_status("⚙ Iniciando extrações EKKO CONTRATO REFERENT CDHDR")
    atualizar_barra_progresso(0.53)
    app.update()

    esperar_carregamento(session)

    mensagem = session.findById("wnd[0]/sbar").Text
    if "Nenhum valor encontrado" in mensagem:
        atualizar_status("⚠ Nenhum dado encontrado para exportar.")
        atualizar_barra_progresso(0.55)

        # Voltar para tela inicial
        try:
            session.findById("wnd[0]/tbar[0]/btn[12]").press()
        except:
            pass

        # Pular exportação e conversão — segue direto com o restante do fluxo
    else:
        # --- Continua se houver dados ---

        # Abre o menu de variantes
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&LOAD")
        time.sleep(1.5)

        # Seleciona a variante /RPA CONT
        grid = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")

        # Tenta encontrar a linha onde a coluna VARIANT tem o valor '/RPA'
        for i in range(0, grid.RowCount):
            try:
                valor = grid.GetCellValue(i, "VARIANT")
                if valor.strip().upper() == "/RPA_CONT":
                    grid.currentCellRow = i
                    grid.selectedRows = str(i)
                    grid.clickCurrentCell()
                    break
            except:
                pass

        # clicar no botão de OK 
        #session.findById("wnd[1]/tbar[0]/btn[0]").press()

        time.sleep(2)

        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

        # Seleciona "Planilha eletrônica"
        for i in range(5):
            try:
                opcao = session.findById(f"wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[{i},0]")
                if opcao.text == "Planilha eletrônica":
                    opcao.select()
                    opcao.setFocus()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    break
            except:
                pass

        # Gera o nome do arquivo
        data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
        nome_arquivo = f"CDHDR_CONTRATO_{data_atual}.XLS"

        # Salva o arquivo
        try:
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception as e:
            atualizar_status("❌ Erro ao salvar o arquivo.")
            executando = False
        else:
            # Volta para a tela anterior
            try:
                session.findById("wnd[0]/tbar[0]/btn[12]").press()
            except Exception as e:
                atualizar_status(f"❌ Erro inesperado: {e}")

            print_log_sap("✅ Dados EKKO referente CDHDR extraídos e salvos com sucesso")
            # Finaliza com sucesso
            atualizar_status("✅ Dados EKKO referente CDHDR extraídos e salvos com sucesso")

            # ==== Conversão de XLS para XLSX ====
            def converter_xls_para_xlsx(caminho_arquivo_xls):
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False
                try:
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

            def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("CDHDR_CONTRATO_*.xls"))
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
                return max(arquivos, key=os.path.getmtime)

            # Executa conversão
            arquivo_xls = encontrar_arquivo_export_mais_recente()
            if arquivo_xls:
                print(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

        atualizar_barra_progresso(0.60)
        app.update()

# ==== Manipulando dados para USR21   ====

    def localizar_arquivo_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("CDHDR_EKKO_*.xlsx"))

            if not arquivos:
                print_log_sap("Nenhum arquivo encontrado para manipulação da URS21.")
                return None

            return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_m(arquivo_xlsx):
    # Usando pandas para ler o arquivo Excel
        df = pd.read_excel(arquivo_xlsx)

        # Acessando a coluna M (a 13ª coluna, que tem o índice 12)
        if df.shape[1] > 6:  # Verifica se a coluna M (índice 12) existe
            coluna_m = df.iloc[4:, 6]  # Começando da linha 6 (índice 5) e acessando a coluna M (índice 12)

            # Limpando os dados (removendo valores nulos e espaços extras)
            dados_coluna_m = coluna_m.dropna().apply(lambda x: str(x).strip()).tolist()

            return dados_coluna_m
        else:
            print_log_sap("A coluna G não foi encontrada no arquivo.")
            return None

# Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
            print_log_sap(f"Arquivo encontrado: {arquivo}")
            criadopor = obter_dados_coluna_m(arquivo)

            if criadopor:
                # Juntar os valores da coluna I com \r\n (quebra de linha para o SAP)
                texto_para_copiar = '\r\n'.join(criadopor)

        # Copiar para a área de transferência
                try:
                    pyperclip.copy(texto_para_copiar)
                    print("\nDados extraídos e copiados para a área de transferência:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(criadopor)} valores copiados.")
                except pyperclip.PyperclipException:
                    print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                    print("Dados extraídos e copiados:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(criadopor)} valores copiados.")
            else:
                print_log_sap("Nenhum dado encontrado na coluna I.")

    print_log_sap("⚙ Iniciando extração USR21")
    atualizar_status("⚙ Iniciando extração USR21")
    atualizar_barra_progresso(0.64)
    app.update() 

# ==== EXTRAÇÃO USR21  ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "USR21"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # (Opcional) Carregar variante, se necessário
    # session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").select()
    # session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "/EXTPO USR21"
    # session.findById("wnd[1]/tbar[0]/btn[0]").press()
    # time.sleep(0.5)

    # Abrir seleção múltipla
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    time.sleep(0.5)

    # Colar da área de transferência
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Executar
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Abrir menu de exportação
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Selecionar formato de exportação
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()

    # Confirmar formato
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    print_log_sap("💾 Salvando tabela USR21...")
    # Atualizar status e progresso
    atualizar_status("💾 Salvando tabela USR21...")
    atualizar_barra_progresso(0.68)
    app.update()

    # Definir nome do arquivo com data e hora
    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"CDHDR_USR21_{data_atual}.XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo

    # Confirmar exportação
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Tentar voltar à tela anterior
    try:
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    # Finalizar status e barra
    print_log_sap("✅ Dados USR21 referente CDHDR extraídos e salvos com sucesso")
    atualizar_status("✅ Dados USR21 referente CDHDR extraídos e salvos com sucesso")
    atualizar_barra_progresso(0.70)
    app.update()

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("CDHDR_USR21_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")


# ==== Manipulando dados para ADR6   ====

    def localizar_arquivo_mais_recente():
            pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
            arquivos = list(pasta.glob("CDHDR_USR21_*.xlsx"))

            if not arquivos:
                print_log_sap("Nenhum arquivo encontrado para manipluar dados para ADR6.")
                return None

            return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_dados_coluna_m(arquivo_xlsx):
    # Usando pandas para ler o arquivo Excel
        df = pd.read_excel(arquivo_xlsx)

        # Acessando a coluna M (a 13ª coluna, que tem o índice 12)
        if df.shape[1] > 2:  # Verifica se a coluna M (índice 12) existe
            coluna_m = df.iloc[4:, 2]  # Começando da linha 6 (índice 5) e acessando a coluna M (índice 12)

            # Limpando os dados (removendo valores nulos e espaços extras)
            dados_coluna_m = coluna_m.dropna().apply(lambda x: str(x).strip()).tolist()

            return dados_coluna_m
        else:
            print_log_sap("A coluna G não foi encontrada no arquivo.")
            return None

# Execução principal
    arquivo = localizar_arquivo_mais_recente()
    if arquivo:
            print_log_sap(f"Arquivo encontrado: {arquivo}")
            criadopor = obter_dados_coluna_m(arquivo)

            if criadopor:
                # Juntar os valores da coluna I com \r\n (quebra de linha para o SAP)
                texto_para_copiar = '\r\n'.join(criadopor)

        # Copiar para a área de transferência
                try:
                    pyperclip.copy(texto_para_copiar)
                    print("\nDados extraídos e copiados para a área de transferência:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(criadopor)} valores copiados.")
                except pyperclip.PyperclipException:
                    print_log_sap("\nNão foi possível copiar para a área de transferência. Certifique-se de ter o 'xclip' (Linux) ou 'clip' (Windows) instalado.")
                    print("Dados extraídos e copiados:")
                    print(texto_para_copiar)
                    print(f"\nTotal de {len(criadopor)} valores copiados.")
            else:
                print_log_sap("Nenhum dado encontrado na coluna I.")

    print_log_sap("⚙ Iniciando extração ADR6 referente CDHDR")
    atualizar_status("⚙ Iniciando extração ADR6 referente CDHDR")
    atualizar_barra_progresso(0.75)
    app.update()

# ==== EXTRAÇÃO ADR6 ====

    session.StartTransaction("SE16N")
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "ADR6"
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""  # Sem limite de linhas
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Abrir seleção múltipla (linha 3 da seleção - coluna 5)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").press()
    time.sleep(0.5)

    # Colar dados da área de transferência
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Executar a transação
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Iniciar exportação
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

    # Selecionar formato .XLS (Spreadsheet)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print_log_sap("💾 Salvando tabela ADR6...")
    # Atualizar status e progresso
    atualizar_status("💾 Salvando tabela ADR6...")
    atualizar_barra_progresso(0.80)
    app.update()

    # Gerar nome do arquivo com data/hora
    data_atual = datetime.datetime.now().strftime("%d_%m_%H_%M")
    nome_arquivo = f"CDHDR_ADR6_{data_atual}.XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Tentar retornar à tela anterior
    try:
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
    except Exception as e:
        atualizar_status(f"❌ Erro inesperado: {e}")
        print_log_sap(f"❌ Erro inesperado: {e}")

    # Finalizar status
    print_log_sap("✅ Dados ADR6 referente CDHDR extraídos e salvos com sucesso")
    atualizar_status("✅ Dados ADR6 referente CDHDR extraídos e salvos com sucesso")
    atualizar_barra_progresso(0.83)
    app.update()

# ==== Função para converter .xls para .xlsx   ====

    def converter_xls_para_xlsx(caminho_arquivo_xls):
        # Verificar se o arquivo existe
                if not os.path.exists(caminho_arquivo_xls):
                    print_log_sap(f"O arquivo {caminho_arquivo_xls} não foi encontrado.")
                    return None

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = False  # Não mostrar a interface do Excel
                try:
            # Abrir o arquivo diretamente no modo de fundo sem exibir
                    wb = excel.Workbooks.Open(caminho_arquivo_xls, ReadOnly=True)  # Modo leitura
                    novo_caminho = str(Path(caminho_arquivo_xls).with_suffix(".xlsx"))
                    wb.SaveAs(novo_caminho, FileFormat=51)  # 51 = formato xlsx
                    wb.Close()
                    print_log_sap(f"Arquivo convertido com sucesso: {novo_caminho}")
                    return novo_caminho
                except Exception as e:
                    print_log_sap("Erro ao converter:", e)
                    return None
                finally:
                    excel.Quit()

    # Função para encontrar o arquivo .xls mais recente na pasta especificada
    def encontrar_arquivo_export_mais_recente():
                pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
                arquivos = list(pasta.glob("CDHDR_ADR6_*.xls"))
        
                if not arquivos:
                    print_log_sap("Nenhum arquivo .xls encontrado na pasta.")
                    return None
        
        # Ordenar os arquivos por data de modificação (mais recente primeiro)
                arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                return arquivo_mais_recente

    # Executar a conversão
    arquivo_xls = encontrar_arquivo_export_mais_recente()

    if arquivo_xls:
                print_log_sap(f"Arquivo encontrado: {arquivo_xls}")
                arquivo_convertido = converter_xls_para_xlsx(arquivo_xls)
                if arquivo_convertido:
                    print_log_sap(f"Arquivo convertido: {arquivo_convertido}")

# ==== MANIPULANDO DADOS MM03 ====
    print_log_sap("⏳ Manipulando dados MM03 referente CDHDR")
    atualizar_status("⏳ Manipulando dados MM03 referente CDHDR")
    atualizar_barra_progresso(0.85)
    app.update()

    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("CDHDR_MARA_*.xlsx"))
        if not arquivos:
            print_log_sap("Nenhum arquivo encontrado para manipular dados para MM03.")
            return None
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_docmateriais(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        return [str(c.value) for c in sheet['B'][5:] if c.value is not None]

    def existe(session, id):
        try:
            session.findById(id)
            return True
        except:
            return False

    def select_gui_table_row_by_text(session, field_text, column_index):
        tabela = session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")
        for i in range(tabela.RowCount):
            try:
                celula = session.findById(f"wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[{column_index},{i}]")
                texto = celula.text.strip()
                if texto.upper() == field_text.upper():
                    tabela.getAbsoluteRow(i).selected = True
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    return
            except:
                pass
        raise Exception(f"Texto '{field_text}' não encontrado na coluna {column_index}")

    def salvar_resultados(lista_com_texto, lista_sem_texto, pasta_destino):
        # Obtendo data e hora no formato dd_mm_hh_mm_ss
        timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")
        
        # Caminhos dos arquivos
        path_com = pasta_destino / f"CDHDR_mm03_com_texto_{timestamp}.txt"
        path_sem = pasta_destino / f"CDHDR_mm03_sem_texto_{timestamp}.txt"
        
        # Salvando como Excel
        df = pd.DataFrame(lista_com_texto)
        df.to_excel(path_com, index=False, header=False, engine="openpyxl")

        # Salvando como TXT
        with open(path_sem, "w", encoding="utf-8") as file:
            file.write("\n".join(lista_sem_texto))

        print_log_sap(f"Arquivos salvos:\nExcel: {path_com}\nTXT: {path_sem}")

        with open(path_com, "w", encoding="utf-8") as f_com:
            for numero, texto in lista_com_texto:
                texto_linha_unica = texto.replace("\n", " ").replace("\r", "").strip()
                f_com.write(f"{numero} | {texto_linha_unica}\n")

        with open(path_sem, "w", encoding="utf-8") as f_sem:
            for numero in lista_sem_texto:
                f_sem.write(f"{numero}\n")

    # --- Execução principal ---

    arquivo = localizar_arquivo_mais_recente()
    if not arquivo:
        print_log_sap("Arquivo de origem não encontrado.")
        exit()

    docmateriais = obter_docmateriais(arquivo)
    if not docmateriais:
        print_log_sap("Nenhum número encontrado na coluna C.")
        exit()

    print_log_sap("⚙ Iniciando extrações MM03 referente CDHDR")
    atualizar_status("⚙ Iniciando extrações MM03 referente CDHDR")
    atualizar_barra_progresso(0.90)
    app.update() 

    com_texto = []
    sem_texto = []
    pasta_destino = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"

    for numero in docmateriais:
        try:
            # Transação MM03
            session.findById("wnd[0]/tbar[0]/okcd").text = "MM03"
            session.findById("wnd[0]/tbar[0]/btn[0]").press()

            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = numero
            session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = len(numero)
            session.findById("wnd[0]/tbar[0]/btn[0]").press()

            select_gui_table_row_by_text(session, field_text="Dados básicos 1", column_index=0)
            session.findById("wnd[0]/tbar[1]/btn[30]").press()
            session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05").select()
            time.sleep(1)

            texto_completo = ""

            try:
                control = session.findById(
                    "wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/"
                    "subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell"
                )

                if hasattr(control, "RowCount"):
                    linhas = [control.GetCellValue(i, 0) for i in range(control.RowCount)]
                    texto_completo = "\n".join(linhas)
                else:
                    try:
                        texto_completo = control.Text
                    except Exception:
                        texto_completo = ""

            except:
                try:
                    texto_completo = session.findById(
                        "wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/"
                        "subSUB2:SAPLMGD1:2031/txtRSTXT-TXLINE"
                    ).Text
                except:
                    texto_completo = ""

            if texto_completo.strip():
                com_texto.append((numero, texto_completo))
                print(f"Texto capturado para {numero}")
            else:
                sem_texto.append(numero)
                print(f"Sem texto para {numero}")

            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()

        except Exception as e:
            print_log_sap(f"Erro ao processar {numero}: {e}")
            sem_texto.append(numero)
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except:
                pass

    salvar_resultados(com_texto, sem_texto, pasta_destino)
    print_log_sap("✅ Dados MM03 extraidos e salvos com sucesso")

    atualizar_status("✅ Dados MM03 extraidos e salvos com sucesso")
    atualizar_barra_progresso(0.94)
    app.update()  

# ==== MANIPULANDO DADOS ME23N ====
    print("⏳ Manipulando dados ME23N referente CDHDR")

    def conectar_sap():
        try:
            SapGuiAuto = win32.GetObject("SAPGUI")
            if not SapGuiAuto:
                print("SAP GUI não está em execução.")
                return None
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            return session
        except Exception as e:
            print(f"Erro ao conectar com o SAP GUI: {e}")
            return None

    def localizar_arquivo_mais_recente():
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        arquivos = list(pasta.glob("CDHDR_EKKO_*.xlsx"))
        if not arquivos:
            print("Nenhum arquivo encontrado para manipulação de dados para ME23N.")
            return None
        return max(arquivos, key=lambda f: f.stat().st_mtime)

    def obter_docmateriais(arquivo_xlsx):
        wb = openpyxl.load_workbook(arquivo_xlsx)
        sheet = wb.active
        return [str(c.value) for c in sheet['B'][5:] if c.value is not None]

    def extrair_texto_texto_de_cabecalho(session):
        try:
            time.sleep(1)
            editor_id = (
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/"
                "subSUB1:SAPLMEVIEWS:1100/"
                "subSUB2:SAPLMEVIEWS:1200/"
                "subSUB1:SAPLMEGUI:1102/"
                "tabsHEADER_DETAIL/tabpTABHDT3/"
                "ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/"
                "subTEXTS:SAPLMMTE:0100/"
                "subEDITOR:SAPLMMTE:0101/"
                "cntlTEXT_EDITOR_0101/shellcont/shell"
            )
            editor = session.findById(editor_id)
            try:
                texto = editor.text
            except:
                texto = editor.getProperty("Text")
            return texto.strip()
        except Exception:
            return ""

    def salvar_resultados(lista_com_texto, lista_sem_texto, pasta_destino):
        timestamp = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")
        path_com = Path(pasta_destino) / f"CDHDR_me23n_com_texto_{timestamp}.txt"
        path_sem = Path(pasta_destino) / f"CDHDR_me23n_sem_texto_{timestamp}.txt"

        with open(path_com, "w", encoding="utf-8") as f_com:
            for numero, texto in lista_com_texto:
                texto_linha_unica = texto.replace("\n", " ").replace("\r", "").strip()
                f_com.write(f"{numero} | {texto_linha_unica}\n")

        with open(path_sem, "w", encoding="utf-8") as f_sem:
            for numero in lista_sem_texto:
                f_sem.write(f"{numero}\n")

        print(f"\nArquivos salvos:")
        print(f"✓ Pedidos COM texto: {path_com}")
        print(f"✓ Pedidos SEM texto: {path_sem}")

    print("⚙ Iniciando extrações ME23N")

    def executar_automacao_ME23N(session):
        arquivo = localizar_arquivo_mais_recente()
        if not arquivo:
            return

        docmateriais = obter_docmateriais(arquivo)

        if not docmateriais:
            print("Nenhum número de pedido encontrado.")
            return

        pasta_base = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        os.makedirs(pasta_base, exist_ok=True)

        lista_com_texto = []
        lista_sem_texto = []

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME23N"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        time.sleep(2)

        try:
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/"
                            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/"
                            "tabsHEADER_DETAIL/tabpTABHDT3").select()
            time.sleep(1)
        except Exception as e:
            print("❌ Erro ao selecionar aba 'Textos':", e)
            return

        for pedido in docmateriais:
            try:
                session.findById("wnd[0]/tbar[1]/btn[17]").press()
                time.sleep(0.5)
                session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = pedido
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                time.sleep(2)

                texto = extrair_texto_texto_de_cabecalho(session)

                if texto:
                    lista_com_texto.append((pedido, texto))
                else:
                    lista_sem_texto.append(pedido)

            except Exception as e:
                print(f"Erro com o pedido {pedido}: {str(e)}")
                lista_sem_texto.append(f"{pedido} (erro)")
                continue

        salvar_resultados(lista_com_texto, lista_sem_texto, pasta_base)

    # ====== INÍCIO DO SCRIPT ======
    session = conectar_sap()
    if session:
        executar_automacao_ME23N(session)

        # Apagar arquivos .xls após execução
        pasta = Path.home() / "OneDrive - Accenture" / "Documents" / "SAP" / "SAP GUI"
        for arquivo in pasta.glob("*.xls"):
            try:
                os.remove(arquivo)
                print(f"🗑 Arquivo removido: {arquivo.name}")
            except Exception as e:
                print(f"❌ Erro ao remover {arquivo.name}: {e}")
    else:
        print("❌ Não foi possível conectar ao SAP.")


   # ---- CIF ----
    print_log_sap("📤 Adicionando IncTm e Incotm.2\n")

    # Função para verificar se uma célula está realmente vazia
    def is_realmente_vazio(valor):
        if valor is None:
            return True
        valor_str = str(valor)
        valor_str = unicodedata.normalize('NFKC', valor_str).strip()
        return valor_str == ""

    # Obter nome do usuário e montar o caminho
    usuario = getpass.getuser()
    pasta = Path(f"C:/Users/{usuario}/OneDrive - Accenture/Documents/SAP/SAP GUI")

    # Localizar o arquivo CDHDR_EKKO_
    arquivos_ekko = list(pasta.glob("CDHDR_EKKO_*.xlsx"))
    if not arquivos_ekko:
        print_log_sap("Arquivo CDHDR_EKKO_ não encontrado.")
    else:
        caminho_arquivo = arquivos_ekko[0]
        wb = openpyxl.load_workbook(caminho_arquivo)
        ws = wb.active

        # Preencher coluna F (coluna 6) a partir da linha 6 com "CIF" se estiver vazia
        for row in ws.iter_rows(min_row=6, min_col=6, max_col=6):
            cell = row[0]
            if is_realmente_vazio(cell.value):
                cell.value = "CIF"

        # Preencher coluna M (coluna 13) a partir da linha 6 com "Custo, seguro & frete" se estiver vazia
        for row in ws.iter_rows(min_row=6, min_col=13, max_col=13):
            cell = row[0]
            if is_realmente_vazio(cell.value):
                cell.value = "Custo, seguro & frete"

        # Salvar alterações
        wb.save(caminho_arquivo)
        print_log_sap(f"Arquivo atualizado com sucesso: {caminho_arquivo}")


    # ---- GRUPO MERCADORIA -----
    print_log_sap("📤 Convertendo Grupo Mercadoria\n")

    # --- Caminho da pasta SAP ---
    pasta_sap = os.path.expanduser(r"~\OneDrive - Accenture\Documents\SAP\SAP GUI")
    if not os.path.exists(pasta_sap):
        os.makedirs(pasta_sap)

    # --- Localiza o arquivo EKPO mais recente ---
    arquivos_ekpo = [f for f in os.listdir(pasta_sap) if f.startswith("CDHDR_EKPO_") and f.endswith(".xlsx")]
    arquivos_ekpo.sort(key=lambda x: os.path.getmtime(os.path.join(pasta_sap, x)), reverse=True)
    arquivo_ekpo = os.path.join(pasta_sap, arquivos_ekpo[0]) if arquivos_ekpo else None

    # --- Caminho do arquivo GrupoMercadoria ---
    arquivo_grupo = os.path.join(pasta_sap, "GrupoMercadoria.xlsx")

    # --- Verificação de existência dos arquivos ---
    if not arquivo_ekpo or not os.path.exists(arquivo_grupo):
        print_log_sap("Arquivo EKPO ou GrupoMercadoria não encontrado.")
        exit()

    # --- Carrega os dados do GrupoMercadoria ---
    wb_grupo = openpyxl.load_workbook(arquivo_grupo)
    ws_grupo = wb_grupo.active

    # Cria dicionário: {codigo_grupo: nome_grupo}
    mapa_grupo = {}
    for row in ws_grupo.iter_rows(min_row=2, values_only=True):  # assumindo cabeçalho
        if row[0] is not None and row[1] is not None and row[2] is not None:
            mapa_grupo[str(row[0]).strip()] = str(row[2]).strip()

    # --- Carrega o arquivo EKPO ---
    wb_ekpo = openpyxl.load_workbook(arquivo_ekpo)
    ws_ekpo = wb_ekpo.active

    # --- Substitui valores da coluna Q (17ª) a partir da linha 6 ---
    for row in ws_ekpo.iter_rows(min_row=6, min_col=17, max_col=17):
        cell = row[0]
        valor_original = str(cell.value).strip() if cell.value is not None else ""
        if valor_original in mapa_grupo:
            cell.value = mapa_grupo[valor_original]

    # --- Salva com nome dinâmico baseado na data e hora ---
    agora = datetime.datetime.now().strftime("%d_%m_%H_%M_%S")
    nome_arquivo = f"CDHDR_EKPO_{agora}.xlsx"
    caminho_saida = os.path.join(pasta_sap, nome_arquivo)
    wb_ekpo.save(caminho_saida)

    print_log_sap(f"[{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Arquivo salvo como: {caminho_saida}")


# Configuração da interface
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Automação SAP")
app.geometry("500x600")
app.resizable(False, False)

frame_inicial = ctk.CTkFrame(app, width=880, height=500, fg_color="#dbe3fa")
frame_inicial.pack(fill="both", expand=True)
frame_inicial.grid_propagate(False)

frame_logo = ctk.CTkFrame(frame_inicial, fg_color="transparent")
frame_logo.place(relx=0.5, rely=0.25, anchor="center")

logo_img = ctk.CTkImage(light_image=Image.open(recurso_caminho("logo_zzzz.png")), size=(160, 100))
ctk.CTkLabel(frame_logo, image=logo_img, text="").pack(pady=15)
            
app.iconbitmap(recurso_caminho("iconaccenture.ico"))

frame_botao = ctk.CTkFrame(frame_inicial, fg_color="#dbe3fa")
frame_botao.place(relx=0.5, rely=0.6, anchor="center")

ctk.CTkButton(
    frame_botao, text="SAP", width=120, height=40,
    fg_color="#2c2c3c", hover_color="#0098d1",
    font=ctk.CTkFont("Segoe UI", 16, "bold"),
    corner_radius=30, command=exibir_sap
).pack(side="left", padx=20)

ctk.CTkButton(
    frame_botao,
    text="Api ",
    width=120,
    height=40,
    fg_color="#ff5c5c",
    hover_color="#d03f3f",
    font=ctk.CTkFont("Segoe UI", 16, "bold"),
    corner_radius=30,
    command=exibir_api  # chama a função ao clicar
).pack(side="right", padx=20)


# Rodapé
rodape = Label(
    app,
    text="Powered by Accenture    v.1.1.0",
    bg="white",
    font=("Arial", 12),
    fg="gray",
    height=2
)
rodape.pack(side="bottom", fill="x", ipady=6)

app.mainloop()
