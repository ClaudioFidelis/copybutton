import tkinter as tk
import pywinauto
from pywinauto.application import Application
import configparser
import re
import ait
import pyautogui
import time
import pygetwindow as gw
import pyperclip
import win32clipboard
import subprocess

# Variáveis para controlar o arrastar e soltar
x = 0
y = 0


# Configuração inicial
config = configparser.ConfigParser()
config['window'] = {'x': '0', 'y': '0'}

def limpar_area_de_transferencia_powershell():
    subprocess.run("cmd /c \"echo off | clip\"", shell=True)

def limpar_area_de_transferencia():
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        print("Área de transferência limpa!")
    except Exception as e:
        print(f"Erro ao limpar a área de transferência: {e}")

def load_config():
    try:
        config.read('config.ini')
        return int(config['window']['x']), int(config['window']['y'])
    except (configparser.Error, ValueError):
        return None, None

def save_config(x, y):
    with open('config.ini', 'w') as configfile:
        config['window']['x'] = str(x)
        config['window']['y'] = str(y)
        config.write(configfile)


def start_move(event):
    global x, y
    x = event.x
    y = event.y

def move_window(event):
    global x, y
    root.geometry("+%d+%d" % (event.x_root - x, event.y_root - y))

# Salvar a posição da janela ao fechar
def on_closing():
    x = root.winfo_x()
    y = root.winfo_y()
    save_config(x, y)
    root.destroy()

def on_click():

    contatosolicitante = "Boa tarde, tudo bem? Faço parte da equipe do Service Desk e o motivo do meu contato é o chamado, " \
    "registro – descricaoresumida, que se encontra em andamento. Nossa equipe está comprometida em acompanhar " \
    "o seu chamado até a sua finalização.\n" \
    "Gostaria de informar que serei o ponto de contato responsável pelo seu chamado. Qualquer informação referente " \
    "ao chamado pode ser fornecida por aqui, para agilização da resolução do problema enfrentado. Estamos em contato " \
    "com a equipe técnica responsável pelo atendimento e qualquer nova informação sobre o andamento do chamado será repassada.\n" \
    "Desde já agradeço a sua atenção e retorno."

    contatotecnico = "Boa tarde, tudo bem? \n" \
    "Faço parte da equipe do Service Desk, estamos realizando o acompanhamento de chamados originados por nós e que ainda estão "\
    "pendentes nas equipes de suporte.  Estou com a análise de atualização do chamado XXXXXX, atribuído ao seu nome na mesa XXXXXX.\n"\
    "Tem alguma previsão para o atendimento?\n"\
    "Desde já agradeço a sua atenção e retorno."

    informacaotecnico = "Foi realizada uma tentativa de contato com o técnico XXXXXX atribuído ao chamado, através do Microsoft Teams.\n" \
    "Aguardando retorno."

    contatosemtecnico = "Boa tarde, tudo bem?\n" \
    "Faço parte da equipe do Service Desk, estamos realizando o acompanhamento de chamados originados por nós e que ainda estão "\
    "pendentes nas equipes de suporte. Estou com a atualização do chamado XXXXXXX, atribuído a mesa XXXXXXX, porém sem técnico "\
    "atribuído, e com isso, sem atualização ou tratativa.\n"\
    "Vi que faz parte do grupo de membros da mesa. Consegue me ajudar ou me orientar sobre quem seria o responsável pelo roteamento de chamados?\n"\
    "Desde já agradeço a sua atenção e retorno."

    informacaosemtecnico = "Tentativa de contato com técnicos da mesa, através do Microsoft Teams:\n\n"

    #limpar_area_de_transferencia_powershell()
    #time.sleep(1)

    pyperclip.copy(informacaosemtecnico)
    time.sleep(0.3)
    pyperclip.copy(contatosemtecnico)
    time.sleep(0.3)
    pyperclip.copy(informacaotecnico)
    time.sleep(0.3)
    pyperclip.copy(contatotecnico)
    time.sleep(0.3)
    pyperclip.copy(contatosolicitante)
    time.sleep(0.3)
    pyautogui.click(3260, 595)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.click(3260, 519)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.click(3260, 484)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.click(2420, 791)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.click(2420, 403)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.click(2420, 211)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')



# Carregar a configuração da posição da janela
x, y = load_config()

# Se a configuração não existir, centralizar a janela
if x is None or y is None:
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = screen_width // 2 - width // 2
    y = screen_height // 2 - height // 2

# Definir o tamanho da janela (ajuste conforme necessário)
width = 90
height = 40

# Cria a janela principal
root = tk.Tk()
root.overrideredirect(True)
root.wm_attributes('-topmost', True)
root.geometry(f"{width}x{height}+{x}+{y}")

# Cria um botão que ocupa toda a janela
button = tk.Button(root, text="Copy", command=on_click)
button.pack(fill='both', expand=True)

# Binds os eventos para arrastar e soltar
root.bind('<Button-1>', start_move)
root.bind('<B1-Motion>', move_window)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()