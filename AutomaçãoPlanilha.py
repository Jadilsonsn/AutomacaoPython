import pyautogui 
import time 


# ABRE O DOCUMENDO 
pyautogui.press("Win")
time.sleep(2)
pyautogui.write("Google Chorme")
time.sleep(2)
pyautogui.press("Enter")
time.sleep(3)
pyautogui.write("login one drive")
time.sleep(4)
pyautogui.press("enter")
time.sleep(3)
pyautogui.write("@gmail.com ")
time.sleep(3)
pyautogui.press("Enter") 
time.sleep(3)


# FUNÇÕES MOUSE / ENCONTRA POSIÇÕES COM BASE NO PIXEL DA TELA. 
pyautogui.click(x=122, y=451)
time.sleep(5)
pyautogui.click(x=694, y=725)
time.sleep(5)
pyautogui.click(x=673, y=734)
time.sleep(7)
pyautogui.click(x=588, y=509)
time.sleep(5)

#Abrir o arquivo online na area de trabalho 
pyautogui.click(x=1692, y=223)
time.sleep(4)
pyautogui.click(x=1574, y=432)
time.sleep(7)


# Ajustar o arquivo para o tamanho 100% 
pyautogui.click(x=1890, y=988)
time.sleep(3)
pyautogui.click(x=1204, y=535)
time.sleep(2)
pyautogui.click(x=1239, y=745)
time.sleep(2) 

#Garantir que não tenha nenhum dado filtrado  
pyautogui.hotkey('ctrl','shift','l')
time.sleep(1)
pyautogui.hotkey('ctrl','shift','l')

#Ir para celula A1
pyautogui.hotkey('ctrl','left')
pyautogui.hotkey('ctrl','left')
time.sleep(2)
pyautogui.hotkey('ctrl','up')
pyautogui.hotkey('ctrl','up')


# Filtragem da coluna Ponto Focal 
time.sleep(3)
for _ in range(15):
    pyautogui.press("right")

time.sleep(3)
pyautogui.click(x=1484, y=424) 
time.sleep(3)  
pyautogui.click(x=1238, y=649)   
time.sleep(3)
pyautogui.write('filtro')
time.sleep(2)
pyautogui.press('enter')
time.sleep(3)

#Volta para celula de referência A1
pyautogui.hotkey('ctrl','up')
pyautogui.hotkey('ctrl','up')
time.sleep(2)
pyautogui.hotkey('ctrl','left')
pyautogui.hotkey('ctrl','left')


#ABRIR ARQUIVO EXCEL (TABELA AUXILIAR)
pyautogui.press('win')
time.sleep(3)
pyautogui.write('Explorador')
time.sleep(3)
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=120, y=475)
time.sleep(3)
pyautogui.doubleClick(x=654, y=313)
time.sleep(5)
pyautogui.hotkey('ctrl','left')
pyautogui.hotkey('ctrl','left')
time.sleep(2)
pyautogui.hotkey('ctrl','up')
pyautogui.hotkey('ctrl','up')
time.sleep(3)
# Copia uma informação da tabela auxiliar 

pyautogui.press('down')    
pyautogui.hotkey('ctrl','c')
time.sleep(3) 

# Vai para tabela principal e faz uma filtragem de dados 
table = 1
pyautogui.keyDown('alt')
for _ in range(table):
    pyautogui.press('tab') 
time.sleep(2)    
pyautogui.keyUp('alt')

#  Ir a coluna Serviço para selecionar a opção desejada  
time.sleep(5)
for _ in range(5):
    pyautogui.press('right')
time.sleep(2)    
pyautogui.click(x=1257, y=424)
time.sleep(2)
pyautogui.click(x=1008, y=643)
time.sleep(2)
pyautogui.hotkey('ctrl','v')
time.sleep(2)
pyautogui.press('enter') 
time.sleep(2) 


# Na tabela principal, o cursor retornar para celula A1 
pyautogui.hotkey('ctrl','left')
time.sleep(1)
pyautogui.hotkey('ctrl','left')
time.sleep(3)
pyautogui.hotkey('ctrl','up')
time.sleep(1)
pyautogui.hotkey('ctrl','up')
time.sleep(3)


#Retornará a tabela auxiliar - Uma informação será copiada e colada na tabela principal
pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.keyUp('alt')
time.sleep(3)
pyautogui.press('right')
pyautogui.hotkey('ctrl','c')
time.sleep(3)
pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.keyUp('alt')
time.sleep(3)

# Na tabela principal, será filtrado uma informação da coluna indicador 
for _ in range(6):
    pyautogui.press('right')
time.sleep(3)
pyautogui.click(x=1825, y=419) 
time.sleep(5)
pyautogui.click(x=1581, y=645)
time.sleep(3)
pyautogui.hotkey('ctrl','v')
time.sleep(2)
pyautogui.press('enter')
time.sleep(5)

#Na tabela principal, voltaremos  a celula A1
pyautogui.hotkey('ctrl','left')
time.sleep(1)
pyautogui.hotkey('ctrl','left')
time.sleep(2)
pyautogui.hotkey('ctrl','up')
time.sleep(1)
pyautogui.hotkey('ctrl','up')
time.sleep(5) 

# Vai até a coluna medição 
for _ in range(14):
    pyautogui.press('right')
time.sleep(2)
for _ in range(1):
    pyautogui.press('down')
time.sleep(3)    

# volta a tabela auxiliar pega uma informação e cola na tabela principal. 
pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.keyUp('alt') 
time.sleep(2)   
pyautogui.press('right')
time.sleep(2)
pyautogui.hotkey('ctrl','c')
time.sleep(2)
pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.keyUp('alt') 
time.sleep(3)
pyautogui.hotkey('ctrl','v')    
time.sleep(4)


for _ in range(25):
    pyautogui.press('right')
    pyautogui.press('up')
    pyautogui.hotkey('ctrl','shift','l')
    time.sleep(2)
    time.sleep(1) 
    pyautogui.hotkey('ctrl','shift','l')
    time.sleep(2)
    pyautogui.click(x=1484, y=426)
    time.sleep(2)
    pyautogui.click(x=1235, y=646)
    time.sleep(2)
    pyautogui.write('texto no filtro')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)

    pyautogui.hotkey('ctrl','up')
    pyautogui.hotkey('ctrl','up')
    time.sleep(2)
    pyautogui.hotkey('ctrl','left')
    pyautogui.hotkey('ctrl','left')
    time.sleep(3)
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')
    time.sleep(3)
    pyautogui.hotkey('ctrl','left')
    pyautogui.hotkey('ctrl','left')
    time.sleep(1)
    pyautogui.press('down')
    time.sleep(1)
    pyautogui.hotkey('ctrl','c')
    time.sleep(2)
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')
    time.sleep(3)

    for _ in range(5):
        pyautogui.press('right')
    time.sleep(2) 
    pyautogui.click(x=1259, y=425)
    time.sleep(2)
    pyautogui.click(x=1012, y=642)
    pyautogui.hotkey('ctrl','v')
    time.sleep(2)
    pyautogui.press('enter')

    time.sleep(2)
    pyautogui.hotkey('ctrl','up')
    pyautogui.hotkey('ctrl','up')
    time.sleep(2)
    pyautogui.hotkey('ctrl','left')
    pyautogui.hotkey('ctrl','left')
    time.sleep(2)

    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')
    time.sleep(3)
    pyautogui.press('right')
    pyautogui.hotkey('ctrl','c')
    time.sleep(3)
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')
    time.sleep(3)


    for _ in range(6):
        pyautogui.press('right')
    time.sleep(3)
    pyautogui.click(x=1827, y=424) 
    time.sleep(4)
    pyautogui.click(x=1590, y=646)
    time.sleep(3)
    pyautogui.hotkey('ctrl','v')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(3) 


    pyautogui.hotkey('ctrl','left')
    pyautogui.hotkey('ctrl','left')
    time.sleep(2)
    pyautogui.hotkey('ctrl','up')
    pyautogui.hotkey('ctrl','up')
    time.sleep(3)

    for _ in range(14):
        pyautogui.press('right')
    time.sleep(2)
    for _ in range(1):
        pyautogui.press('down')
    time.sleep(3)    
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt') 
    time.sleep(2)   

    pyautogui.press('right')
    time.sleep(2)
    pyautogui.hotkey('ctrl','c')
    time.sleep(2)
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt') 
    time.sleep(3)
    pyautogui.hotkey('ctrl','v')    
    time.sleep(3)











    







    























