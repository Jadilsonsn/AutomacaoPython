<h1>Automação com PyAutoGUI para Preenchimento de Planilhas</h1>

<h3>Descrição do Projeto</h3>
Neste projeto utilizei a biblioteca PyAutoGUI para realizar a automação do preenchimento de planilhas Excel. Ele foi desenvolvido pensando em otimizar tarefas manuais, como o cadastro de produtos ou transferência de dados entre planilhas.

A automação lê informações de uma planilha de origem e, de forma automatizada, preenche os campos correspondentes em outra planilha. Isso elimina a necessidade de fazer essas operações manualmente, reduzindo erros e economizando tempo.

<h3>Requisitos:</h3>
Python 3.7 ou superior<br></br>
Bibliotecas necessárias:<br></br>
- pyautogui <br></br>
- time (Essa biblioteca já vem instalada) <br></br>

Você pode instalar o PyAutoGui com o seguinte comando:<br></br>
<strong>pip install pyautogui</strong> 

📚  Confira a documentação oficial do [PyAutoGUI](https://pyautogui.readthedocs.io/en/latest/) para mais informações. 

<h3>Dica Extra:</h3>
O comando pyautogui.PAUSE = 2.5 é uma forma prática de definir um intervalo padrão de 2,5 segundos entre a execução de cada comando do PyAutoGUI. Isso elimina a necessidade de adicionar várias chamadas ao time.sleep no código, tornando-o mais limpo e eficiente.
