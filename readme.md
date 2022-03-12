<h2 align='center'>Automação - Tratamento de dados e envio por email
(Automation - Data handling and Mailing)
</h2>
<br>

### Descrição (Description)

<p>Script em Python para automação de processo de obtenção de dados a partir de um banco de dados, tratamento desses dados, composição e envio de email com informações tratadas.</p>

<p>Automation script that imports from a DB, treat data and send them by email.</p>
<br>

### Funcionalidades do projeto (Project Functionalities)

* Acessa e baixa arquivo Excel em um Google Drive através do navegador
* Acessa os dados do arquivo e calcula a quantidade de items vendidos e total faturado
* Acessa email do usuário através do navegador, compõe email com os dados calculados e envia

<br>

* Access and download Excel file from Google Drive
* Access file data and sums sales and total income
* Access user email through browser and email the data
<br>


### Tecnologias utilizadas (Dependencies)

* Pyautogui
* Pyperclip
* Pandas
* Openpyxl
* Numpy
<br>


### Rodando o projeto (Running the project)

* Esse script não funciona em segundo plano, o usuário precisa aguarda-lo finalizar para que não haja conflito nas funções.
* O processo de automação através do Pyautogui controla os inputs do mouse e teclado para simular o acesso humano. Para isso, ele depende das coordenadas do mouse sobre os pontos a serem clicados. O script get_mouse_location.py, quando executado, aguarda 5 segundos, captura e imprime na tela as coordenadas (x e y) da posição do mouse. Execute o get_mouse_location.py sempre que necessário capturar novas coordenadas e atualize as coordenadas de cada clique nas funções pyautogui.click, modificando também os parâmetros click e button para definir a quantidade de cliques e qual botão do mouse utilizar.
* Atualize a url de acesso ao arquivo de banco de dados na função pyperclip.copy
* Atualize o nome do arquivo para que o pandas abra o arquivo
* Trate os dados conforme necessário
* Atualize a url de acesso ao email,  email de destino, título e mensagem nas respectivas funções pyperclip.copy
<br>

* This script doesn't run in background. User needs to wait for it's conclusion so there are no problems.
* The automation process through Pyautogui controls mouse and keyboard inputs to simulate human access. To do that it requires the mouse coordinates when hoving the elements to be clicked. The script get_mouse_location.py waits 5 seconds then captures and display the mouse position coordinates. Run get_mouse_location.py as many times as required to get new coordinates and update each click coordinate on their pyautogui.click functions, also updating the click (how many clicks) and button (which mouse button) parameters.
* Update the DB access url on pyperclip.copy
* Update the file name on pandas.read_...
* Handle the data according to your needs
* Update the email url, email address, subject and message on respectives pyperclip.copy
