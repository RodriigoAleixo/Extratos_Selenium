🤖 Automação de Extração de Extratos - Newcon Web
Este projeto é um script de automação desenvolvido em Python que utiliza as bibliotecas Selenium e PyAutoGUI para acessar o sistema Newcon Web, realizar a busca de cotas de consórcio a partir de uma planilha Excel e salvar os extratos correspondentes em formato PDF.

📋 Funcionalidades
Login Automático: Realiza o acesso ao portal Newcon com credenciais pré-definidas.

Integração com Excel: Lê dados de Grupo, Cota e Nome diretamente de um arquivo .xlsx.

Navegação Inteligente: Percorre os menus do sistema para localizar o extrato de cada cliente.

Impressão via Interface (GUI): Utiliza o PyAutoGUI para interagir com a janela de impressão do sistema operacional e salvar o arquivo com um nome padronizado.

🛠️ Tecnologias Utilizadas
Python (Linguagem principal)

Selenium (Automação Web)

Webdriver Manager (Gerenciamento automático do driver do Chrome)

Openpyxl (Manipulação de arquivos Excel)

PyAutoGUI (Automação de interface gráfica/mouse/teclado)

⚙️ Pré-requisitos
Antes de executar o script, você precisará ter instalado:

Python 3.x

Google Chrome instalado.

As bibliotecas necessárias. Você pode instalá-las via terminal com o comando:

Bash
pip install selenium webdriver-manager openpyxl pyautogui
Arquivo de Dados: Uma planilha chamada extrato.xlsx na mesma pasta do script, contendo as colunas:

Coluna A: Grupo

Coluna B: Cota

Coluna C: Nome do Cliente

Referência Visual: Uma imagem chamada imprimir-2.png que represente o botão ou ícone de impressão que o PyAutoGUI deve clicar.

🚀 Como Usar
Configuração de Credenciais: No código, substitua os campos 'USUÁRIO' e 'SENHA' pelas suas credenciais de acesso ao Newcon.

Preparação do Ambiente: Certifique-se de que a janela do navegador e a de impressão não serão obstruídas, pois o PyAutoGUI depende da visão da tela.

Execução:

Bash
python nome_do_seu_arquivo.py
Processo: O script abrirá o Chrome, fará login, lerá a planilha linha por linha, gerará o extrato e o salvará com a nomenclatura: Grupo-Cota - Extrato 10-07-2025.

⚠️ Observações Importantes
Pausas (time.sleep): O código possui diversos intervalos de tempo para garantir que o sistema carregue. Dependendo da velocidade da sua internet, esses tempos podem precisar de ajuste.

Resolução de Tela: O PyAutoGUI utiliza reconhecimento de imagem (imprimir-2.png). Se você mudar de monitor ou resolução, pode ser necessário tirar um novo print do botão.
