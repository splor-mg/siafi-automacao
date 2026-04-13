# Automação SIAFI com py3270

Automação de operações no SIAFI (Sistema Integrado de Administração Financeira) utilizando Python e emulação de terminal TN3270, desenvolvido para a Cidade Administrativa de Minas Gerais.

---

## Sobre o projeto

O SIAFI roda em um terminal TN3270, acessado localmente pelo emulador **pw3270**. Este projeto utiliza a biblioteca **py3270** para controlar o emulador via Python, permitindo automatizar operações repetitivas como aprovação e anulação de cotas orçamentárias, eliminando a necessidade de interação manual com o terminal.

---

## Funcionalidades

- Login automático no SIAFI com tratamento de telas de aviso
- Leitura de planilha Excel com os dados das operações
- Aprovação de cotas orçamentárias (global e amarrada)
- Anulação de cotas orçamentárias
- Captura do retorno do SIAFI (sucesso ou erro) para cada operação

---

## Pré-requisitos

### Sistema operacional

O projeto roda no Linux **Ubuntu**, podendo ser utilizado WSL (Windows Subsystem for Linux), instalado diretamente no Windows.

### Dependências do sistema

Com o Ubuntu aberto, instale as dependências necessárias:

```bash
# Atualiza a lista de pacotes
sudo apt update

# Instala o emulador x3270 (necessário para o py3270 funcionar)
sudo apt install x3270

# Instala o s3270 (versão sem interface gráfica, para rodar em background)
sudo apt install s3270

# Instala o xdotool (controle de janelas, opcional para modo visível)
sudo apt install xdotool

# Instala o Python e o pip
sudo apt install python3 python3-pip python3-venv
```

---

## Instalação do projeto

### 1. Clone ou copie os arquivos para o Ubuntu

```bash
# Cria a pasta do projeto
clone git@github.com:splor-mg/siafi-automacao.git
cd siafi_automacao
```

### 2. Crie e ative o ambiente virtual

```bash
# Cria o ambiente virtual
python3 -m venv venv

# Ativa o ambiente virtual
source venv/bin/activate
```

### 3. Instale as dependências Python

```bash
pip install -r requirements.txt
```

---

## Como usar

### Ativar o ambiente virtual (sempre que abrir o terminal)

```bash
cd ~/siafi_automacao
source venv/bin/activate
```

### Executar o script

```bash
python teste_py3270.py
```

---

## Estrutura do projeto

```
siafi_automacao/
│
├── venv/                  # Ambiente virtual Python
├── teste_py3270.py        # Script principal de automação
├── planilha.xlsx          # Planilha com os dados das operações (a implementar)
└── README.md              # Este arquivo
```

---

## Fluxo de automação

```
1. Login no SIAFI
      ↓
2. Leitura da planilha Excel
      ↓
3. Para cada linha da planilha:
   ├── Navega até a transação de cotas
   ├── Preenche os campos (mês, fonte, UO, grupo de despesa...)
   ├── Executa aprovação ou anulação
   └── Captura o retorno do SIAFI (sucesso ou erro)
      ↓
4. Gera relatório de resultados
```

---

## Principais bibliotecas utilizadas

| Biblioteca | Descrição | Instalação |
|------------|-----------|------------|
| `py3270` | Interface Python para o emulador x3270/s3270 | `pip install py3270` |
| `openpyxl` | Leitura e escrita de planilhas Excel (.xlsx) | `pip install openpyxl` |
| `time` | Controle de pausas entre operações | Nativa do Python |

---

## Funções principais do py3270

| Função | Descrição |
|--------|-----------|
| `Emulator(visible=True)` | Cria o emulador (visível ou invisível) |
| `em.connect('host')` | Conecta ao servidor SIAFI |
| `em.wait_for_field()` | Aguarda a tela estar pronta para digitação |
| `em.fill_field(linha, col, texto, tam)` | Preenche um campo na tela |
| `em.send_enter()` | Pressiona Enter |
| `em.send_pf(n)` | Pressiona tecla de função (F1–F24) |
| `em.string_found(linha, col, texto)` | Verifica se um texto está na posição indicada |
| `em.string_get(linha, col, tam)` | Lê um trecho de texto da tela |
| `em.terminate()` | Encerra o emulador |

---

## Observações

- O emulador `x3270` precisa estar instalado e disponível no PATH do sistema para o py3270 funcionar
- Em modo `visible=True` o script abre a janela gráfica do terminal — útil para testes e depuração
- Em modo `visible=False` o script roda em segundo plano usando o `s3270` — recomendado para produção
- Sempre que o SIAFI demorar para responder, utilize `time.sleep()` antes do `wait_for_field()` para evitar o erro `Keyboard locked`

---

## Referências

- [py3270 no PyPI](https://pypi.org/project/py3270/)
- [pw3270 — Perry Werneck](https://github.com/PerryWerneck/pw3270)
- [x3270 — IBM 3270 Terminal Emulator](http://x3270.bgp.nu/)