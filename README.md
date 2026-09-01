# 📧 E-mail Automático — Envio de Credenciais

Ferramenta com interface gráfica (Tkinter) para envio automático de PDFs de credenciais (SSHD) para colaboradores via e-mail, com cópia automática salva na pasta "Enviados" via IMAP.

## ✨ Funcionalidades

- Envio de e-mails em HTML com logo institucional
- Anexo automático de PDF de credenciais por colaborador
- Leitura de dados a partir de planilha/CSV
- Cópia automática do e-mail enviado salva via IMAP
- Interface gráfica simples (Tkinter)
- Configuração de SMTP e IMAP direto na interface

## 🚀 Como usar

1. Instale as dependências:
```bash
   pip install openpyxl pdfplumber unidecode
```
2. Execute o script:
```bash
   python enviar_credenciais_gui_v4.py
```
3. Preencha os dados de SMTP (e IMAP, opcional) na interface e envie.

## 🛠️ Gerar executável (.exe)

```bash
python -m PyInstaller --onefile --windowed --name EnviarCredenciais enviar_credenciais_gui_v4.py
```

## 📁 Versões

| Arquivo | Descrição |
|---|---|
| `enviar_credenciais_gui.py` | Versão inicial |
| `enviar_credenciais_gui_v2.py` | Melhorias na interface |
| `enviar_credenciais_gui_v3.py` | Ajustes no fluxo de senha/CPF |
| `enviar_credenciais_gui_v4.py` | Versão atual |

## ⚠️ Aviso

Este repositório contém **apenas o código-fonte**. Nenhuma credencial, e-mail de colaborador ou arquivo de saída (CSV, PDFs, executável) está versionado — todos esses itens estão listados no `.gitignore`.

## 📄 Licença

Uso interno.
