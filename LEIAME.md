# Cadastro de Visitantes · Batizado Brincapoeira

Formulário só para **capoeiristas de outros grupos** que vêm ao Batizado (28/11, 9h–12h, IFSul Sapucaia do Sul).
Campos: **Nome, Telefone, CPF (validado) e Grupo**. Cada envio vira uma linha numa **planilha Google separada**
(dá para baixar em Excel: *Arquivo → Fazer download → Microsoft Excel*).

## 1. Criar a planilha e o script

1. Crie uma planilha nova no Google Planilhas, por exemplo **"Visitantes Batizado Brincapoeira"**.
2. Menu **Extensões → Apps Script**.
3. Apague o que estiver lá e cole todo o conteúdo de `apps-script/Code.gs`. Salve (ícone de disquete).
4. (Opcional) Escolha a função `testeManual` e clique em **▶ Executar**. Autorize com sua conta Google.
   Deve aparecer a aba **Visitantes** com uma linha *Visitante TESTE*. Depois apague essa linha.

## 2. Publicar o script

1. **Implantar → Nova implantação** → engrenagem → **App da Web**.
2. *Executar como*: **Eu**. *Quem pode acessar*: **Qualquer pessoa**.
3. **Implantar** e copie a URL que termina em `/exec`.

## 3. Ligar o site à planilha

Abra `index.html`, procure `SCRIPT_URL: ""` (perto do fim do arquivo) e cole a URL entre as aspas.

> Se mudar o `Code.gs` depois: **Implantar → Gerenciar implantações → lápis → Versão: Nova versão**.
> Assim a URL continua a mesma.

## 4. Colocar no ar

Mesmo esquema do cadastro principal: repositório no GitHub + GitHub Pages (ou Vercel).
Sugestão de endereço: `visitantes.brincapoeira.com` (registro CNAME `visitantes` → `thallesyy.github.io` na Hostinger).

## Planilha

| Data/hora | Protocolo | Nome | Telefone | CPF | Grupo | Presença |
|---|---|---|---|---|---|---|

- O protocolo começa com **VS-** (o do cadastro dos alunos começa com BC-).
- **Presença** é uma caixinha para marcar na portaria.
- Menu **Brincapoeira → Arrumar planilha** recoloca o cabeçalho e apaga linhas vazias.
- A planilha tem CPF e telefone: **não compartilhe o link** com quem não é da organização.
