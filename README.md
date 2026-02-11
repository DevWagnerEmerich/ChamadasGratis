# Gerador de Lista de Presença

Webapp simples para gerar listas de presença em Excel (.xlsx) formatadas para impressão em A4 paisagem.

## Funcionalidades
- **Personalização**: Nome da turma e lista de alunos.
- **Ajuste Automático**: O layout se adapta para caber em uma página A4.
- **Excel Profissional**: Formatação limpa, cabeçalhos, datas rotacionadas e células preparadas para preenchimento manual.
- **Design Moderno**: Interface intuitiva e agradável.

## Como Usar

1. Certifique-se de ter o [Node.js](https://nodejs.org/) instalado.
2. Abra o terminal na pasta do projeto.
3. Instale as dependências (se ainda não fez):
   ```bash
   npm install
   ```
4. Inicie o servidor de desenvolvimento:
   ```bash
   npm run dev
   ```
5. Acesse o endereço exibido (geralmente `http://localhost:5173`) no seu navegador.

## Requisitos
- Conexão com a internet (para carregar a biblioteca ExcelJS via CDN e fontes).

## Tecnologias
- React + Vite
- ExcelJS (Geração de planilhas)
- CSS Moderno
