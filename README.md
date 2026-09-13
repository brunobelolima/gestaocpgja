# gestaocpgja — IAC-Pal + Gestão & BI

Projeto estático compatível com **GitHub Pages** para o site:
`https://brunobelolima.github.io/gestaocpgja/`

## Páginas

- `index.html` — IAC-Pal v1.2, fluxo interativo de complexidade paliativa.
- `gestao.html` — planilha interativa e painel de BI.

## Gestão & BI

O painel funciona inteiramente no navegador e aceita arquivos `.xlsx`/`.xls`.
Ele reconhece as abas:

- Pacientes
- Leads
- Atendimentos
- Programas
- Financeiro
- Custos
- Capacidade
- Qualidade

A planilha modelo está em `assets/Painel_Gestao_Clinica_Cuidados_Paliativos_v1.1_BI.xlsx`.

### Indicadores

- pacientes ativos;
- leads;
- taxa de conversão;
- MRR estimado;
- receita recebida;
- custos;
- margem;
- inadimplência;
- pacientes por programa;
- funil e origem de leads;
- receita x custos por competência;
- capacidade;
- satisfação e eventos não planejados.

### Planilha interativa

Na aba **Planilha**, é possível editar células, adicionar/excluir linhas e exportar um novo Excel.
Os dados ficam somente na memória do navegador durante a sessão. Não há backend próprio.

## Publicar no GitHub Pages

1. Faça backup do repositório atual.
2. Copie todos os arquivos deste pacote para a raiz do repositório `gestaocpgja`.
3. Commit e push:

```bash
git add .
git commit -m "Integra painel de gestão e BI"
git push origin main
```

4. Em **Settings → Pages**, mantenha a publicação pelo branch `main`/raiz, se já estiver configurada assim.
5. Acesse:
   - IAC-Pal: `https://brunobelolima.github.io/gestaocpgja/`
   - Gestão & BI: `https://brunobelolima.github.io/gestaocpgja/gestao.html`

## Dependências web

O painel usa CDNs públicos no navegador:
- SheetJS `xlsx@0.18.5`
- Chart.js `4.4.7`

Por isso, é necessário acesso à internet para carregar essas bibliotecas.

## Segurança e LGPD

**Não coloque planilhas contendo dados identificáveis de pacientes dentro do repositório GitHub.**
A planilha incluída neste pacote é apenas o modelo operacional. O arquivo real deve ser carregado pelo usuário diretamente no navegador.

O GitHub Pages é estático e não oferece autenticação, prontuário, banco de dados, controle de acesso clínico ou backup multiusuário. Para uso compartilhado entre profissionais, será necessário migrar para uma arquitetura com autenticação e backend adequado.
