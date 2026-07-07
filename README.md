# Dashboard Facilities — Hype Coworking

Site estático com os dashboards Comercial e Operacional + geração automática
de relatórios PDF via GitHub Actions.

## Estrutura

```
index.html                        → redireciona para comercial.html
comercial.html                    → dashboard Comercial (dados embutidos)
operacional.html                  → dashboard Operacional (dados embutidos)
admin.html                        → ferramenta: sobe planilha, revisa textos, gera PDFs
relatorios-core.js                → lógica pura (parse, insights, HTML de PDF)
                                     usada por admin.html E pelo script Node
dados/dados.xlsx                  → planilha do mês (commitar aqui dispara automação)
scripts/gerar-relatorios.js       → script Node que gera PDFs + atualiza dashboards
.github/workflows/gerar-relatorios.yml  → GitHub Action que roda automaticamente
netlify.toml                      → configuração de deploy e rotas
template_dados.xlsx               → modelo em branco para preencher
```

## Fluxo mensal automatizado (depois de conectado — ver abaixo)

1. Preencha a planilha com os dados do mês.
2. Salve como `dados/dados.xlsx`.
3. `git add dados/dados.xlsx && git commit -m "dados: Jun-26" && git push`
4. O GitHub Action dispara sozinho:
   - Lê a planilha
   - Gera os insights automaticamente
   - Renderiza os dois PDFs com Chromium headless
   - Injeta os dados nos dashboards HTML
   - Commita tudo de volta (`relatorios/*.pdf`, `comercial.html`, `operacional.html`)
5. O Netlify detecta o commit e republica o site.

**Resultado:** você só faz push da planilha → PDFs e dashboards prontos em ~2 min.

## Fluxo manual (alternativa, sem automação)

1. Abra `seusite.netlify.app/admin`
2. Arraste a planilha → revise os textos
3. Clique **Gerar PDFs** → salve cada relatório como PDF na caixa de impressão
4. Clique **Gerar dashboard.zip** → suba no Netlify ou commite os arquivos atualizados

## Primeira configuração

### 1. Criar repositório no GitHub
Crie um repo vazio (sem README) e anote a URL.

### 2. Subir este código
```bash
git remote add origin SUA_URL_AQUI
git branch -M main
git push -u origin main
```

### 3. Conectar ao Netlify
No painel do Netlify: **Add new site → Import an existing project → GitHub**
→ escolha este repositório → deixe Build command vazio e Publish directory como `.`

### 4. Rodar localmente (opcional)
```bash
npm install
npm run gerar:local   # detecta Chromium local automaticamente
```
Se não funcionar, instale o Chromium do Playwright: `npx playwright install chromium`
e então: `npm run gerar`

## Por que `relatorios-core.js` existe

O admin.html e o script de automação usam **exatamente a mesma função** para
ler a planilha e gerar os textos de insight (`relatorios-core.js`). Sem isso,
a lógica ficaria duplicada e qualquer ajuste numa planilha nova precisaria ser
feito em dois lugares — foi exatamente isso que causou o bug do Desfalque (Edgar
esquecido num lado).
