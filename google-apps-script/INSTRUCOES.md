# Como publicar o backend (Google Sheets + Apps Script)

## 1. Criar a planilha

Crie uma Google Sheet nova (na conta da Regiane, `divarebel.on1@gmail.com`, como você faz nos seus outros sites) com estas 5 abas. A **primeira linha de cada aba é o cabeçalho**, exatamente com esses nomes (maiúsculas/minúsculas importam):

### Aba `Pacientes`
| Email | Nome | DataInicio | RetornosRealizados | ReceitasSalvas | ProgressoPercent | ProximoRetornoData | ProximoRetornoHora | PlanoTexto | PontosTotal |
|---|---|---|---|---|---|---|---|---|---|
| paciente@exemplo.com | Marta Silva | 01/06/2026 | 3 | 12 | 78 | 28/07/2026 | 15h00 | Siga as orientações da última consulta... | 340 |

- `DataInicio` precisa ser uma data de verdade (não texto) — a coluna calcula os "dias de acompanhamento" sozinha.
- `PontosTotal` começa em 0 pra paciente nova; o sistema soma sozinho quando ela comenta.

### Aba `Materiais`
| Id | Email | Tipo | Titulo | Descricao | Link |
|---|---|---|---|---|---|
| 1 | paciente@exemplo.com | PDF | Guia de intestino | Feito especialmente pro seu caso | https://... |
| 2 | TODOS | PDF | Planner semanal | Vale pra todo mundo do programa | https://... |

- Coloque `TODOS` na coluna `Email` pra um material aparecer pra todas as pacientes.

### Aba `Comentarios`
| Id | PostId | Email | Nome | Texto | DataHora |
|---|---|---|---|---|---|

Deixe só o cabeçalho — o sistema preenche sozinho conforme as pacientes comentam.

### Aba `PontosLog`
| Id | Email | Tipo | Pontos | Data |
|---|---|---|---|---|

Também só o cabeçalho — é o histórico de cada ponto ganho.

### Aba `Horarios`
| Data | Hora | Disponivel |
|---|---|---|
| 28/07/2026 | 15h00 | Sim |
| 29/07/2026 | 10h00 | Sim |

Você (ou a Regiane) preenche os horários livres aqui. Quando uma paciente solicita, o sistema marca `Disponivel` como `Não` automaticamente.

### Aba `Agendamentos`
| Email | Data | Hora | Status | DataSolicitacao |
|---|---|---|---|---|

Só o cabeçalho — preenchido sozinho quando alguém solicita um horário. A Regiane confirma manualmente (não há checagem de conflito automática nesta versão).

## 2. Copiar o ID da planilha

Na URL da planilha, o ID é o trecho entre `/d/` e `/edit`:
`https://docs.google.com/spreadsheets/d/`**`ESTE_TRECHO_AQUI`**`/edit`

## 3. Publicar o Apps Script

1. Na própria planilha, vá em **Extensões → Apps Script**.
2. Apague o conteúdo padrão e cole todo o conteúdo de `Code.gs` (está na mesma pasta deste arquivo).
3. No topo do script, troque `COLE_AQUI_O_ID_DA_PLANILHA` pelo ID que você copiou no passo 2.
4. Clique em **Implantar → Nova implantação**.
5. Tipo: **App da Web**.
6. "Executar como": **Eu** (sua conta, `babadosdaaline@gmail.com`, do jeito que você já faz nos outros sites).
7. "Quem pode acessar": **Qualquer pessoa**.
8. Clique em **Implantar** e autorize as permissões pedidas.
9. Copie a **URL do app da Web** gerada (termina em `/exec`).

## 4. Colar a URL no site

Abra o arquivo `assets/js/api.js` no projeto do site e troque:

```js
const API_URL = 'COLE_AQUI_A_URL_DO_WEB_APP';
```

pela URL que você copiou no passo 9. Salve, suba pro GitHub, e me avise — eu termino de testar a integração com você.

## Se algo não funcionar

- Toda vez que você editar o `Code.gs`, precisa fazer **Implantar → Gerenciar implantações → editar (ícone de lápis) → Nova versão → Implantar** pra as mudanças valerem (só salvar o script não é suficiente).
- Se o navegador mostrar erro de permissão/CORS, confirme que "Quem pode acessar" está mesmo como **Qualquer pessoa**.
