# Controle de Produtividade — Dobra

Sistema web para monitorar o **tempo de produção** de um setor. O operador
escolhe a máquina, **inicia** o cronômetro, aponta a **OP**, **pausa** nos
intervalos e **finaliza** com a quantidade produzida. A gerência acompanha
tudo por um **dashboard**, e o **admin** cadastra máquinas, usuários e setores.

Foi desenhado para **Dobra**, mas serve igual para **Corte, Estamparia,
Solda, Acabamento** — basta cadastrar as máquinas de cada setor.

---

## O que já vem pronto

**Painel do operador** (`/painel`)
- Grade de máquinas do setor, com status ao vivo (Livre / Produzindo / Pausado).
- Iniciar → **bipe a OP com o leitor de código de barras ou digite**; informa
  operador, código, qtd prevista. (O leitor envia Enter ao final: se o operador
  já estiver preenchido, a produção inicia sozinha ao bipar.)
- Cronômetro por máquina que corre em tempo real.
- **Pausar / Retomar** com motivo (café, almoço, manutenção…). O tempo pausado
  é **descontado** do tempo produtivo.
- Finalizar → quantidade produzida, refugo e observação.

**Painel gerencial** (`/dashboard`)
- KPIs: OPs, peças, horas produtivas, horas em pausa, eficiência, peças/hora,
  refugo e quantos apontamentos estão em andamento agora.
- **OEE** em destaque (medidor + os 3 pilares: Disponibilidade, Desempenho e
  Qualidade), com a meta de OEE como referência.
- Gráficos de produção por dia e de motivos de pausa.
- Rankings por operador e por máquina (com coluna de OEE).
- Tabela dos últimos apontamentos + **exportação para Excel** (com colunas de
  meta, disponibilidade, desempenho, qualidade e OEE).
- Filtros por período (hoje, ontem, 7/30 dias, mês) e por setor.

**Administração** (edita e ajusta tudo)
- `/admin/maquinas` — cadastro de máquinas (nome, código, setor, ativa, ordem
  e **meta de peças/hora** — a capacidade nominal usada no Desempenho do OEE).
- `/admin/usuarios` — usuários e acessos: **Operador** (aponta), **Gerencial**
  (vê dashboard e OEE) e **Administração** (configura tudo). É aqui que você
  define quem enxerga o OEE e o painel gerencial.
- `/admin/config` — setores, setor em foco, motivos de pausa, **meta de
  peças/hora padrão**, **meta de OEE (%)** e reset.
- No dashboard, o admin pode **editar ou excluir** cada apontamento (corrigir
  OP, quantidade, refugo etc.). Toda alteração fica registrada em auditoria.

---

## Acessos padrão (troque depois!)

| Usuário    | Senha     | Acessos                     |
|------------|-----------|-----------------------------|
| `admin`    | `admin123`| Admin + Gerencial + Operador|
| `operador` | `123456`  | Operador                    |

> Ao subir para produção, entre como `admin`, mude a senha e crie os usuários reais.

---

## Rodar na sua máquina

```bash
pip install -r requirements.txt
python app.py
```

Abre em `http://localhost:5000`. Sem `DATABASE_URL`, ele usa um arquivo
`dobra.db` (SQLite) — ótimo para testar.

---

## Publicar (Railway ou similar)

1. Suba este projeto num repositório Git.
2. No Railway: **New Project → Deploy from GitHub** e selecione o repositório.
3. Adicione um banco **PostgreSQL** (o Railway cria a variável `DATABASE_URL`).
4. (Recomendado) defina uma variável `SECRET_KEY` com um texto aleatório.
5. Deploy. O `Procfile`/`railway.json` já sobem o app com gunicorn.

> **Por que Postgres em produção?** O disco do Railway é efêmero — a cada
> redeploy o SQLite seria zerado. Com `DATABASE_URL` de Postgres, os dados
> (máquinas, usuários e apontamentos) ficam salvos.

Variáveis de ambiente:
- `DATABASE_URL` — string do Postgres (em produção). Sem ela → SQLite local.
- `SECRET_KEY` — chave das sessões. Defina em produção.

---

## Usar em outro setor (Corte, Solda, etc.)

1. Em **Configurações**, confira/adicione o setor e escolha o **setor em foco**.
2. Em **Máquinas**, cadastre as máquinas daquele setor.
3. Pronto — o painel do operador e o dashboard passam a operar naquele setor.

Cada apontamento guarda o setor, então o dashboard filtra por setor sem misturar.

---

## Como o tempo é calculado

```
tempo produtivo = (fim − início) − soma de todas as pausas
```

O cronômetro do operador congela enquanto a máquina está pausada e volta a
correr ao retomar. Datas são gravadas em UTC e exibidas no horário local
(Brasil, UTC−3; ajuste `FUSO_LOCAL_HORAS` no `app.py` se precisar).

## Como o OEE é calculado

**OEE = Disponibilidade × Desempenho × Qualidade**

```
Disponibilidade = tempo produtivo ÷ (produtivo + pausas)
Desempenho      = peças produzidas ÷ (meta pç/h × horas produtivas)   [limitado a 100%]
Qualidade       = (produzidas − refugo) ÷ produzidas
```

A **meta de peças/hora** vem da máquina (tela de Máquinas); se a máquina não
tiver meta, usa-se a **meta padrão** de Configurações. A meta em vigor é
congelada em cada apontamento no momento do início, para que relatórios
antigos não mudem quando você reajustar a meta. Sem nenhuma meta cadastrada,
Disponibilidade e Qualidade continuam sendo calculadas, mas Desempenho e OEE
aparecem como "—" até você definir a meta.

## Estrutura

```
app.py                 backend (Flask + SQLAlchemy)
requirements.txt       dependências
Procfile / railway.json  deploy
static/style.css       design system
templates/             telas (login, painel, dashboard, admin…)
```
