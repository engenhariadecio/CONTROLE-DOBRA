# DECIOMES

Sistema web de **controle integrado de produção** (MES) da Décio, para todas as
máquinas e setores da fábrica — Corte (puncionadeiras, laser), Estamparia
(furadeira, prensas), Dobra (dobradeiras), Solda (solda ponto, solda) e
Acabamento (polimento, acabamento). O operador escolhe a máquina, **inicia** o
cronômetro, **pausa** nos intervalos (Almoço/Janta, Café, Laboral, Outros…) e,
ao concluir, **bipa a OP** — que puxa material, descrição e quantidade da
**lista mestra do SAP**. A gerência acompanha em um **painel de visualização em
tempo real** (histórico + produção ao vivo + OEE), e o **admin** configura e
ajusta tudo.

Cada **usuário pertence a um setor** (sua área), e as máquinas são cadastradas
por setor — então o mesmo sistema atende toda a fábrica de forma integrada.

---

## O que já vem pronto

**Painel do operador** (`/painel`)
- Grade de máquinas do setor, com status ao vivo (Livre / Produzindo / Pausado)
  nas cores da Décio (azul e verde).
- **Iniciar** → informa só o operador; o cronômetro começa na hora.
- **Pausar / Retomar** com os motivos do Banho (Almoço/Janta, Café, Laboral,
  Outros…) em botões grandes; o tempo pausado é **descontado** do produtivo.
- **Finalizar** → **bipe a OP** (leitor de código de barras ou digite). A OP é
  buscada na **lista mestra do SAP** e preenche material, descrição e
  quantidade automaticamente; depois informe a quantidade produzida e o refugo.

### Lista mestra do SAP (mesma do Banho)
Coloque o arquivo **`lista_mestra.xlsx`** (ou `.csv`/`.txt`) na raiz do
projeto — mesmo formato do sistema Banho (colunas Ordem, Material, Texto breve,
Quantidade). O sistema a carrega em memória no startup; a bipagem da OP no final
usa essa lista. O leitor de código de barras que traz dígitos extras é tratado
automaticamente (remove 4 no início e 4 no fim). Em `Configurações` há o status
da lista e um botão para recarregar. O arquivo de exemplo já acompanha o projeto.

**Painel gerencial** (`/dashboard`) — **somente visualização**
- Para os usuários com acesso **Gerencial**. Não edita nada: é um monitor.
- **Atualiza sozinho a cada 20s** (produção em tempo real, estilo MES).
- KPIs: OPs, peças, horas produtivas, horas em pausa, eficiência, peças/hora,
  refugo e quantos apontamentos estão em andamento agora.
- **OEE** em destaque (medidor + os 3 pilares: Disponibilidade, Desempenho e
  Qualidade), com a meta de OEE como referência.
- **Linha do tempo por máquina** (estilo Banho): barras de produção e pausas ao
  longo do dia, com filtro por **turno** e por **máquina**.
- Gráficos de produção por dia e de motivos de pausa.
- Rankings por operador e por máquina (com OEE) e tabela ao vivo do que está
  em andamento agora.
- Filtros por período, **por setor** e **por turno**; exportação para Excel.

**Administração** (configura e ajusta tudo)
- `/admin/maquinas` — máquinas por setor (nome, código, setor, ativa, ordem e
  **meta de peças/hora** — capacidade nominal usada no Desempenho do OEE).
- `/admin/usuarios` — usuários por **área/setor** e acessos: **Operador**
  (aponta na sua área), **Gerencial** (visualiza o dashboard/OEE da sua área) e
  **Administração** (tudo). É aqui que você define quem enxerga o painel
  gerencial e o OEE, e a que setor cada pessoa pertence.
- `/admin/config` — setores, setor em foco, **motivos de pausa** (café,
  laboral, almoço, janta…), **turnos** (3 turnos com horários configuráveis,
  aceita virada de meia-noite), **meta de peças/hora padrão**, **meta de OEE
  (%)** e reset.
- `/admin/apontamentos` — consulta com filtros (período, setor, turno) e
  **edição/exclusão** de cada registro de produção. Toda alteração fica
  registrada na auditoria. (A edição fica aqui, no admin — o painel gerencial é
  só visualização.)

**Vários administradores.** Basta dar o acesso **Administração** a mais de um
usuário. Todos podem editar a base (máquinas, usuários, configurações, turnos)
e os dados (apontamentos).

**Backup e restauração** (em `/admin/config`)
- **Baixar backup** gera um único arquivo `.json` com a base completa: usuários
  (com senha), máquinas, configurações, turnos, apontamentos e auditoria.
- **Restaurar backup** substitui toda a base atual pelo conteúdo do arquivo
  (pede confirmação). É também a forma de migrar a base entre ambientes.
- Recomendação: baixe um backup com regularidade e sempre antes de restaurar.

**Setores e máquinas de exemplo.** Numa base nova, o sistema já vem com um
parque de exemplo cobrindo os cinco setores — Corte (puncionadeiras, laser),
Estamparia (furadeira, prensas), Dobra (dobradeiras), Solda (solda ponto,
solda) e Acabamento (polimento, acabamento). Ajuste/adicione tudo em Máquinas.

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
