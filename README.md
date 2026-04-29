# Galp Frota — Gerador de Cartas

## Deploy (gratuito) — Vercel + Supabase

### Passo 1 — Supabase (base de dados)

1. Vai a **supabase.com** → cria conta gratuita
2. Clica **New project**, dá um nome (ex: `galp-cartas`), escolhe uma password e região
3. Aguarda o projeto criar (~1 min)
4. Vai a **SQL Editor** e corre este comando:

```sql
CREATE TABLE enderecos (
  empresa TEXT PRIMARY KEY,
  avenida TEXT DEFAULT '',
  numero  TEXT DEFAULT '',
  postal  TEXT DEFAULT ''
);
```

5. Vai a **Settings → API** e copia:
   - **Project URL** → vai ser o `SUPABASE_URL`
   - **anon public key** → vai ser o `SUPABASE_KEY`

---

### Passo 2 — GitHub

1. Vai a **github.com** → New repository → nome `galp-cartas` → Create
2. Faz upload de todos estes ficheiros:
   - `api/index.py`
   - `public/index.html`
   - `requirements.txt`
   - `vercel.json`

---

### Passo 3 — Vercel

1. Vai a **vercel.com** → Sign up with GitHub
2. **Add New Project** → importa o repositório `galp-cartas`
3. Em **Environment Variables** adiciona:
   - `SUPABASE_URL` → (o Project URL do Supabase)
   - `SUPABASE_KEY` → (a anon public key do Supabase)
4. Clica **Deploy**

Fica com um link tipo `https://galp-cartas.vercel.app` — 100% gratuito, sem limite de tempo.

---

## Como usar

1. Abre o link da app
2. Faz upload do ficheiro principal (descarregas do SharePoint)
3. Faz upload dos 4 templates .docx
4. Indica o número do lote → **Analisar**
5. Preenche endereços em falta (ficam guardados para sempre)
6. **Gerar** → descarregas o ZIP
