# Cash Register 

A web application for tracking cash register income and expenses. Written in
Python (Flask), runs as a website in Docker.

## Features

- **10 sub-registers (objects)** — switch between them using the tabs at the
  top of the page. Each register keeps its own separate operations log and
  its own balance. The editor (`user2`) can rename each register.
- **3 currencies**: hryvnia (UAH), US dollar (USD), euro (EUR). The balance
  is shown separately for each currency, plus an overall total in UAH
  equivalent (calculated using each individual operation's exchange rate).
- For every operation, the **exchange rate and purpose** (description) are
  **required** — an operation cannot be added without these fields.
- **Counterparties directory** (like in accounting software): income
  operations specify "From whom", expense operations specify "To whom".
  Counterparties can be added, edited and deleted (deletion is blocked if
  a counterparty is already used in operations) on a separate
  "Counterparties" page, and a new one can also be added quickly right from
  the operation-entry form with the "+ New counterparty" button.
- **Log filters** (currency, operation type, counterparty, date range) —
  a collapsed panel above the log that expands on click.
- **Transfers between registers** — move funds from one register to another:
  it automatically creates an "Expense" in the source register (marked "To
  whom" = the destination register) and an "Income" in the destination
  register (marked "From whom" = the source register). The two operations
  are linked and are deleted together.
- **Excel export** — the "Download Excel" button above the log generates an
  .xlsx file containing exactly the operations currently shown on screen
  (respecting the active filters), plus a summary table with income,
  expenses and the register's overall balance per currency.

 **Important:** if you are updating an already-running project from an
older version — delete the `data/kasa.db` file before restarting, since the
database schema has changed (registers, currency, exchange rate were added).

## Running it

You need Docker and Docker Compose installed.

```bash
docker compose up -d --build
```

Once started, the site will be available at:

```
http://localhost:5000
```

Data (the SQLite database) is stored in the `./data` folder, which is
mounted as a volume — so all operations are preserved after the container
restarts.

## Stopping it

```bash
docker compose down
```

## Logins and roles

| Login   | Password  | Role                                        |
|---------|-----------|----------------------------------------------|
| User1   | password  | View operations and balance only              |
| user2   | password  | View + add / delete operations                |

> Logins, passwords and roles are stored in `app.py`, in the `USERS`
> dictionary. Change the passwords there if needed, or rewrite it to store
> them in a database with hashing (e.g. via `werkzeug.security`).

## Changing the session key

`docker-compose.yml` has a `SECRET_KEY` environment variable — make sure to
change it to your own random string before using this in production.

## Project structure

```
kasa/
├── app.py                 # Flask application
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── data/                  # kasa.db is stored here (volume)
├── static/
│   └── style.css          # interface styles
└── templates/
    ├── login.html         # login page
    ├── index.html         # register operations log
    └── contragents.html   # counterparties directory page
```

## Running without Docker (locally)

```bash
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

The site will be available at `http://localhost:5000`.
