# RaspiCash

Applicazione web per tenere le spese di casa, pensata per girare su un Raspberry Pi e usata dal telefono.
Flask + SQLite, nessun servizio esterno: i dati restano sul Raspberry.

Funzioni principali:

- inserimento spese con categoria, buoni pasto e spese private (visibili solo a chi le inserisce);
- storico con filtri e riepilogo annuale per mese e categoria;
- auto: rifornimenti, consumi (km/L, L/100 km), costo ogni 100 km e altre spese dell'auto;
- viaggi, anche di lavoro con spese pagate dall'azienda;
- prestiti e anticipi con saldo per persona;
- esportazione e importazione Excel;
- lettura dello scontrino da foto (opzionale, Tesseract);
- più utenti, backup e ripristino.

Versione 1.0.0.

## Requisiti

- Raspberry Pi con Raspberry Pi OS e Python 3.11 o successivo
- Facoltativo, per la lettura degli scontrini: `tesseract-ocr` e `tesseract-ocr-ita`

## Installazione

```
git clone git@github.com:simo94pasqui/raspicash.git
cd raspicash
python3 -m venv venv
venv/bin/pip install -r requirements.txt
```

Il servizio systemd è in `raspicash.service`: se il tuo utente non è `pi`, correggi `User=` e i percorsi, poi:

```
sudo cp raspicash.service /etc/systemd/system/
sudo systemctl enable --now raspicash
```

Apri `http://raspberry.local:8000` (o l'indirizzo IP del Pi) dal telefono. Al primo avvio crei l'amministratore;
gli altri utenti si aggiungono da Admin, Gestione utenti. Dal menu del browser puoi aggiungere l'app alla schermata Home.

Controlla anche il fuso orario del Pi (`sudo raspi-config`, Localisation): le date proposte dipendono da quello.

L'app usa HTTP senza cifratura ed è pensata per la rete di casa: non esporla direttamente su Internet.

## Dati

Tutto sta nella cartella `data/` (non versionata):

- `raspicash.db`: il database;
- `secret.key`: chiave delle sessioni, creata al primo avvio;
- `prima_import_*.db`: copia di sicurezza prima di ogni importazione (si tengono le ultime 5);
- `prima_del_ripristino_*.db`: copia dello stato precedente a un ripristino.

## Backup e ripristino

Da Admin, Backup completo si scarica un unico file `.db` con tutto (utenti, categorie, spese, anche le private di tutti).
Conservalo fuori dal Raspberry. In alternativa basta copiare `data/raspicash.db`.

Se la microSD si rompe: reinstalla il sistema e RaspiCash come sopra, apri l'app e nella pagina del primo avvio usa
"Oppure ripristina da backup" caricando il file `.db`.

## Esporta e importa Excel

Il menu Esporta/Importa ha due schede.

**Esporta**: spese dell'anno (condivise o solo le tue private), storico di un'auto, rifornimenti, viaggio, prestiti.

**Importa** (per le proprie spese): si carica un `.xlsx`, si vede un'anteprima (nuove, già presenti, da completare,
da controllare, errori con numero di riga) e solo dopo la conferma si scrive. Gli errori si possono correggere uno alla volta.

- Spese: stesso formato del foglio Spese esportato. Obbligatorie Data, Descrizione, Categoria, Importo.
  Le colonne si riconoscono dal nome, l'ordine non conta e quelle in più si ignorano.
- Spese di un viaggio: elenco semplice con Data, Descrizione, Importo, Buono, Pagata da (Io o Azienda), Privata.
  Il viaggio si sceglie o si crea nel modulo.
- Rifornimenti: stesso formato del foglio Rifornimenti. Obbligatorie Data e Auto, più Importo o Litri
  (con due valori su tre il terzo viene calcolato). Una spesa Benzina uguale ma senza dettagli viene completata
  invece di essere duplicata.
- Le righe già presenti (stessa data, descrizione, categoria, importo e buono) vengono saltate: reimportare lo stesso
  file non cambia nulla. Il proprietario è sempre chi importa.
- Categorie, auto, persone e viaggi che non esistono si creano con la spunta "Crea i nomi mancanti".

In `examples/` ci sono tre file di esempio.

## Lettura scontrino

Nella pagina Inserisci, "Scansiona scontrino" legge la foto e compila importo, negozio, data e categoria;
non viene salvato nulla finché non premi Salva. La lettura avviene sul Raspberry, la foto non esce di casa.

```
sudo apt install tesseract-ocr tesseract-ocr-ita
```

Su un Pi la lettura può richiedere qualche secondo: per questo `raspicash.service` usa `--timeout 90`.
Senza Tesseract il pulsante mostra un avviso e il resto dell'app funziona normalmente.

## Aggiornamento e log

```
git pull
sudo systemctl restart raspicash
journalctl -u raspicash -f
```

## Sviluppo

```
venv/bin/python app.py
```

Avvia il server di sviluppo su `http://localhost:8000`.

## Licenza

Rilasciato con licenza MIT, vedi il file `LICENSE`.
