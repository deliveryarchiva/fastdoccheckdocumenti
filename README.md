# Archiva File DB

Web app per la ricerca cross-referenziata di documenti tra i fogli **Estrazione Archiva** e **Estrazione Postel**.

## Funzionalità

- Ricerca per Ragione Sociale, P.IVA, Nome File (parziale o esatto), Range di date
- Data estratta automaticamente dal nome file (`PIVA_AAAA_M_G_...`)
- Colonne risultato: Presente in Archiva (SI/NO), Presente in Postel (SI/NO)
- Esportazione risultati in CSV (con BOM per Excel italiano)
- Autenticazione con ruoli admin/user
- Upload del file Excel via pannello Admin

## Struttura file Excel attesa

Il file deve avere i fogli:
- `Estrazione Archiva` — col 8 (file_name), col 95 (PIVA), col 98 (RAGIONE_SOCIALE), col 116 (r_object_id)
- `Estrazione Postel` — col 1 (r_object_id), col 3 (object_name), col 33 (pt_ragione_sociale), col 89 (pt_piva)

## Deploy su Railway

1. Push su GitHub
2. Nuovo progetto Railway → Connect GitHub repo
3. Aggiungi **Volume** con mount path `/data`
4. Variabili d'ambiente:
   - `RAILWAY_VOLUME_MOUNT_PATH=/data`
   - `DEFAULT_PASSWORD=archiva2026`  ← cambiare subito dopo il primo login

5. Primo login con `marco.pastore` / `archiva2026`
6. Admin → Carica il file `database.xlsx`
7. Attendere il parsing (30–60 secondi per ~110k righe)

## Sviluppo locale

```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

Aprire http://localhost:8000/login
