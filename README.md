# OLTA Newsletter Extractor — versione 0.1

Prima versione Streamlit per estrarre automaticamente informazioni dalle newsletter OLTA con sole regole Python, senza modelli AI.

## Formati supportati

- `.msg` Outlook
- `.eml`
- `.html` / `.htm`
- `.txt`

## Output Excel

Il file scaricato dall'app contiene tre fogli:

1. `Database`: righe estratte con le colonne principali e tecniche.
2. `Summary`: riepilogo quantitativo.
3. `Text summary`: breve testo riepilogativo modificabile nell'app.

## Colonne principali

- OLTA MITTENTE
- DATA NEWSLETTER
- TIPOLOGIA
- COMPAGNIA PROMOSSA
- DESTINAZIONE/MERCATO/TRATTA
- PROMOZIONE

## Colonne tecniche

- OGGETTO EMAIL
- FILE ORIGINE
- TESTO ESTRATTO
- CONFIDENCE
- NOTE

## Installazione

Aprire il terminale nella cartella del progetto e lanciare:

```bash
pip install -r requirements.txt
```

## Avvio app

```bash
streamlit run app.py
```

## Nota importante

La versione 0.1 è rule-based. Funziona bene sui formati ricorrenti ma richiede revisione manuale, soprattutto quando:

- la promozione è contenuta solo in immagini;
- la compagnia è presente solo nel logo;
- la newsletter contiene promozioni multiple in layout molto variabile;
- la meccanica commerciale è descritta in modo implicito.

La colonna `CONFIDENCE` serve a identificare le righe da controllare con maggiore attenzione.
