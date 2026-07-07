#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pseudo-anonimizzatore avanzato per atti legali italiani. (v5)

Novità della v5:
- Modalità interattiva guidata: avviando senza argomenti (o con -i) parte un
  menu a domande; ora spiega file vs cartella con esempi e accetta il
  trascinamento del percorso.
- Avvio con un clic: pausa finale in modalità interattiva così, aperto con un
  doppio clic, il programma non chiude la finestra prima di mostrare i risultati.
- OCR per atti scansionati e immagini: i PDF senza testo vengono letti pagina
  per pagina via OCR; accettati anche file immagine (.png .jpg .tiff...);
  nuova opzione --ocr per forzare l'OCR; usa la lingua italiana se disponibile.

Correzioni rispetto alla v3 (introdotte nella v4):
- TOPONIMO: la prima parola non può più essere una preposizione nuda, quindi
  "in corso di causa", "in via di definizione" ecc. non sono più falsi
  positivi INDIRIZZO ("via di Ripetta" e simili restano coperti).
- LUOGO_NASCITA: lookahead esteso (a-capo, ":", congiunzioni "e"/"ed"),
  prima non matchava a fine riga né in "nata a Milano e residente...".
- Audit: usa TUTTI i pattern di una categoria (prima il secondo pattern
  TELEFONO, cellulari senza etichetta, era escluso); nuova categoria
  LUOGO_NASCITA_RESIDUO; INDIRIZZO_RESIDUO senza falsi positivi
  ("via preliminare"); max_esempi rispettato davvero.
- leggi_txt: cp1252 provato PRIMA di latin-1 (latin-1 non fallisce mai,
  quindi il ramo cp1252 era irraggiungibile e i file Windows con ’ e –
  venivano decodificati male).
- elabora: un file che fallisce (es. PDF scansionato) non abortisce più
  l'intero fascicolo prima del salvataggio della mappa; la password di
  --encrypt-map è chiesta e validata PRIMA dell'elaborazione.
- _residui_da_verificare.txt contiene dati reali: ora è scritto 0600 come
  mappa e report, e l'avviso finale lo menziona.
- Sostituzione su gruppo basata sugli span del match (str.replace poteva
  colpire un'occorrenza identica precedente dentro il match).
- Range Unicode corretti (esclusi × U+00D7 e ÷ U+00F7); minimo di
  iterazioni KDF quando si decifra una mappa.
"""

from __future__ import annotations
import argparse, base64, getpass, json, os, re, stat, sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

try:
    import pypdf  # type: ignore
except ImportError:
    pypdf = None

_NLP = None
try:
    import spacy  # type: ignore
    for _MODELLO in ("it_core_news_lg", "it_core_news_md", "it_core_news_sm"):
        try:
            _NLP = spacy.load(_MODELLO); break
        except Exception:
            continue
except ImportError:
    pass

# OCR opzionale: pytesseract + Pillow per riconoscere il testo dalle immagini,
# PyMuPDF (fitz) per rasterizzare le pagine dei PDF scansionati.
try:
    import pytesseract  # type: ignore
    from PIL import Image  # type: ignore
except ImportError:
    pytesseract = None
    Image = None
try:
    import fitz  # type: ignore  # PyMuPDF
except ImportError:
    fitz = None

_ESTENSIONI_IMG = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp"}


def _tesseract_ok() -> bool:
    if pytesseract is None:
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _lingua_ocr() -> Optional[str]:
    """Usa l'italiano se il pacchetto lingua è presente, altrimenti il default."""
    try:
        if "ita" in pytesseract.get_languages(config=""):
            return "ita"
    except Exception:
        pass
    return None


def ocr_disponibile_immagini() -> bool:
    return _tesseract_ok() and Image is not None


def ocr_disponibile_pdf() -> bool:
    return ocr_disponibile_immagini() and fitz is not None


def _ocr_immagine(img) -> str:
    lang = _lingua_ocr()
    return pytesseract.image_to_string(img, **({"lang": lang} if lang else {}))


def _modulo_presente(nome: str) -> bool:
    try:
        __import__(nome)
        return True
    except Exception:
        return False


def stato_dipendenze() -> List[Tuple[str, str, bool, str, str]]:
    """Elenco (nome, gruppo, disponibile, descrizione, comando) delle dipendenze.
    gruppo: 'pdf' = utile per i PDF, 'opz' = facoltativa."""
    voci: List[Tuple[str, str, bool, str, str]] = []
    voci.append(("pypdf", "pdf", pypdf is not None,
                 "Lettura dei file PDF (con testo).", "pip install pypdf"))

    voci.append(("spaCy + modello italiano", "opz", _NLP is not None,
                 "Riconoscimento dei nomi più preciso.",
                 "pip install spacy  e poi  python -m spacy download it_core_news_lg"))

    voci.append(("cryptography", "opz", _modulo_presente("cryptography"),
                 "Cifratura della mappa con password (--encrypt-map).",
                 "pip install cryptography"))

    ocr_ok = ocr_disponibile_pdf()
    libs_ocr = (pytesseract is not None and Image is not None and fitz is not None)
    if libs_ocr and not _tesseract_ok():
        descr_ocr = "OCR per PDF scansionati e immagini (librerie OK, manca il motore Tesseract)."
    elif ocr_ok and _lingua_ocr() != "ita":
        descr_ocr = "OCR per PDF scansionati e immagini (attivo, ma senza la lingua italiana)."
    else:
        descr_ocr = "OCR per PDF scansionati e immagini."
    voci.append(("OCR: pytesseract + pillow + pymupdf + motore Tesseract", "opz", ocr_ok,
                 descr_ocr,
                 "pip install pytesseract pymupdf pillow  +  motore Tesseract con lingua 'ita'"))
    return voci


def stampa_stato_dipendenze() -> None:
    voci = stato_dipendenze()
    def riga(nome, ok, descr, cmd):
        print(f"  {'[ OK ]' if ok else '[MANCA]'}  {nome}")
        print(f"          {descr}")
        if not ok:
            print(f"          -> {cmd}")
    print("-" * 70)
    print("DIPENDENZE")
    print(" I file .txt funzionano sempre, senza installare nulla.")
    print("\n Utile per i PDF con testo:")
    for nome, g, ok, descr, cmd in voci:
        if g == "pdf":
            riga(nome, ok, descr, cmd)
    print("\n Facoltative (migliorano il programma):")
    for nome, g, ok, descr, cmd in voci:
        if g == "opz":
            riga(nome, ok, descr, cmd)
    print("-" * 70)


PLACEHOLDER_RE = re.compile(r"\[[A-Z_]+_\d+\]")

# Classi di caratteri: esclusi × (U+00D7) e ÷ (U+00F7) dai range accentati.
_MAIU = "A-ZÀ-ÖØ-Ý"
_LETT = "A-Za-zÀ-ÖØ-öø-ÿ"

MESI = ("gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|"
        "agosto|settembre|ottobre|novembre|dicembre")
DATA_NUM = r"\b\d{1,2}[/\.\-]\d{1,2}[/\.\-]\d{2,4}\b"
DATA_TESTUALE = rf"\b\d{{1,2}}\s+(?:{MESI})\s+\d{{4}}\b"
DATA_QUALSIASI = rf"(?:{DATA_NUM}|{DATA_TESTUALE})"

_PAROLA = rf"(?:[{_MAIU}][{_LETT}]{{1,}}|[{_MAIU}][’'][{_MAIU}][{_LETT}]{{1,}})"

# Parole che non possono far parte di un nome di persona (confronto case-insensitive).
_NON_PERSONA = {
    # --- enti, uffici, organi ---
    "tribunale","corte","cassazione","suprema","costituzionale","appello","giudice",
    "giudici","comune","provincia","regione","regionale","statale","nazionale","comunale",
    "provinciale","ministero","ministro","agenzia","entrate","dogane","repubblica","banca",
    "condominio","societa","società","srl","spa","sas","snc","cooperativa","coop","fondazione",
    "associazione","ente","enti","istituto","istituti","camera","commercio","consiglio",
    "ordine","forense","foro","collegio","sezione","sezioni","distretto","circondario",
    "procura","procuratore","questura","prefettura","comando","carabinieri","polizia",
    "guardia","finanza","inps","inail","asl","usl","azienda","sanitaria","ospedale",
    # --- ruoli processuali / parti ---
    "convenuto","convenuta","convenuti","attore","attrice","attori","ricorrente","ricorrenti",
    "resistente","resistenti","appellante","appellato","opponente","opposto","controparte",
    "parte","parti","terzo","chiamato","intervenuto","creditore","debitore","istante",
    "rappresentato","rappresentata","rappresentante","difeso","difesa","difensore","patrocinio",
    "contro","oggetto","nonche","nonché",
    # --- atti / documenti / fasi ---
    "atto","atti","citazione","ricorso","comparsa","memoria","memorie","difensiva","conclusioni",
    "conclusionale","replica","note","nota","verbale","udienza","provvedimento","sentenza",
    "ordinanza","decreto","ingiuntivo","precetto","pignoramento","perizia","consulenza",
    "relazione","tecnica","perizale","allegato","allegati","documento","documenti","fascicolo",
    "istanza","eccezione","domanda","deposito","notifica","notificazione","costituzione",
    # --- concetti / branche del diritto ---
    "codice","fiscale","partita","iva","civile","penale","amministrativo","tributario",
    "societario","fallimentare","processuale","procedura","diritto","legge","leggi","norma",
    "norme","normativa","regolamento","articolo","art","comma","commi","principio","principi",
    "clausola","clausole","contratto","contratti","locazione","comodato","compravendita",
    "commerciale","costituzione","italiana","italiano","europea","europeo","giustizia",
    "pubblico","ministero","cancelleria","cancelliere","ruolo","registro","generale",
    "prima","primo","seconda","secondo","terza","terzo","quarta","quarto","quinta","quinto",
    "ordinario","ordinaria","speciale","straordinario","unico","unica",
    # --- indirizzi / topografia / cortesia ---
    "via","piazza","viale","corso","largo","vicolo","strada","località","contrada","borgo",
    "frazione","san","santa","santo","studio","legale","ufficio","residente","domiciliato",
    "domiciliata","nato","nata","nascita","spettabile","egregio","egregia","chiarissimo",
    "illustrissimo","onorevole","gentile","signor","signora","signori",
    # --- visura/dati catastali: intestazioni ed etichette di tabella ---
    "catasto","catastale","catastali","visura","terreni","terreno","fabbricati","fabbricato",
    "immobile","immobili","intestazione","intestatario","intestatari","dati","dato","identificativi",
    "identificativo","classamento","informazioni","informazione","foglio","particella","particelle",
    "subalterno","sub","porzione","porz","qualità","qualita","classe","superficie","reddito","redditi",
    "dominicale","agrario","agraria","anagrafici","anagrafico","anagrafe","tributaria","tributario",
    "oneri","reali","derivanti","frazionamento","scrittura","privata","sede","registrazione","volume",
    "ulteriori","ulteriore","euro","mappale","planimetria","fine","ur",
}
# Preposizioni/articoli: non possono essere la PRIMA parola di un nome di persona.
_PREP_ARTICOLI = {
    "di","de","del","dello","della","dei","degli","delle","da","dal","dalla",
    "lo","la","le","il","i","gli","l","d","e","ed","ai","al","alla","con","per",
}
_NON_PERSONA_RX = r"(?i:(?!(?:" + "|".join(map(re.escape, sorted(_NON_PERSONA))) + r")\b))"
_PAROLA_PERSONA = rf"{_NON_PERSONA_RX}{_PAROLA}"
# Separatore tra le parole di un nome: UN solo spazio, oppure un trattino.
# Evita di unire colonne di tabella separate da più spazi (es. visure catastali).
_SEP_NOME = r"(?:[ ]|[ ]?-[ ]?)"
NOME_PERSONA_1_4 = rf"{_PAROLA_PERSONA}(?:{_SEP_NOME}{_PAROLA_PERSONA}){{0,3}}"
NOME_PERSONA_2_4 = rf"{_PAROLA_PERSONA}(?:{_SEP_NOME}{_PAROLA_PERSONA}){{1,3}}"

# FIX v4: un toponimo può iniziare con al più 2 preposizioni ("via di Ripetta",
# "Piazza della Repubblica") ma DEVE poi contenere una parola con la maiuscola:
# "corso di causa", "via del tutto", "via di definizione" non matchano più.
_TOPO_PREP = r"(?:della|delle|degli|dei|del|di|da|san|santa|santo|ss\.)"
_TOPO_WORD = rf"(?:[{_MAIU}][{_LETT}'’]{{1,}}|{_TOPO_PREP})"
TOPONIMO = rf"(?:{_TOPO_PREP}[ ]+){{0,2}}[{_MAIU}][{_LETT}'’]{{1,}}(?:[ \-]+{_TOPO_WORD}){{0,5}}"

TITOLI = (r"Avv(?:\.|ocato|ocata)?|Dott\.ssa|Dott|Dr|Sig\.ra|Sig\.na|Sig|Ing|Geom|Arch|Rag|"
          r"Prof\.ssa|Prof|On|Notaio|Notaia|Giudice|C\.T\.U\.|CTU|C\.T\.P\.|CTP|P\.M\.|PM|"
          r"G\.I\.|G\.U\.")

NON_PERSONA_WORDS = {w.capitalize() for w in _NON_PERSONA} | {w.upper() for w in _NON_PERSONA}


@dataclass(frozen=True)
class PatternDef:
    categoria: str
    pattern: str
    flags: int = 0
    gruppo: Optional[str] = None
    categoria_placeholder: Optional[str] = None


PATTERN_STRUTTURATI: List[PatternDef] = [
    PatternDef("CODICE_FISCALE",
        r"\b[A-Z]{6}[0-9LMNPQRSTUV]{2}[ABCDEHLMPRST][0-9LMNPQRSTUV]{2}"
        r"[A-Z][0-9LMNPQRSTUV]{3}[A-Z]\b", re.IGNORECASE),
    PatternDef("IBAN",
        r"\bIT\d{2}[ ]?[A-Z][ ]?(?:\d[ ]?){10}(?:[0-9A-Z][ ]?){12}\b", re.IGNORECASE),
    PatternDef("PEC",
        r"\b(?:pec|posta\s+elettronica\s+certificata)[:\s]*[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b",
        re.IGNORECASE),
    PatternDef("EMAIL", r"\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b"),
    PatternDef("PARTITA_IVA",
        r"\b(?:P(?:artita)?\.?\s?IVA|P\.IVA)[:\s n\.]*\d{11}\b", re.IGNORECASE),
    PatternDef("TELEFONO",   # con etichetta esplicita
        r"\b(?:tel(?:efono)?|cell(?:ulare)?|mobile|fax|recapito)\.?\s*[:\-]?\s*"
        r"(?:\+39\s*)?(?:0\d{1,4}|3\d{2})[\s.\-/]?\d{5,8}\b", re.IGNORECASE),
    PatternDef("TELEFONO",   # cellulare senza etichetta, ancorato ai prefissi mobili
        r"(?<!\d)(?:\+39\s*)?3[2-9]\d[\s.\-/]?\d{6,7}\b", re.IGNORECASE),
    PatternDef("NUMERO_RUOLO",
        r"\bR\.?\s*G\.?\s*(?:N\.?\s*R\.?|E\.?|A\.?\s*C\.?|Trib\.?)?\s*(?:n\.?|nr\.?|numero)?\s*"
        r"\d{1,8}\s*/\s*\d{2,4}\b", re.IGNORECASE),
    PatternDef("NUMERO_PROVVEDIMENTO",
        r"\b(?:sentenza|ordinanza|decreto(?:\s+ingiuntivo)?|provvedimento|procedimento)"
        r"\s*(?:n\.?|nr\.?|numero)?\s*\d{1,8}\s*/\s*\d{2,4}\b", re.IGNORECASE),
    PatternDef("REPERTORIO_NOTARILE",
        r"\b(?:rep(?:ertorio)?\.|racc(?:olta)?\.)\s*(?:n\.?|nr\.?|numero)?\s*\d{1,10}\b",
        re.IGNORECASE),
    PatternDef("DOCUMENTO_IDENTITA",
        r"\b(?:carta\s+d[’' ]identità|carta\s+di\s+identita|c\.\s*i\.|patente|passaporto|"
        r"permesso\s+di\s+soggiorno|documento\s+d[’' ]identità)\s*"
        r"(?:n\.?|nr\.?|numero)?\s*[A-Z0-9]{5,20}\b", re.IGNORECASE),
    PatternDef("TESSERA_SANITARIA",
        r"\b(?:tessera\s+sanitaria|codice\s+tessera\s+sanitaria|t\.s\.)\s*"
        r"(?:n\.?|nr\.?|numero)?\s*[A-Z0-9]{8,25}\b", re.IGNORECASE),
    PatternDef("TARGA", r"\b[A-Z]{2}\s?\d{3}\s?[A-Z]{2}\b"),
    PatternDef("DATI_CATASTALI",
        r"\b(?:foglio|particella|mappale|subalterno|fg\.?|p\.lla|mapp\.?|sub\.?|part\.|f\.)\s*"
        r"(?:n\.?|nr\.?|numero)?\s*\d+[A-Za-z0-9/\-]*\b", re.IGNORECASE),
    PatternDef("COORDINATA",
        r"\b(?:lat(?:itudine)?|lng|long(?:itudine)?|coordinate?)\s*[:=]?\s*-?\d{1,3}[\.,]\d+\b",
        re.IGNORECASE),
    PatternDef("DATA_NASCITA",
        rf"\b(?:nato|nata|nascita)\b[^\n]{{0,80}}?\b(?:il|data)\s+(?P<data>{DATA_QUALSIASI})",
        re.IGNORECASE, gruppo="data", categoria_placeholder="DATA_NASCITA"),
    # FIX v4: lookahead esteso — copre fine riga (\r\n), ":" e le congiunzioni
    # "e"/"ed" ("nata a Milano e residente..."); il vecchio "$" senza MULTILINE
    # valeva solo a fine testo.
    PatternDef("LUOGO_NASCITA",
        rf"\b(?i:nato|nata)\s+(?i:a|in)\s+(?P<luogo>{TOPONIMO})"
        rf"(?=\s*(?:,|\(|;|\.|:|[\r\n]|$)|\s(?i:il|in|e|ed)\b)",
        0, gruppo="luogo", categoria_placeholder="LUOGO_NASCITA"),
    PatternDef("INDIRIZZO",
        rf"\b(?i:Via|V\.le|Viale|Piazza|P\.zza|Corso|C\.so|Largo|Vicolo|Strada|"
        rf"Località|Loc\.|Contrada|C\.da|Borgo|Frazione)\s+{TOPONIMO}"
        rf"(?:\s*,?\s*(?:n\.?|num\.?|numero)?\s*\d+[A-Za-z]?)?", 0),
    PatternDef("CONDOMINIO",
        rf"\bCondominio\s+(?:denominato\s+)?[\"“]?{TOPONIMO}[\"”]?", 0),
    PatternDef("SOCIETA",
        rf"\b[{_MAIU}][{_LETT}0-9&'’\.\-]*(?:\s+[{_MAIU}0-9][{_LETT}0-9&'’\.\-]*){{0,5}}\s+"
        r"(?i:S\.?\s?R\.?\s?L\.?|S\.?\s?P\.?\s?A\.?|S\.?\s?N\.?\s?C\.?|S\.?\s?A\.?\s?S\.?|"
        r"Soc\.?\s?Coop\.?|Società\s+Cooperativa|Cooperativa|Associazione|Fondazione)\b"),
    PatternDef("PERSONA_CON_TITOLO",
        rf"\b(?P<titolo>(?i:{TITOLI})\.?)\s+(?P<nome>{NOME_PERSONA_1_4})",
        0, gruppo="nome", categoria_placeholder="PERSONA"),
]

PATTERN_DATE_GENERICHE: List[PatternDef] = [
    PatternDef("DATA", DATA_NUM),
    PatternDef("DATA", DATA_TESTUALE, re.IGNORECASE),
]

# FIX v4: l'audit copre TUTTI i pattern di una categoria (la v3 teneva solo il
# primo: il pattern dei cellulari senza etichetta era escluso dall'audit).
_BY_CAT: Dict[str, List[PatternDef]] = {}
for _d in PATTERN_STRUTTURATI:
    _BY_CAT.setdefault(_d.categoria, []).append(_d)

def _audit_da_categoria(nome_audit: str, categoria: str) -> List[PatternDef]:
    return [PatternDef(nome_audit, d.pattern, d.flags) for d in _BY_CAT[categoria]]

AUDIT_PATTERNS: List[PatternDef] = (
    _audit_da_categoria("CODICE_FISCALE_RESIDUO", "CODICE_FISCALE")
    + _audit_da_categoria("IBAN_RESIDUO", "IBAN")
    + _audit_da_categoria("EMAIL_RESIDUA", "EMAIL")
    + _audit_da_categoria("PARTITA_IVA_RESIDUA", "PARTITA_IVA")
    + _audit_da_categoria("TELEFONO_RESIDUO", "TELEFONO")
    + _audit_da_categoria("TARGA_RESIDUA", "TARGA")
    + _audit_da_categoria("DATI_CATASTALI_RESIDUI", "DATI_CATASTALI")
    + [
        # FIX v4: usa TOPONIMO (il vecchio "\w+" IGNORECASE segnalava
        # "via preliminare", "in corso di causa" ecc.)
        PatternDef("INDIRIZZO_RESIDUO",
            rf"\b(?i:Via|Viale|Piazza|Corso|Largo|Vicolo|Strada|Località|Contrada)\s+{TOPONIMO}", 0),
        PatternDef("DATA_NASCITA_RESIDUA",
            rf"\b(?:nato|nata|nascita)\b[^\n]{{0,80}}?(?:{DATA_QUALSIASI})", re.IGNORECASE),
        # FIX v4: nuova categoria — i luoghi di nascita sfuggiti erano invisibili.
        PatternDef("LUOGO_NASCITA_RESIDUO",
            rf"\b(?i:nato|nata)\s+(?i:a|in)\s+{TOPONIMO}", 0),
    ]
)


# ---- sicurezza/mappa ----
def scrivi_file_privato(percorso: Path, contenuto: str) -> None:
    percorso.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(str(percorso), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(contenuto)
    try:
        os.chmod(percorso, stat.S_IRUSR | stat.S_IWUSR)
    except Exception:
        pass

def _crypto():
    try:
        from cryptography.fernet import Fernet
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    except ImportError as exc:
        raise RuntimeError("Per --encrypt-map serve: pip install cryptography") from exc
    return Fernet, hashes, PBKDF2HMAC

def _deriva_chiave(password: str, salt: bytes, iterations: int = 480_000) -> bytes:
    _, hashes, PBKDF2HMAC = _crypto()
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=iterations)
    return base64.urlsafe_b64encode(kdf.derive(password.encode("utf-8")))

def cifra_json(dati: Dict[str, str], password: str) -> str:
    Fernet, _, _ = _crypto()
    salt = os.urandom(16); iterations = 480_000
    token = Fernet(_deriva_chiave(password, salt, iterations)).encrypt(
        json.dumps(dati, ensure_ascii=False, indent=2).encode("utf-8"))
    return json.dumps({"encrypted": True, "kdf": "PBKDF2HMAC-SHA256",
        "iterations": iterations, "salt_b64": base64.b64encode(salt).decode("ascii"),
        "payload_b64": token.decode("ascii"),
        "created_at": datetime.now().isoformat(timespec="seconds")}, ensure_ascii=False, indent=2)

def decifra_json(pkt: Dict[str, object], password: str) -> Dict[str, str]:
    Fernet, _, _ = _crypto()
    salt = base64.b64decode(str(pkt["salt_b64"]))
    # FIX v4: minimo di iterazioni (un file manomesso non può degradare il KDF).
    iterations = int(pkt.get("iterations", 480_000))
    if iterations < 100_000:
        raise ValueError("Numero di iterazioni KDF troppo basso: mappa sospetta o manomessa.")
    chiave = _deriva_chiave(password, salt, iterations)
    return json.loads(Fernet(chiave).decrypt(str(pkt["payload_b64"]).encode("ascii")).decode("utf-8"))

def chiedi_password_cifratura() -> str:
    """FIX v4: chiamata PRIMA dell'elaborazione, così un errore di password
    non fa perdere la mappa a lavoro già svolto."""
    pw = getpass.getpass("Password per cifrare la mappa: ")
    if pw != getpass.getpass("Conferma password: "):
        raise RuntimeError("Le password non coincidono. Nessun file elaborato.")
    if not pw.strip():
        raise RuntimeError("Password vuota non ammessa. Nessun file elaborato.")
    return pw

def salva_mappa(mappa: Dict[str, str], percorso: Path, password: Optional[str] = None) -> Path:
    if password is not None:
        if not str(percorso).endswith(".enc"):
            percorso = percorso.with_name(percorso.name + ".enc")
        scrivi_file_privato(percorso, cifra_json(mappa, password))
    else:
        scrivi_file_privato(percorso, json.dumps(mappa, ensure_ascii=False, indent=2))
    return percorso

def carica_mappa(percorso: Path) -> Dict[str, str]:
    dati = json.loads(percorso.read_text(encoding="utf-8"))
    if isinstance(dati, dict) and dati.get("encrypted") is True:
        return decifra_json(dati, getpass.getpass("Password per decifrare la mappa: "))
    if not isinstance(dati, dict):
        raise ValueError("Formato mappa non valido.")
    return {str(k): str(v) for k, v in dati.items()}


class Anonimizzatore:
    def __init__(self, anonimizza_date=False, fallback_nomi=True):
        self.anonimizza_date = anonimizza_date
        self.fallback_nomi = fallback_nomi
        self.registro: Dict[str, str] = {}
        self.mappa: Dict[str, str] = {}
        self.contatori: Dict[str, int] = {}

    def _normalizza(self, categoria: str, valore: str) -> str:
        base = re.sub(r"\s+", " ", valore).strip().upper()
        if categoria in {"CODICE_FISCALE", "IBAN", "PARTITA_IVA"}:
            base = re.sub(r"\s+", "", base)
        if categoria == "TELEFONO":
            base = re.sub(r"[^0-9+]", "", base)
        return base

    def _segnaposto(self, categoria: str, valore: str) -> str:
        # FIX v4: strip dei soli spazi (":" e "-" facevano perdere caratteri
        # del testo originale al momento del ripristino).
        valore = valore.strip(" \t\u00a0")
        chiave = f"{categoria}::{self._normalizza(categoria, valore)}"
        if chiave in self.registro:
            return self.registro[chiave]
        self.contatori[categoria] = self.contatori.get(categoria, 0) + 1
        ph = f"[{categoria}_{self.contatori[categoria]}]"
        self.registro[chiave] = ph
        self.mappa[ph] = valore
        return ph

    def _segnaposto_preserva_bordi(self, categoria: str, valore: str) -> str:
        m = re.match(r"^(?P<pre>\s*)(?P<core>.*?)(?P<post>\s*)$", valore, flags=re.DOTALL)
        if not m:
            return self._segnaposto(categoria, valore)
        return m.group("pre") + self._segnaposto(categoria, m.group("core")) + m.group("post")

    def _fuori_placeholder(self, testo: str, fn: Callable[[str], str]) -> str:
        parti, ultimo = [], 0
        for m in PLACEHOLDER_RE.finditer(testo):
            parti.append(fn(testo[ultimo:m.start()])); parti.append(m.group(0)); ultimo = m.end()
        parti.append(fn(testo[ultimo:]))
        return "".join(parti)

    def _sostituisci(self, testo: str, d: PatternDef) -> str:
        cat = d.categoria_placeholder or d.categoria
        def seg(s: str) -> str:
            def repl(m):
                if d.gruppo:
                    # FIX v4: ricostruzione basata sugli span (str.replace poteva
                    # sostituire un'occorrenza identica precedente nel match).
                    ini, fin = m.span(d.gruppo)
                    base = m.start()
                    tutto = m.group(0)
                    return (tutto[:ini - base]
                            + self._segnaposto(cat, m.group(d.gruppo))
                            + tutto[fin - base:])
                return self._segnaposto_preserva_bordi(cat, m.group(0))
            return re.sub(d.pattern, repl, s, flags=d.flags)
        return self._fuori_placeholder(testo, seg)

    def _entita_persona_valida(self, v: str) -> bool:
        """Filtra i falsi positivi di spaCy (intestazioni, etichette, blocchi di testo)."""
        # Un nome di persona sta su una sola riga: gli a-capo indicano un blocco di layout.
        if "\n" in v or "\r" in v:
            return False
        parole = [p.strip(".,;:()[]{}\"'’") for p in re.split(r"[\s\-]+", v.strip()) if p.strip()]
        if not parole:
            return False
        # Contiene un termine istituzionale/giuridico/catastale: non è una persona.
        if any(p.lower() in _NON_PERSONA for p in parole):
            return False
        # Non può iniziare con preposizione/articolo.
        if parole[0].lower() in _PREP_ARTICOLI:
            return False
        # Troppe parole: probabile frammento di testo, non un nome.
        if len(parole) > 4:
            return False
        # Deve contenere almeno una parola con iniziale maiuscola (evita spazzatura).
        if not any(p[:1].isupper() for p in parole):
            return False
        return True

    def _nomi_ner(self, testo: str) -> str:
        if _NLP is None:
            return testo
        def seg(s: str) -> str:
            doc = _NLP(s); nuovo = s
            for ent in sorted(doc.ents, key=lambda e: e.start_char, reverse=True):
                if ent.label_ in {"PER", "PERSON"}:
                    v = ent.text.strip()
                    if v and not PLACEHOLDER_RE.fullmatch(v) and self._entita_persona_valida(v):
                        nuovo = nuovo[:ent.start_char] + self._segnaposto("PERSONA", v) + nuovo[ent.end_char:]
            return nuovo
        return self._fuori_placeholder(testo, seg)

    def _sembra_nome(self, valore: str, contesto: str = "") -> bool:
        parole = [p.strip(".,;:()[]{}\"'’") for p in re.split(r"[\s\-]+", valore.strip()) if p.strip()]
        if len(parole) < 2:
            return False
        if any(p.lower() in _NON_PERSONA for p in parole):
            return False
        if parole[0] in {"San", "Santa", "Santo"}:
            return False
        # Un nome di persona non inizia con preposizione/articolo (es. "DI Milano", "DELLA Corte")
        if parole[0].lower() in _PREP_ARTICOLI:
            return False
        if re.search(r"(?:tribunale|corte|comune|provincia|regione|agenzia|banca|ministero|"
                     r"foro|camera|consiglio|procura|sezione|ordine|collegio|prefettura|questura)"
                     r"(?:\s+\w+){0,2}\s+(?:di|del|della|dello|dei|degli|delle)\s*$",
                     contesto[-40:].lower()):
            return False
        return True

    def _nomi_fallback(self, testo: str) -> str:
        pat = re.compile(rf"\b(?P<nome>{NOME_PERSONA_2_4})\b")
        def seg(s: str) -> str:
            out, ultimo = [], 0
            for m in pat.finditer(s):
                cand = m.group("nome")
                if self._sembra_nome(cand, s[max(0, m.start()-50):m.start()]):
                    out.append(s[ultimo:m.start()]); out.append(self._segnaposto("PERSONA", cand)); ultimo = m.end()
            out.append(s[ultimo:])
            return "".join(out)
        return self._fuori_placeholder(testo, seg)

    def anonimizza(self, testo: str) -> str:
        for d in PATTERN_STRUTTURATI:
            testo = self._sostituisci(testo, d)
        testo = self._nomi_ner(testo)
        if self.fallback_nomi:
            testo = self._nomi_fallback(testo)
        if self.anonimizza_date:
            for d in PATTERN_DATE_GENERICHE:
                testo = self._sostituisci(testo, d)
        return testo

    def audit_residui(self, testo: str, max_esempi: int = 10) -> Dict[str, List[str]]:
        residui: Dict[str, List[str]] = {}
        def cerca(s: str, d: PatternDef) -> str:
            es = residui.setdefault(d.categoria, [])
            for m in re.finditer(d.pattern, s, flags=d.flags):
                # FIX v4: controllo PRIMA dell'append (il limite poteva essere superato).
                if len(es) >= max_esempi:
                    break
                v = re.sub(r"\s+", " ", m.group(0)).strip()
                if v and v not in es:
                    es.append(v[:180])
            return s
        for d in AUDIT_PATTERNS:
            self._fuori_placeholder(testo, lambda s, d=d: cerca(s, d))
        nomi: List[str] = []
        pat = re.compile(rf"\b(?P<nome>{NOME_PERSONA_2_4})\b")
        def an(s: str) -> str:
            for m in pat.finditer(s):
                if len(nomi) >= max_esempi:
                    break
                c = m.group("nome")
                if self._sembra_nome(c, s[max(0, m.start()-50):m.start()]) and c not in nomi:
                    nomi.append(c[:180])
            return s
        self._fuori_placeholder(testo, an)
        if nomi: residui["POSSIBILE_PERSONA_RESIDUA"] = nomi
        return {k: v for k, v in residui.items() if v}


# ---- I/O ----
def leggi_txt(p: Path) -> str:
    # FIX v4: cp1252 PRIMA di latin-1 (latin-1 non fallisce mai: il ramo cp1252
    # era irraggiungibile e ’ – ecc. venivano decodificati come caratteri di
    # controllo, rompendo i pattern sui nomi con apostrofo tipografico).
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return p.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("Impossibile decodificare il file di testo.")

def _ocr_pagine_pdf(percorso: Path, testi: List[str], indici: List[int]) -> List[str]:
    """Rasterizza e riconosce via OCR le pagine PDF prive di testo estraibile."""
    import io
    try:
        doc = fitz.open(str(percorso))
    except Exception:
        return testi
    lang = _lingua_ocr()
    kw = {"lang": lang} if lang else {}
    for i in indici:
        try:
            pix = doc.load_page(i).get_pixmap(dpi=300)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            t = pytesseract.image_to_string(img, **kw)
            if t and t.strip():
                testi[i] = t
                print(f"[OCR] Pagina {i + 1}: testo riconosciuto tramite OCR.")
        except Exception:
            pass
    try:
        doc.close()
    except Exception:
        pass
    return testi

def leggi_pdf(p: Path, force_ocr: bool = False) -> str:
    if pypdf is None:
        raise RuntimeError("pypdf non installato. Usa: pip install pypdf")
    reader = pypdf.PdfReader(str(p))
    testi: List[str] = []
    for pagina in reader.pages:
        if force_ocr:
            testi.append("")
            continue
        try:
            t = pagina.extract_text(extraction_mode="layout")
        except Exception:
            t = pagina.extract_text() or ""
        testi.append(t or "")
    # Pagine senza testo: probabile scansione -> OCR (se disponibile).
    da_ocr = [i for i, t in enumerate(testi) if not t.strip()]
    if da_ocr and ocr_disponibile_pdf():
        testi = _ocr_pagine_pdf(p, testi, da_ocr)
    parti = [f"\n--- PAGINA {i + 1} ---\n{t}" for i, t in enumerate(testi)]
    testo = "\n".join(parti)
    if not testo.strip():
        if not ocr_disponibile_pdf():
            raise ValueError(
                "PDF senza testo (probabile scansione) e OCR non disponibile. "
                "Installa: pip install pytesseract pymupdf pillow, più il motore "
                "Tesseract con la lingua italiana. Vedi il manuale.")
        raise ValueError("PDF vuoto: nessun testo riconosciuto nemmeno con l'OCR.")
    return testo

def leggi_immagine(p: Path) -> str:
    if not ocr_disponibile_immagini():
        raise ValueError(
            "Per leggere le immagini serve l'OCR. Installa: pip install pytesseract "
            "pillow, più il motore Tesseract con la lingua italiana. Vedi il manuale.")
    try:
        img = Image.open(str(p))
    except Exception as exc:
        raise ValueError(f"Impossibile aprire l'immagine: {exc}") from exc
    testo = _ocr_immagine(img)
    if not testo.strip():
        raise ValueError("Nessun testo riconosciuto nell'immagine.")
    return testo

def carica_contenuto(p: Path, force_ocr: bool = False) -> str:
    ext = p.suffix.lower()
    if ext == ".pdf": return leggi_pdf(p, force_ocr=force_ocr)
    if ext == ".txt": return leggi_txt(p)
    if ext in _ESTENSIONI_IMG: return leggi_immagine(p)
    raise ValueError(f"Formato non supportato: {ext}")

def file_elaborabile(p: Path) -> bool:
    if p.suffix.lower() not in ({".txt", ".pdf"} | _ESTENSIONI_IMG): return False
    n = p.name.lower()
    return not any(n.endswith(e) for e in (
        "_anonimizzato.txt","_ripristinato.txt","_residui_da_verificare.txt",
        "_mappa_privata.json","_mappa_privata.json.enc",
        "_report_privacy.json","mappa_fascicolo_privata.json","mappa_fascicolo_privata.json.enc",
        "report_fascicolo_privacy.json"))

class ErroreUso(Exception):
    """Errore d'uso da parte dell'utente: mostra un messaggio chiaro + l'help."""


def risolvi_bersagli(arg: Optional[str]) -> Tuple[List[Path], Optional[Path]]:
    if arg:
        p = Path(arg).expanduser().resolve()
        if p.is_dir():
            files = sorted([f for f in p.iterdir() if f.is_file() and file_elaborabile(f)])
            if not files:
                raise ErroreUso(f"La cartella '{p}' non contiene file .txt o .pdf da elaborare.")
            return files, p
        if p.is_file():
            if p.suffix.lower() not in ({".txt", ".pdf"} | _ESTENSIONI_IMG):
                raise ErroreUso(f"Formato non supportato: '{p.suffix or p.name}'. "
                                "Sono ammessi file .txt, .pdf o immagini (png, jpg, tiff...).")
            return [p], None
        raise ErroreUso(f"File o cartella inesistente: '{p}'. "
                        "Controlla il nome e il percorso.")
    for nome in ("atto_originale.txt", "atto_originale.pdf"):
        p = Path(nome).resolve()
        if p.exists(): return [p], None
    raise ErroreUso("Nessun file indicato e nessun 'atto_originale.txt/.pdf' trovato "
                    "nella cartella corrente. Indica il file da elaborare.")


_ETICHETTE_RESIDUI = {
    "CODICE_FISCALE_RESIDUO": "Codici fiscali",
    "IBAN_RESIDUO": "IBAN",
    "EMAIL_RESIDUA": "Email",
    "PARTITA_IVA_RESIDUA": "Partite IVA",
    "TELEFONO_RESIDUO": "Telefoni",
    "TARGA_RESIDUA": "Targhe",
    "INDIRIZZO_RESIDUO": "Indirizzi",
    "DATA_NASCITA_RESIDUA": "Date di nascita",
    "LUOGO_NASCITA_RESIDUO": "Luoghi di nascita",
    "DATI_CATASTALI_RESIDUI": "Dati catastali",
    "POSSIBILE_PERSONA_RESIDUA": "Possibili nomi di persona",
}

def formatta_residui(nome_atto: str, residui: Dict[str, List[str]]) -> str:
    tot = sum(len(v) for v in residui.values())
    righe = ["=" * 70,
             "RESIDUI DA VERIFICARE MANUALMENTE",
             f"Atto: {nome_atto}",
             f"Generato: {datetime.now().isoformat(timespec='seconds')}",
             "=" * 70, ""]
    if not tot:
        righe.append("Nessun residuo evidente rilevato dall'audit automatico.")
        righe.append("")
        righe.append("NOTA: l'audit non garantisce l'assenza totale di dati sensibili.")
        righe.append("Consigliata comunque una rilettura del file _anonimizzato.txt.")
        return "\n".join(righe) + "\n"
    righe.append(f"Trovati {tot} possibili residui, suddivisi per categoria.")
    righe.append("Controlla nel file _anonimizzato.txt se vanno rimossi a mano.")
    righe.append("")
    for cat, esempi in residui.items():
        titolo = _ETICHETTE_RESIDUI.get(cat, cat)
        righe.append(f"[{titolo}] ({len(esempi)})")
        for e in esempi:
            righe.append(f"  - {e}")
        righe.append("")
    return "\n".join(righe) + "\n"


def elabora(bersagli, cartella, anonimizza_date, fallback_nomi, encrypt_map, force_ocr=False):
    if not bersagli:
        print("[ERRORE] Nessun file da elaborare."); return
    # FIX v4: dipendenza e password verificate PRIMA di elaborare qualsiasi file.
    password: Optional[str] = None
    if encrypt_map:
        _crypto()
        password = chiedi_password_cifratura()
    anon = Anonimizzatore(anonimizza_date, fallback_nomi)
    report = {"created_at": datetime.now().isoformat(timespec="seconds"),
              "spacy_attivo": _NLP is not None, "fallback_nomi": fallback_nomi,
              "date_generiche": anonimizza_date, "ocr": ocr_disponibile_pdf(),
              "force_ocr": force_ocr, "files": {}}
    errori: Dict[str, str] = {}
    print("=" * 70); print("PSEUDO-ANONIMIZZATORE ATTI LEGALI v5"); print("=" * 70)
    stampa_stato_dipendenze()
    for p in bersagli:
        print(f"\n[INFO] Elaboro: {p.name}")
        # FIX v4: un file che fallisce (PDF scansionato, pypdf mancante...) non
        # abortisce più l'intero fascicolo prima del salvataggio della mappa.
        try:
            testo = anon.anonimizza(carica_contenuto(p, force_ocr=force_ocr))
            residui = anon.audit_residui(testo)
        except Exception as exc:
            errori[p.name] = str(exc)
            report["files"][p.name] = {"errore": str(exc)}
            print(f"[ERRORE] {p.name}: {exc} — file saltato.")
            continue
        out = p.with_name(p.stem + "_anonimizzato.txt"); out.write_text(testo, encoding="utf-8")
        res_file = p.with_name(p.stem + "_residui_da_verificare.txt")
        # FIX v4: il file dei residui contiene dati reali -> permessi 0600.
        scrivi_file_privato(res_file, formatta_residui(p.name, residui))
        report["files"][p.name] = {"output": out.name, "residui_file": res_file.name,
                                   "residui_potenziali": residui}
        print(f"[OK] {out.name}")
        tot = sum(len(v) for v in residui.values())
        if tot:
            print(f"[ATTENZIONE] {tot} possibili residui da verificare -> {res_file.name}")
        else:
            print(f"[OK] Audit: nessun residuo evidente (dettagli in {res_file.name}).")
    if cartella:
        base = cartella / f"{cartella.name}_mappa_privata.json"
        rep = cartella / f"{cartella.name}_report_privacy.json"
    else:
        base = bersagli[0].with_name(bersagli[0].stem + "_mappa_privata.json")
        rep = bersagli[0].with_name(bersagli[0].stem + "_report_privacy.json")
    mp = salva_mappa(anon.mappa, base, password)
    report["mappa"] = {"file": mp.name, "cifrata": password is not None,
                       "segnaposto": len(anon.mappa), "per_categoria": anon.contatori}
    if errori:
        report["errori"] = errori
    scrivi_file_privato(rep, json.dumps(report, ensure_ascii=False, indent=2))
    print("\n" + "-" * 70)
    print(f"[OK] Mappa privata: {mp.name} ({len(anon.mappa)} sostituzioni)")
    print(f"[OK] Report: {rep.name}")
    if errori:
        print(f"[ATTENZIONE] {len(errori)} file saltati per errori: {', '.join(errori)}")
    print("[AVVISO] Mappa, report e file _residui contengono dati reali:")
    print("         non inviarli all'AI, conservali offline.")
    print("=" * 70)


def ripristina(risposta_path, mappa_path, output=None):
    risposta = Path(risposta_path).expanduser().resolve()
    mappa = carica_mappa(Path(mappa_path).expanduser().resolve())
    testo = PLACEHOLDER_RE.sub(lambda m: mappa.get(m.group(0), m.group(0)),
                               risposta.read_text(encoding="utf-8"))
    # FIX v4: conserva l'estensione originale della risposta (es. .md).
    out = (Path(output).expanduser().resolve() if output
           else risposta.with_name(risposta.stem + "_ripristinato" + (risposta.suffix or ".txt")))
    out.write_text(testo, encoding="utf-8")
    print(f"[OK] Risposta ripristinata: {out}")


class _FormatterIT(argparse.HelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):
        return super().add_usage(usage, actions, groups, prefix or "uso: ")

def _input_nonvuoto(prompt: str) -> str:
    while True:
        v = input(prompt).strip()
        if v:
            return v
        print("  Inserisci un valore (o Ctrl+C per uscire).")

def _input_file_esistente(prompt: str) -> str:
    while True:
        v = _input_nonvuoto(prompt)
        if Path(v).expanduser().exists():
            return v
        print(f"  [!] Percorso inesistente: {v}")

def _chiedi_si_no(prompt: str, default: bool = False) -> bool:
    suff = " [S/n] " if default else " [s/N] "
    while True:
        v = input(prompt + suff).strip().lower()
        if not v:
            return default
        if v in ("s", "si", "sì", "y", "yes"):
            return True
        if v in ("n", "no"):
            return False
        print("  Rispondi 's' oppure 'n'.")

def _processi_windows() -> Dict[int, Tuple[int, str]]:
    """Mappa pid -> (pid_padre, nome_exe) di tutti i processi (solo Windows)."""
    import ctypes
    from ctypes import wintypes
    class PROCESSENTRY32(ctypes.Structure):
        _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
                    ("th32ProcessID", wintypes.DWORD),
                    ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
                    ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
                    ("th32ParentProcessID", wintypes.DWORD),
                    ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD),
                    ("szExeFile", ctypes.c_char * 260)]
    k = ctypes.windll.kernel32
    snap = k.CreateToolhelp32Snapshot(0x00000002, 0)  # TH32CS_SNAPPROCESS
    entry = PROCESSENTRY32(); entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
    info: Dict[int, Tuple[int, str]] = {}
    if k.Process32First(snap, ctypes.byref(entry)):
        while True:
            info[entry.th32ProcessID] = (entry.th32ParentProcessID,
                                         entry.szExeFile.decode(errors="ignore").lower())
            if not k.Process32Next(snap, ctypes.byref(entry)):
                break
    k.CloseHandle(snap)
    return info

def _avviato_da_shell_windows() -> bool:
    """True se il programma è stato lanciato da un terminale (cmd, PowerShell...)."""
    _SHELL = {"cmd.exe", "powershell.exe", "pwsh.exe", "wt.exe",
              "windowsterminal.exe", "bash.exe", "code.exe"}
    try:
        info = _processi_windows()
        mio = os.getpid()
        mio_nome = info.get(mio, (0, ""))[1]
        pid = info.get(mio, (0, ""))[0]
        visti = set()
        while pid and pid not in visti:
            visti.add(pid)
            ppid, nome = info.get(pid, (0, ""))
            if nome and nome != mio_nome:        # salta il bootloader PyInstaller
                if nome in _SHELL:
                    return True                  # avviato da una shell
                return False                     # explorer.exe o altro: doppio clic
            pid = ppid
    except Exception:
        pass
    return False

def _serve_pausa() -> bool:
    """Decide se attendere INVIO: solo quando il programma ha una finestra propria
    (doppio clic), non quando è lanciato da un terminale già aperto."""
    if os.environ.get("ANON_ATTI_NO_PAUSE"):
        return False
    if os.environ.get("ANON_ATTI_PAUSE"):
        return True
    try:
        if not sys.stdin.isatty() or not sys.stdout.isatty():
            return False                         # output rediretto/pipe: niente pausa
    except Exception:
        return False
    if os.name == "nt":
        return not _avviato_da_shell_windows()
    # macOS/Linux: non è possibile distinguere in modo affidabile il doppio clic
    # dal Terminale; il file .command apre comunque una finestra propria, quindi
    # la pausa è utile. Chi lancia da shell può disattivarla con ANON_ATTI_NO_PAUSE=1.
    return True

def _pausa_finale() -> None:
    """Tiene aperta la finestra quando il programma è avviato con un clic."""
    if not _serve_pausa():
        return
    try:
        input("\nPremi INVIO per chiudere...")
    except (EOFError, KeyboardInterrupt):
        pass

def modalita_interattiva() -> int:
    print("=" * 70)
    print("PSEUDO-ANONIMIZZATORE ATTI LEGALI v5 — modalità interattiva")
    print("=" * 70)
    print("Cosa vuoi fare?")
    print("  1) Anonimizzare un atto (o una cartella di atti)")
    print("  2) Ripristinare una risposta già anonimizzata")

    while True:
        scelta = input("Scelta [1/2]: ").strip()
        if scelta in ("1", "2"):
            break
        print("  Digita 1 o 2.")

    if scelta == "1":
        # Chiede il file/cartella e riprova in caso di percorso errato,
        # senza chiudere il programma.
        print("\nPuoi indicare DUE cose diverse:")
        print("  1. UN SINGOLO FILE  -> viene anonimizzato solo quel documento.")
        print("     Esempi:")
        print("       Windows:      C:\\Users\\Mario\\Documenti\\atto.pdf")
        print("       macOS/Linux:  /Users/mario/Documenti/atto.pdf")
        print("  2. UNA CARTELLA     -> vengono anonimizzati TUTTI i file")
        print("     (.txt, .pdf, immagini) contenuti al suo interno, con un'unica")
        print("     mappa per l'intero fascicolo.")
        print("     Esempi:")
        print("       Windows:      C:\\Users\\Mario\\Documenti\\Fascicolo_Rossi")
        print("       macOS/Linux:  /Users/mario/Documenti/Fascicolo_Rossi")
        print("Suggerimento: puoi anche TRASCINARE il file o la cartella in questa")
        print("finestra e premere INVIO: il percorso viene inserito da solo.")
        while True:
            percorso = _input_nonvuoto("\nFile o cartella da anonimizzare: ")
            try:
                bersagli, cartella = risolvi_bersagli(percorso)
                break
            except ErroreUso as exc:
                print(f"  [!] {exc}")

        if cartella:
            print(f"\n[INFO] Cartella riconosciuta: {len(bersagli)} file da elaborare.")
        else:
            print(f"\n[INFO] File singolo pronto per l'elaborazione.")
        anonimizza_date = _chiedi_si_no("Anonimizzare anche le date generiche?", default=False)
        fallback_nomi   = _chiedi_si_no("Usare il fallback regex per i nomi 'nudi'?", default=True)
        encrypt_map     = _chiedi_si_no("Cifrare la mappa con password?", default=False)
        force_ocr = False
        if ocr_disponibile_pdf():
            force_ocr = _chiedi_si_no("Forzare l'OCR su tutte le pagine dei PDF?", default=False)

        elabora(bersagli, cartella, anonimizza_date, fallback_nomi, encrypt_map, force_ocr)
        return 0

    # scelta == "2": ripristino
    print("\n--- Ripristino ---")
    risposta = _input_file_esistente("File con la risposta dell'AI (contiene i segnaposto): ")
    mappa    = _input_file_esistente("File mappa (_mappa_privata.json oppure .json.enc): ")
    out      = input("File di output (INVIO per il nome automatico): ").strip() or None

    ripristina(risposta, mappa, out)
    return 0
def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else list(argv)

    ap = argparse.ArgumentParser(
        description="Pseudo-anonimizzatore atti legali italiani (v5).",
        formatter_class=_FormatterIT, add_help=False)
    ap._positionals.title = "argomenti posizionali"
    ap._optionals.title = "opzioni"
    ap.add_argument("-h", "--help", "--aiuto", action="help", default=argparse.SUPPRESS,
                    help="Mostra questo messaggio di aiuto ed esci.")
    ap.add_argument("percorso", nargs="?", help="File .txt/.pdf/immagine o cartella.")
    ap.add_argument("--date", action="store_true", help="Anonimizza anche le date generiche.")
    ap.add_argument("--no-fallback-nomi", action="store_true", help="Disattiva il fallback regex per i nomi.")
    ap.add_argument("--encrypt-map", action="store_true", help="Cifra la mappa (pip install cryptography).")
    ap.add_argument("--ocr", action="store_true", help="Forza l'OCR su tutte le pagine dei PDF (per atti scansionati).")
    ap.add_argument("--check", action="store_true", help="Mostra lo stato dei pacchetti installati ed esce.")
    ap.add_argument("--ripristina", nargs=2, metavar=("RISPOSTA_AI", "MAPPA"), help="Riapplica i dati reali.")
    ap.add_argument("--output", help="Output per --ripristina.")
    ap.add_argument("-i", "--interattivo", action="store_true", help="Avvia la modalità interattiva guidata.")
    args = ap.parse_args(argv)

    try:
        if args.check:
            print("PSEUDO-ANONIMIZZATORE ATTI LEGALI v5 — stato dei pacchetti")
            stampa_stato_dipendenze()
            return 0
        # Nessun argomento oppure -i esplicito -> modalità guidata.
        # Gestisce da sé errori e pausa finale, cos' avviata con un clic
        # la finestra resta aperta e mostra sempre i risultati.
        if args.interattivo or not argv:
            codice = 0
            try:
                codice = modalita_interattiva()
            except KeyboardInterrupt:
                print("\n[ERRORE] Interrotto."); codice = 130
            except Exception as exc:
                print(f"[ERRORE] {exc}", file=sys.stderr); codice = 1
            _pausa_finale()
            return codice
        if args.ripristina:
            ripristina(args.ripristina[0], args.ripristina[1], args.output); return 0
        bersagli, cartella = risolvi_bersagli(args.percorso)
        elabora(bersagli, cartella, args.date, not args.no_fallback_nomi, args.encrypt_map, args.ocr)
        return 0
    except KeyboardInterrupt:
        print("\n[ERRORE] Interrotto."); return 130
    except (ErroreUso, FileNotFoundError) as exc:
        print(f"[ERRORE] {exc}\n", file=sys.stderr)
        ap.print_help(sys.stderr)
        return 2
    except Exception as exc:
        print(f"[ERRORE] {exc}", file=sys.stderr); return 1

if __name__ == "__main__":
    raise SystemExit(main())