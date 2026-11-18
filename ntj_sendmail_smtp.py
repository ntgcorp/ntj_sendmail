"""
ntJobsApp per invio email tramite SMTP.
Parametro: file .ini contenente la configurazione e il percorso del file JSON con le mail.
"""

import json
import time
import os
import sys

# Librerie esterne
import aiSys
from acJobsApp import acJobsApp
from ncMailSimple import ncMailSimple

# ----------------------------------------------------------------------
# Variabili globali
# ----------------------------------------------------------------------
jData = None                # Istanza di acJobsApp
dictMail = {}               # Dizionario delle mail lette dal JSON
dictMailCurrent = {}        # Mail corrente in spedizione
nCounter = 0                # Contatore mail inviate
sMailEngine = ""            # "SMTP" per questo script
sFileIn = ""                # Nome del file JSON
sTimestamp = ""             # Timestamp aggiornato
sTimestampStart = ""        # Timestamp di avvio
sDelimiter = ","            # Delimitatore per liste
nMaxSizeAttach = 5000000    # Dimensione massima allegati (5 MB)
nWaitStart = 3              # Secondi di attesa prima di iniziare
nWaitMail = 5               # Secondi di attesa tra una mail e l'altra

# Specifico per SMTP
jMail = None                # Istanza di ncMailSimple
bSmtpStart = False          # Flag avvio SMTP

# ----------------------------------------------------------------------
# Funzioni di supporto
# ----------------------------------------------------------------------
def format_timedelta(seconds):
    """Converte secondi in formato HH:MM:SS."""
    # Forziamo il cast a intero per eliminare qualsiasi virgola passata da aiSys
    seconds = int(seconds) 
    
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"

# ----------------------------------------------------------------------
# Funzioni applicative
# ----------------------------------------------------------------------
def app_argread():
    """Legge il file JSON specificato in sFileIn e popola dictMail."""
    global dictMail, sTimestamp
    sProc = "app_argread"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    try:
        with open(sFileIn, 'r', encoding='utf-8') as f:
            data = json.load(f)
        dictMail = {}
        for key, value in data.items():
            new_key = key.strip().upper()
            if isinstance(value, dict):
                new_value = {}
                for k, v in value.items():
                    new_k = k.strip().upper()
                    if isinstance(v, str):
                        v = v.strip().strip('"')
                    elif isinstance(v, list):
                        v = [item.strip() if isinstance(item, str) else item for item in v]
                    new_value[new_k] = v
                dictMail[new_key] = new_value
            else:
                dictMail[new_key] = value
        sTimestamp = aiSys.Timestamp()
    except Exception as e:
        sResult = f"Errore lettura file json: {str(e)}"
        return aiSys.ErrorProc(sResult, sProc)
    return sResult

def app_checkmail():
    """Valida le mail presenti in dictMail."""
    global dictMail, nMaxSizeAttach
    sProc = "app_checkmail"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    try:
        for sKey, dictTemp in dictMail.items():
            if sKey == "CONFIG":
                continue

            if "TO" not in dictTemp or not dictTemp["TO"]:
                sResult = f"Campo TO obbligatorio mancante o vuoto - ID: {sKey}"
                return aiSys.ErrorProc(sResult, sProc)

            for field in ["TO", "CC", "CCN"]:
                if field in dictTemp:
                    for email in dictTemp[field]:
                        email_clean = email.strip()
                        if not aiSys.isEmail(email_clean):
                            sResult = f"Email non valida: {email_clean} - ID: {sKey}"
                            return aiSys.ErrorProc(sResult, sProc)

            if "ATTACH" in dictTemp:
                for sFile in dictTemp["ATTACH"]:
                    sFile_clean = sFile.strip()
                    if not aiSys.isValidPath(sFile_clean):
                        sResult = f"File attach non presente: {sFile_clean} - ID: {sKey}"
                        return aiSys.ErrorProc(sResult, sProc)
                    if os.path.getsize(sFile_clean) > nMaxSizeAttach:
                        sResult = f"File {sFile_clean} supera dimensione massima - ID: {sKey}"
                        return aiSys.ErrorProc(sResult, sProc)
        return sResult
    except Exception as e:
        return aiSys.ErrorProc(str(e), sProc)

def app_start_smtp():
    """Inizializza la connessione SMTP."""
    global jMail, bSmtpStart
    sProc = "app_start_smtp"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""

    jMail = ncMailSimple()
    smtp_port_str = jData.Config("SMTP.PORT")
    nPort = 25
    if smtp_port_str:
        nPort = aiSys.StringToNum(smtp_port_str)

    sResult = jMail.Start(
        sSmtp_User_par=jData.Config("SMTP.USER"),
        sSmtp_Password_par=jData.Config("SMTP.PASSWORD"),
        sSmtp_Host_par=jData.Config("SMTP.SERVER"),
        nSmtp_Port_par=nPort,
        bSSL_par=aiSys.StringBool(jData.Config("SMTP.SSL"))
    )

    if sResult == "":
        bSmtpStart = True
    return sResult

def app_sendmail_smtp():
    """Invia una singola mail usando SMTP."""
    global dictMailCurrent, sTimestamp, bSmtpStart
    sProc = "app_sendmail_smtp"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")

    if not bSmtpStart:
        return aiSys.ErrorProc("SMTP non avviato", sProc)

    try:
        # Determina formato
        format_type = dictMailCurrent.get("FORMAT", "").upper()
        sFormat = "HTML" if format_type == "HTML" else "TXT"

        sResult = jMail.Send(
            asTo_par=dictMailCurrent.get("TO", []),
            sSubject_par=dictMailCurrent.get("SUBJECT", ""),
            asAttach_par=dictMailCurrent.get("ATTACH", []),
            sFormat_par=sFormat,
            sBody_par=dictMailCurrent.get("BODY", "")
        )
        sTimestamp = aiSys.Timestamp()
        return sResult
    except Exception as e:
        return aiSys.ErrorProc(str(e), sProc)

def app_sendmail():
    """Ciclo principale di invio mail per SMTP."""
    global nCounter, dictMailCurrent, sTimestampStart, sTimestamp, nWaitStart, nWaitMail
    sProc = "app_sendmail"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""

    if nWaitStart > 0:
        print(f"{aiSys.Timestamp()} - Aspetto secondi in partenza: {nWaitStart}")
        time.sleep(nWaitStart)

    for sKey, mail_dict in dictMail.items():
        if sKey == "CONFIG":
            continue
        dictMailCurrent = mail_dict
        print(f"{aiSys.Timestamp()} - Invio mail: {sKey}")
        sResult = app_sendmail_smtp()
        if sResult != "":
            break
        nCounter += 1
        if nWaitMail > 0 and sKey != list(dictMail.keys())[-1]:
            print(f"{aiSys.Timestamp()} - Aspetto secondi: {nWaitMail}")
            time.sleep(nWaitMail)

    sTimestampEnd = aiSys.Timestamp()
    diff_seconds = aiSys.TimestampDiff(sTimestampStart, sTimestampEnd, "s")
    formatted_diff = format_timedelta(diff_seconds)
    print(f"Mail inviate: {nCounter}")
    print(f"Tempo totale: {formatted_diff}")

    return sResult

def cbCommands(dictJob):
    """Callback per jData.Run."""
    sCommand = jData.sCommand.upper()
    if sCommand == "SEND":
        return app_sendmail()
    else:
        return f"Comando non riconosciuto: {sCommand}"

def app_start(cbCommands):
    """Entry point dell'applicazione."""
    global jData, sMailEngine, sFileIn, sTimestampStart, nWaitStart, nWaitMail, sDelimiter, nMaxSizeAttach
    sProc = "app_start"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""

    jData = acJobsApp()
    sResult = jData.Start()
    if sResult != "":
        print(f"Errore Start: {sResult}")
        jData.End(sResult)
        return

    nCounter = 0
    nMaxSizeAttach = 5000000
    sTimestampStart = aiSys.Timestamp()

    nWaitStart = aiSys.StringToNum(jData.Config("WAIT.START"))
    if nWaitStart == 0:
        nWaitStart = 3

    nWaitMail = aiSys.StringToNum(jData.Config("WAIT.MAIL"))
    if nWaitMail == 0:
        nWaitMail = 5

    sDelimiter = jData.Config("DELIMITER")
    if sDelimiter == "":
        sDelimiter = ","

    sFileIn = jData.Config("FILE.IN.JSON")
    if not os.path.isfile(sFileIn):
        sResult = f"File json non presente: {sFileIn}"
        print(f"Errore Start: {sResult}")
        jData.End(sResult)
        return

    sResult = app_argread()
    if sResult == "":
        sResult = app_checkmail()

    sMailEngine = jData.Config("MAIL.ENGINE").upper()
    if sMailEngine != "SMTP":
        sResult = f"MAIL.ENGINE deve essere SMTP per questo script, trovato: {sMailEngine}"
        print(f"Errore Start: {sResult}")
        jData.End(sResult)
        return

    if sResult != "":
        print(f"Errore Start: {sResult}")
        jData.End(sResult)
        return

    # Avvio SMTP
    sResult = app_start_smtp()
    if sResult != "":
        print(f"Errore Start: {sResult}")
        jData.End(sResult)
        return

    sResult = jData.Run(cbCommands)
    jData.End(sResult)

# ----------------------------------------------------------------------
if __name__ == "__main__":
    app_start(cbCommands)