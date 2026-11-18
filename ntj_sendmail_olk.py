"""
ntj_sendmail_olk.py - Applicazione ntJobsApp per invio mail via Outlook
Orchestrata da acJobsApp, supporta aiSys e ncMailSimple (esterno)
"""
import sys
import os
import time
import json
import win32com.client
import aiSys
import acJobsApp

# =============================================================================
# HELPER LOCALI
# =============================================================================
def NormalizePath(sPath: str) -> str:
    """Normalizza i separatori di percorso in base al sistema operativo."""
    if not sPath: return ""
    if sys.platform == "win32":
        return os.path.normpath(sPath.replace("/", os.sep))
    else:
        return os.path.normpath(sPath.replace("\\", os.sep))

def format_timedelta(seconds):
    """Converte secondi in formato HH:MM:SS. Gestisce float e None."""
    if seconds is None:
        return "00:00:00"
    seconds = int(float(seconds))  # Forza conversione a intero
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"

# =============================================================================
# VARIABILI GLOBALI
# =============================================================================
jData = None
dictMail = {}
nCounter = 0
dictMailCurrent = {}
jMail = None
sMailId = ""
sMailEngine = ""
sFileIn = ""
sTimestamp = ""
sTimestampStart = ""
sDelimiter = ","
nWaitStart = 3
nWaitMail = 5
nMaxSizeAttach = 5000000
outlook_app = None
bSmtpStart = False  # Dichiarato per coerenza, usato solo in SMTP
bWaitStartDone = False

# =============================================================================
# FUNZIONI APPLICATIVE
# =============================================================================
def app_argread():
    global sFileIn, dictMail, sTimestamp, sDelimiter
    sProc = "app_argread"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    try:
        sFileInNorm = NormalizePath(sFileIn)
        if not os.path.exists(sFileInNorm):
            return aiSys.ErrorProc(f"File JSON non trovato: {sFileInNorm}", sProc)
        
        with open(sFileInNorm, 'r', encoding='utf-8') as f:
            raw_data = json.load(f)
            
        dictMail.clear()
        for key, value in raw_data.items():
            new_key = str(key).strip().upper()
            if isinstance(value, dict):
                new_value = {}
                for k, v in value.items():
                    new_k = str(k).strip().upper()
                    if isinstance(v, str):
                        v = v.strip().strip('"')
                        if new_k in ["TO", "CC", "CCN", "ATTACH"]:
                            v = [x.strip() for x in v.split(sDelimiter) if x.strip()]
                    new_value[new_k] = v
                dictMail[new_key] = new_value
            else:
                dictMail[new_key] = value
                
        sTimestamp = aiSys.Timestamp()
    except Exception as e:
        sResult = aiSys.ErrorProc(str(e), sProc)
    return sResult

def app_checkmail():
    global dictMail, nMaxSizeAttach
    sProc = "app_checkmail"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    try:
        for sKey, dictTemp in dictMail.items():
            if sKey == "CONFIG" or not isinstance(dictTemp, dict):
                continue
                
            if not dictTemp.get("TO"):
                return aiSys.ErrorProc(f"Campo TO obbligatorio mancante - ID: {sKey}", sProc)
                
            for sField in ["TO", "CC", "CCN"]:
                for email in dictTemp.get(sField, []):
                    if not aiSys.isEmail(email):
                        return aiSys.ErrorProc(f"Email non valida nel campo {sField}: {email} - ID: {sKey}", sProc)
                        
            for sFile in dictTemp.get("ATTACH", []):
                sFileNorm = NormalizePath(sFile)
                if not os.path.isfile(sFileNorm):
                    return aiSys.ErrorProc(f"File attach non presente: {sFileNorm} - ID: {sKey}", sProc)
                if os.path.getsize(sFileNorm) > nMaxSizeAttach:
                    return aiSys.ErrorProc(f"File supera dimensione massima ({nMaxSizeAttach}): {sFileNorm} - ID: {sKey}", sProc)
                    
            dictMail[sKey] = dictTemp.copy()
    except Exception as e:
        sResult = aiSys.ErrorProc(str(e), sProc)
    return sResult

def app_olk_start():
    global outlook_app
    sProc = "app_olk_start"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    try:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        sResult = aiSys.ErrorProc(str(e), sProc)
    return sResult

def app_olk_kill():
    global outlook_app
    sProc = "app_olk_kill"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    try:
        if outlook_app:
            outlook_app.Quit()
            outlook_app = None
    except Exception as e:
        sResult = aiSys.ErrorProc(str(e), sProc)
    return sResult

def app_sendmail_olk():
    global dictMailCurrent, sTimestamp
    sProc = "app_sendmail_olk"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    try:
        if not outlook_app:
            return aiSys.ErrorProc("Motore Outlook non avviato", sProc)

        mail = outlook_app.CreateItem(0)
        mail.To = "; ".join(dictMailCurrent.get("TO", []))
        if dictMailCurrent.get("CC"):
            mail.CC = "; ".join(dictMailCurrent["CC"])
        if dictMailCurrent.get("CCN"):
            mail.BCC = "; ".join(dictMailCurrent["CCN"])
            
        mail.Subject = dictMailCurrent.get("SUBJECT", "")
        
        sFormat = str(dictMailCurrent.get("FORMAT", "TXT")).upper()
        sBody = dictMailCurrent.get("BODY", "")
        if sFormat == "HTML":
            mail.HTMLBody = sBody
        else:
            mail.Body = sBody
            
        for att in dictMailCurrent.get("ATTACH", []):
            mail.Attachments.Add(NormalizePath(att))
            
        mail.Send()
        sTimestamp = aiSys.Timestamp()
    except Exception as e:
        sResult = aiSys.ErrorProc(str(e), sProc)
    return sResult

def app_sendmail():
    global dictMail, sMailEngine, sMailId, dictMailCurrent, nCounter, nWaitStart, nWaitMail, sTimestampStart, sTimestamp, bWaitStartDone
    sProc = "app_sendmail"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    
    if nWaitStart > 0 and not bWaitStartDone:
        print(f"{aiSys.Timestamp()} Aspetto secondi in partenza: {nWaitStart}")
        time.sleep(nWaitStart)
        bWaitStartDone = True
        
    for sKey, dictMailCurrent in dictMail.items():
        if sKey == "CONFIG" or not isinstance(dictMailCurrent, dict):
            continue
            
        sMailId = sKey
        print(f"{aiSys.Timestamp()} Invio mail {sMailId}")
        
        if sMailEngine == "OLK":
            sResult = app_sendmail_olk()
        else:
            sResult = aiSys.ErrorProc(f"Motore mail non supportato: {sMailEngine}", sProc)
            
        if sResult:
            break
            
        if sResult == "":
            nCounter += 1
            
        if nWaitMail > 0:
            print(f"{aiSys.Timestamp()} Aspetto secondi: {nWaitMail}")
            time.sleep(nWaitMail)
            
    sTimestampEnd = aiSys.Timestamp()
    diff_seconds = aiSys.TimestampDiff(sTimestampStart, sTimestampEnd, "s")
    formatted_diff = format_timedelta(diff_seconds)
    print(f"{aiSys.Timestamp()} Mail inviate: {nCounter} | Tempo totale: {formatted_diff}")
    return sResult

def cbCommand(dictJob):
    global sResult
    sResult = ""
    sCommand = dictJob.get("COMMAND", jData.sCommand)
    if sCommand == "SEND":
        sResult = app_sendmail()
    elif sCommand == "OLK.START":
        sResult = app_olk_start()
    elif sCommand == "OLK.KILL":
        sResult = app_olk_kill()
    else:
        sResult = aiSys.ErrorProc(f"Comando non riconosciuto: {sCommand}", "cbCommand")
    return sResult

def app_start():
    global jData, nCounter, nMaxSizeAttach, sTimestampStart, nWaitStart, nWaitMail, sDelimiter, sFileIn, sMailEngine, bWaitStartDone
    sProc = "app_start"
    print(f"{aiSys.Timestamp()} - Esecuzione {sProc}")
    sResult = ""
    try:
        jData = acJobsApp.acJobsApp()
        sResult = jData.Start()
        if sResult:
            print(f"Errore Start: {sResult}")
            jData.End(sResult)
            return sResult

        nCounter = 0
        nMaxSizeAttach = 5000000
        sTimestampStart = aiSys.Timestamp()
        bWaitStartDone = False
        
        nWaitStart = aiSys.StringToNum(jData.Config("WAIT.START"))
        if nWaitStart == 0: nWaitStart = 3
        nWaitMail = aiSys.StringToNum(jData.Config("WAIT.MAIL"))
        if nWaitMail == 0: nWaitMail = 5
        
        sDelimiter = jData.Config("DELIMITER")
        if not sDelimiter: sDelimiter = ","
        
        sFileIn = jData.Config("FILE.IN.JSON")
        if not sFileIn:
            sResult = aiSys.ErrorProc("FILE.IN.JSON non specificato nel CONFIG", sProc)
        else:
            sResult = app_argread()
            if not sResult:
                sResult = app_checkmail()
                
        if sResult:
            print(f"Errore Start: {sResult}")
            jData.End(sResult)
            return sResult

        sMailEngine = jData.Config("MAIL.ENGINE").upper()
        if sMailEngine != "OLK":
            sResult = aiSys.ErrorProc(f"Motore {sMailEngine} non supportato in questo script", sProc)
            jData.End(sResult)
            return sResult

        sResult = jData.Run(cbCommand)
        if sResult:
            print(f"Errore Run: {sResult}")
        jData.End(sResult)
    except Exception as e:
        sResult = aiSys.ErrorProc(str(e), sProc)
        if jData: jData.End(sResult)
    return sResult

# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    app_start()