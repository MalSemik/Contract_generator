from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from num2words import num2words

name = ""
hiring = "1"
address1 = ""
address2 = ""
remuneration = 41000
gender = "F"
categoria = "I7"
mansione = "R&D Engineer"
job_req = "T7"
premio_feriale = 1750
PDR = "specifico"
sede = "Via Friuli, 4 – 24044 Dalmine (BG)"
probation = 6
intellectual_property = True
pdr_amount = 5767

def fill_template(name, hiring, address1, address2, remuneration, gender, categoria, mansione, job_req, premio_feriale,
                  PDR, sede, probation, intellectual_property, pdr_amount):
    template = "ABB_iT_LETTERA INDETERMINATO_impiegato.docx"
    document = MailMerge(template)
    print(document.get_merge_fields())
    # {'Date', 'Address1', 'hiring', 'Klauzula3', 'Name', 'Address2', 'Klauzula1', 'remuneration', 'salutation', 'Klauzula2'}
    remuneration = int(remuneration)
    remuneration_word = num2words(remuneration, lang='it') + "/00"
    premio_feriale = int(premio_feriale)
    premio_feriale_word = num2words(premio_feriale, lang="it") + "/00"
    probation = int(probation)
    probation_word = num2words(probation, lang="it")
    pdr_word = num2words(pdr_amount, lang='it')
    if categoria[0] == "I":
        qualifica = "Impiegat"
    elif categoria[0] == "Q":
        qualifica = "Quadro"

    categoria_num = categoria[1] + "a"

    if gender == 'M':
        salutation = "Gentile Sig."
        desinenza = "o"
        if categoria[0] == "I":
            qualifica += desinenza
    else:
        salutation = "Gentile Sig.ra"
        desinenza = "a"
        if categoria[0] == "I":
            qualifica += desinenza

    if PDR == "generale":
        PDR = "Le verrà inoltre riconosciuto secondo le modalità previste dagli accordi aziendali vigenti il premio di risultato variabile.\n"
    elif PDR == "specifico":
        PDR = f"Le verrà riconosciuto secondo le modalità previste dagli accordi aziendali vigenti il premio di risultato variabile che, per l’anno 2021, è pari a Euro {pdr_amount} ({pdr_word}/00) come importo massimo lordo erogabile."
    if intellectual_property:
        intellectual_property = """La informiamo inoltre che la nostra Società si riserva espressamente la facoltà di adibirLa alla ricerca di nuovi ritrovati, anche di carattere inventivo o innovativo, che possano eventualmente essere tutelati con brevetti o altri diritti di proprietà industriale o intellettuale, interessanti la propria sfera di attività e quelle delle altre aziende direttamente controllate o facenti parte del Gruppo. Eventuali invenzioni da Lei realizzate sono pertanto da intendersi di servizio e l’azienda se ne riserva l’eventuale brevettazione ai sensi dell’art. 23 comma 1 del R.D. 1127/39.

Conseguentemente, nei termini previsti dalle norme di legge, i diritti derivanti da tali ritrovati anche di carattere inventivo, frutto in tutto o in parte della Sua attività esplicata alla nostre dipendenze, saranno della Società, la quale potrà sfruttarne a proprio beneficio le applicazioni ed i brevetti.

In ogni caso dovrà osservare scrupolosamente, sia all'interno che all'esterno dell'azienda, le norme che regolano la materia ed il segreto di ufficio che restano vincolanti, anche in caso di risoluzione del rapporto di lavoro secondo quanto previsto dagli artt. 622 (Rivelazione di segreto professionale), 623 (Rivelazione di segreti scientifici o industriali) e 623 bis (altre comunicazioni e conversazioni) del Codice Penale, art. 6 bis R.D. 1127/39 (atti di concorrenza sleale).

Resta inteso che il trattamento economico sopra indicato è comprensivo di ogni e qualsiasi compenso, sia per gli eventuali frutti delle Sue ricerche, inclusi brevetti e invenzioni, che per la Sua normale attività lavorativa.
\n"""
    document.merge(
        Date='{:%d.%m.%Y}'.format(date.today()),
        salutation=salutation,
        Address1=address1,
        Address2=address2,
        hiring=hiring,
        name=name,
        remuneration=str(remuneration),
        remuneration_word=remuneration_word,
        qualifica=qualifica,
        desinenza=desinenza,
        categoria_num=categoria_num,
        mansione=mansione,
        job_req=job_req,
        premio_feriale=str(premio_feriale),
        premio_feriale_word=premio_feriale_word,
        PDR=PDR,
        probation=str(probation),
        probation_word=probation_word,
        intellectual_property=intellectual_property,
        sede= sede,
    )
    document.write(f'01 {name} - lettera.docx')


fill_template(name, hiring, address1, address2, remuneration, gender, categoria, mansione, job_req, premio_feriale,
                  PDR, sede, probation, intellectual_property, pdr_amount)
