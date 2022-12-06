import pandas as pd
from openpyxl.reader.excel import load_workbook
import spacy


def mtm_code(decription):
    aa_words = ['get', 'place', 'approx']
    ab_words = ["get", "place", "exact"]
    ac_words = ["get", "place", "approx", "bulky"]
    ad_words = ["get", "place", "exact", "bulky"]
    pa_words = ["place", "approx"]
    pb_words = ["place", "exact"]
    ha_words = ["handle", "approx"]
    hb_words = ["handle"]
    ba_words = ["operate"]
    bb_words = ["operate", "compound"]
    za_words = ["motion", "without shift", "<=10"]
    zb_words = ["motion", "without shift", ">10 - 30"]
    zc_words = ["motion", "without shift", ">30 - 80"]
    zd_words = ["motion", "with shift", "<=20"]
    ze_words = ["motion", "with shift", "> 20 - 45"]
    zf_words = ["motion", "with shift", "> 45 - 100"]
    zz_words = ["tighten"]
    zz1_words = ["loosen"]
    ka_words = ["walk"]
    kb_words = ["bend"]
    kc_words = ["sit"]
    kc1_words = ["stand"]
    va_words = ["visual"]

    if set(aa_words).issubset(decription):
        return 'AA'
    if set(ab_words).issubset(decription):
        return 'AB'
    if set(ac_words).issubset(decription):
        return 'AC'
    if set(ad_words).issubset(decription):
        return 'AD'
    if set(pa_words).issubset(decription):
        return 'PA'
    if set(pb_words).issubset(decription):
        return 'PB'
    if set(ha_words).issubset(decription):
        return 'HA'
    if set(hb_words).issubset(decription):
        return 'HB'
    if set(ba_words).issubset(decription):
        return 'BA'
    if set(bb_words).issubset(decription):
        return 'BB'
    if set(za_words).issubset(decription):
        return 'ZA'
    if set(zb_words).issubset(decription):
        return 'ZB'
    if set(zc_words).issubset(decription):
        return 'ZC'
    if set(zd_words).issubset(decription):
        return 'ZD'
    if set(ze_words).issubset(decription):
        return 'ZE'
    if set(zf_words).issubset(decription):
        return 'ZF'
    if set(zz_words).issubset(decription):
        return 'ZZ'
    if set(zz1_words).issubset(decription):
        return 'ZZ'
    if set(ka_words).issubset(decription):
        return 'KA'
    if set(kb_words).issubset(decription):
        return 'KB'
    if set(kc_words).issubset(decription):
        return 'KC'
    if set(kc1_words).issubset(decription):
        return 'KC'
    if set(va_words).issubset(decription):
        return 'VA'


code_df = []

wb = load_workbook(filename='MTM_MEK_Template.xlsm', read_only=False, keep_vba=True)
ws = wb["SOS_ERGO - AS-IS"]
df = pd.DataFrame(ws.values)


for row in df[9].values[1:]:
    code_df.append(row)

print(code_df)
value = ""
value1 = []
nlp = spacy.load("en_core_web_sm")
doc = nlp(code_df)
sentence_lemma = []

for token in doc:
    sentence_lemma.append(token.lemma_)

value = mtm_code(sentence_lemma)
value1.append(value)




print(value1)

