#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Génération des lettres de publipostage BELVEDERE.

Les lettres ne sont plus générées en masse : une lettre est produite à la
demande (clic dans l'interface, ou génération automatique si l'option est
cochée). Le résultat est mis en cache dans Letters/<id_nom_cin>/ et n'est
régénéré que si le contenu du client a changé (comparaison d'empreinte).
"""

import os
import re
import json

try:
    import docx
except ImportError:
    docx = None

HERE = os.path.dirname(os.path.abspath(__file__))

try:
    from paths import resource
except Exception:
    def resource(*parts):
        return os.path.join(HERE, *parts)

MERGE_MAP = {
    "NomPrenom":   lambda c: c.get("nom_prenom", ""),
    "Adresse1":    lambda c: c.get("adresse1", ""),
    "Adresse2":    lambda c: c.get("adresse2", ""),
    "Ville":       lambda c: c.get("ville", ""),
    "Pays":        lambda c: c.get("pays", ""),
    "DateCourrier": lambda c: c.get("date_courrier", ""),
    "TypeBien":    lambda c: c.get("type_bien", ""),
    "Ilot":        lambda c: c.get("ilot", ""),
    "Immeuble":    lambda c: c.get("immeuble", ""),
    "NumeroBien":  lambda c: c.get("numero", ""),
}

FIELD_RE = re.compile(r"[«<]{1,2}\s*([A-Za-z0-9_]+)\s*[»>]{1,2}")


def _safe(s):
    s = re.sub(r"[^\w\s-]", "", s or "", flags=re.U).strip()
    return re.sub(r"\s+", "_", s)[:50]


def _norm_cmp(s):
    """Minuscule sans accents/ponctuation, pour comparer deux libellés."""
    import unicodedata
    s = unicodedata.normalize("NFKD", (s or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def address_lines(c):
    """Construit l'encadré adresse SANS doublon ville / pays.

    - « Neuss (Allemagne) » -> ville = « Neuss », pays = « Allemagne ».
    - Si la ville est déjà présente à la fin de l'adresse, on ne la répète pas.
    - Pays jamais répété s'il figure déjà dans la ville/adresse.
    """
    a1 = (c.get("adresse1", "") or "").strip()
    a2 = (c.get("adresse2", "") or "").strip()
    ville = (c.get("ville", "") or "").strip()
    pays = (c.get("pays", "") or "").strip()

    # « Ville (Pays) » -> on isole le pays et on garde la ville pure.
    m = re.match(r"^(.*?)\s*\(([^)]+)\)\s*$", ville)
    if m:
        ville_pure = m.group(1).strip()
        paren = m.group(2).strip()
        if ville_pure:
            ville = ville_pure
        if not pays and paren:
            pays = paren
        # si le pays diffère de la parenthèse, on fait confiance au champ pays.

    lines = [x for x in (a1, a2) if x]
    joined = _norm_cmp(" ".join(lines))

    # N'ajouter la ville que si elle n'est pas déjà dans les lignes d'adresse.
    nv = _norm_cmp(ville)
    if ville and nv and nv not in joined:
        lines.append(ville)
        joined = _norm_cmp(" ".join(lines))

    # N'ajouter le pays que s'il n'est pas déjà présent.
    npays = _norm_cmp(pays)
    if pays and npays and npays not in joined:
        lines.append(pays)

    return lines


def folder_name(c):
    nom = _safe(c.get("nom_prenom", "").split(" / ")[0]) or "CLIENT"
    cin = _safe((c.get("cin_str") or "").split(" / ")[0])
    return f"{c['id']:03d}_{nom}" + (f"_{cin}" if cin else "")


def find_template(root):
    for f in os.listdir(root):
        if "lettre" in f.lower() and f.lower().endswith(".docx"):
            return os.path.join(root, f)
    return None


# --------------------------------------------------------------------------- #
#  Remplissage du .docx (champs pouvant être découpés en plusieurs runs)      #
# --------------------------------------------------------------------------- #

def _fill_paragraph(par, values):
    full = "".join(r.text for r in par.runs)
    if "«" not in full and "<<" not in full:
        return
    new = FIELD_RE.sub(lambda m: str(values.get(m.group(1), m.group(0))), full)
    if new != full and par.runs:
        par.runs[0].text = new
        for r in par.runs[1:]:
            r.text = ""


def fill_docx(template, client, out_path):
    if not docx or not template or not os.path.exists(template):
        return False
    values = {k: fn(client) for k, fn in MERGE_MAP.items()}
    d = docx.Document(template)
    for p in d.paragraphs:
        _fill_paragraph(p, values)
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _fill_paragraph(p, values)
    d.save(out_path)
    return True


# --------------------------------------------------------------------------- #
#  Version HTML (aperçu web + impression PDF)                                 #
# --------------------------------------------------------------------------- #

BODY = """\
<p>Objet : <strong>Disponibilité du titre foncier – Résidence Belvédère</strong></p>
<p>Madame, Monsieur,</p>
<p>Nous vous informons que le titre foncier du bien que vous avez réservé à la
Résidence BELVÉDÈRE, projet de la société CORAIL DEVELOPPEMENT, est désormais
disponible.</p>
<p>Par conséquent, nous vous invitons à vous présenter, dans un délai de deux
semaines à compter de la réception de la présente lettre, à notre bureau de vente
situé avenue Lalla Meryem à Sala Al Jadida, afin de régulariser la situation de
votre dossier et de finaliser les démarches administratives nécessaires. Vous serez
ensuite orienté vers le notaire chargé d'établir votre contrat d'acquisition.</p>
<p>Nous vous prions de bien vouloir accomplir ces démarches dans le délai indiqué
afin d'éviter tout désagrément lié à un éventuel retard.</p>
<p>Nous vous remercions de votre confiance et restons à votre disposition pour tout
complément d'information aux numéros suivants : 06 66 88 66 20 ou 06 66 33 09 13.</p>
<p>Veuillez agréer, Madame, Monsieur, l'expression de nos salutations distinguées.</p>
<p class="sign">La Direction Commerciale</p>
"""

# Mêmes paragraphes en texte brut (pour le rendu PDF fpdf2).
_BODY_PARAS = [
    "Objet : Disponibilité du titre foncier – Résidence Belvédère",
    "Madame, Monsieur,",
    "Nous vous informons que le titre foncier du bien que vous avez réservé à la "
    "Résidence BELVÉDÈRE, projet de la société CORAIL DEVELOPPEMENT, est désormais disponible.",
    "Par conséquent, nous vous invitons à vous présenter, dans un délai de deux "
    "semaines à compter de la réception de la présente lettre, à notre bureau de vente "
    "situé avenue Lalla Meryem à Sala Al Jadida, afin de régulariser la situation de "
    "votre dossier et de finaliser les démarches administratives nécessaires. Vous serez "
    "ensuite orienté vers le notaire chargé d'établir votre contrat d'acquisition.",
    "Nous vous prions de bien vouloir accomplir ces démarches dans le délai indiqué "
    "afin d'éviter tout désagrément lié à un éventuel retard.",
    "Nous vous remercions de votre confiance et restons à votre disposition pour tout "
    "complément d'information aux numéros suivants : 06 66 88 66 20 ou 06 66 33 09 13.",
    "Veuillez agréer, Madame, Monsieur, l'expression de nos salutations distinguées.",
]


def build_html(client, logo_data_uri="", company=None, date_override="",
               show_print=True):
    """Construit le HTML de la lettre (chaîne) — sert au fichier, à l'aperçu et au PDF."""
    company = company or {}
    v = {k: fn(client) for k, fn in MERGE_MAP.items()}
    date_txt = date_override or v["DateCourrier"] or "........................"
    addr = "<br>".join(x for x in (v["Adresse1"], v["Adresse2"], v["Ville"], v["Pays"]) if x)
    ref = " &nbsp;·&nbsp; ".join(x for x in (
        f"Réf. : {v['TypeBien']}" if v["TypeBien"] else "",
        f"Îlot {v['Ilot']}" if v["Ilot"] else "",
        f"Immeuble {v['Immeuble']}" if v["Immeuble"] else "",
        f"N° {v['NumeroBien']}" if v["NumeroBien"] else "") if x)
    logo = f'<img src="{logo_data_uri}" style="height:70px">' if logo_data_uri else ""

    cname = (company.get("name") or "").strip()
    header = f'<div class="coname">{cname}</div>' if cname else ""
    footer_bits = [company.get("address", ""), company.get("phone", ""),
                   company.get("email", ""), company.get("extra", "")]
    footer = " &nbsp;·&nbsp; ".join(x for x in footer_bits if x and x.strip())
    footer_html = f'<div class="foot">{footer}</div>' if footer else ""

    print_btn = ('<div class="noprint"><button onclick="window.print()">'
                 'Imprimer / Enregistrer en PDF</button></div>') if show_print else ""

    return f"""<!doctype html><html lang="fr"><head><meta charset="utf-8">
<title>Lettre – {v['NomPrenom']}</title><style>
 @page {{ size:A4; margin:2.2cm; }}
 body {{ font-family:"Times New Roman",Georgia,serif; font-size:13pt; line-height:1.55;
        color:#111; max-width:18cm; margin:1cm auto; padding:1.4cm; background:#fff; }}
 .top {{ display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:1cm; }}
 .brand {{ display:flex; flex-direction:column; gap:4px; }}
 .coname {{ font-weight:bold; color:#143A82; font-size:12pt; }}
 .addr {{ width:9cm; border:1px solid #ccc; padding:.5cm .7cm; border-radius:4px; align-self:flex-end; }}
 .date {{ text-align:right; margin:.6cm 0 1cm; }}
 .ref {{ font-weight:bold; margin-bottom:1cm; }}
 p {{ text-align:justify; margin:.5cm 0; }}
 .sign {{ margin-top:1.4cm; font-weight:bold; text-align:right; }}
 .foot {{ margin-top:1.6cm; padding-top:.4cm; border-top:1px solid #ccc; font-size:9.5pt;
          color:#555; text-align:center; }}
 .noprint {{ text-align:center; margin-bottom:1cm; }}
 .noprint button {{ padding:.55em 1.3em; font-size:12pt; cursor:pointer; background:#143A82;
        color:#fff; border:0; border-radius:6px; font-weight:600; }}
 @media print {{ body{{margin:0;padding:0;max-width:none}} .noprint{{display:none}} }}
</style></head><body>
{print_btn}
<div class="top"><div class="brand">{logo}{header}</div><div class="addr">{addr or '&nbsp;'}</div></div>
<div class="date">Rabat, le {date_txt}</div>
<div class="ref">{ref}</div>
{BODY}{footer_html}</body></html>"""


def fill_html(client, out_path, logo_data_uri="", company=None, date_override=""):
    html = build_html(client, logo_data_uri, company, date_override)
    open(out_path, "w", encoding="utf-8").write(html)
    return True


# --------------------------------------------------------------------------- #
#  Cache                                                                        #
# --------------------------------------------------------------------------- #

def _meta_path(fol):
    return os.path.join(fol, ".meta.json")


def is_cached(client, fol):
    """La lettre en cache correspond-elle encore au contenu actuel du client ?"""
    html = os.path.join(fol, "lettre.html")
    pdf = os.path.join(fol, "lettre.pdf")
    meta = _meta_path(fol)
    # On exige aussi le PDF : une lettre en cache sans PDF est régénérée.
    if not (os.path.exists(html) and os.path.exists(pdf) and os.path.exists(meta)):
        return False
    try:
        m = json.load(open(meta, encoding="utf-8"))
    except Exception:
        return False
    return m.get("hash") == client.get("hash")


def ensure_letter(client, root, template=None, logo_uri="", force=False, company=None):
    """Génère (ou réutilise le cache) la lettre d'un client. Retourne les chemins."""
    letters_dir = os.path.join(root, "Letters")
    os.makedirs(letters_dir, exist_ok=True)
    fol = os.path.join(letters_dir, folder_name(client))

    if not force and is_cached(client, fol):
        return _paths(client, fol, root, cached=True)

    os.makedirs(fol, exist_ok=True)
    if template is None:
        template = find_template(root)
    html_p = os.path.join(fol, "lettre.html")
    fill_html(client, html_p, logo_uri, company)
    fill_docx(template, client, os.path.join(fol, "lettre.docx"))
    # PDF fidèle : LibreOffice/wkhtmltopdf si présents, sinon rendu fpdf2 intégré.
    pdf_p = os.path.join(fol, "lettre.pdf")
    try:
        if _find_soffice() or __import__("shutil").which("wkhtmltopdf"):
            html_to_pdf(html_p, pdf_p, client=client, logo_uri=logo_uri, company=company)
        else:
            _pdf_via_fpdf(client, pdf_p, logo_uri, company)
    except Exception:
        pass  # le PDF est un bonus ; HTML + DOCX restent disponibles
    json.dump({"hash": client.get("hash"), "nom": client.get("nom_prenom")},
              open(_meta_path(fol), "w", encoding="utf-8"), ensure_ascii=False)
    return _paths(client, fol, root, cached=False)


def _paths(client, fol, root, cached):
    html = os.path.join(fol, "lettre.html")
    docx_p = os.path.join(fol, "lettre.docx")
    pdf_p = os.path.join(fol, "lettre.pdf")
    rel = lambda p: os.path.relpath(p, root) if os.path.exists(p) else ""
    return {
        "letter_dir": os.path.relpath(fol, root),
        "letter_html": rel(html),
        "letter_docx": rel(docx_p),
        "letter_pdf": rel(pdf_p),
        "cached": cached,
    }


def logo_data_uri(custom=""):
    """Logo à utiliser : personnalisé (data URI) sinon logo Romana par défaut."""
    if custom and custom.startswith("data:image"):
        return custom
    import paths
    p = paths.resource("static", "logo.png")
    if not os.path.exists(p):
        return ""
    import base64
    return "data:image/png;base64," + base64.b64encode(open(p, "rb").read()).decode()


# --------------------------------------------------------------------------- #
#  Génération en lot : fichiers plats dans un dossier choisi                   #
# --------------------------------------------------------------------------- #

def _find_soffice():
    import shutil
    for c in ("soffice", "libreoffice",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if os.path.exists(c) or shutil.which(c):
            return c
    return None


def output_filename(client, pattern=""):
    """Nom de fichier plat pour un client (sans extension)."""
    nom = _safe(client.get("nom_prenom", "").split(" / ")[0]) or "CLIENT"
    cin = _safe((client.get("cin_str") or "").split(" / ")[0])
    ref = "-".join(x for x in (client.get("ilot", ""), client.get("immeuble", ""),
                               str(client.get("numero", ""))) if x)
    if pattern:
        name = pattern
        for k, val in (("{nom}", nom), ("{cin}", cin), ("{ilot}", client.get("ilot", "")),
                       ("{imm}", client.get("immeuble", "")), ("{no}", str(client.get("numero", ""))),
                       ("{ref}", ref), ("{type}", client.get("type_bien", ""))):
            name = name.replace(k, str(val))
        return _safe(name) or f"Lettre_{nom}"
    parts = ["Lettre", nom] + ([cin] if cin else []) + ([ref] if ref else [])
    return "_".join(parts)


def _unicode_font_candidates():
    """Chemins possibles d'une police TTF Unicode (regular, bold)."""
    import os as _os
    out = []
    # 1) Police fournie avec l'application (hors-ligne, portable).
    try:
        base = resource("static")
        for reg, bold in (("DejaVuSans.ttf", "DejaVuSans-Bold.ttf"),
                          ("NotoSans-Regular.ttf", "NotoSans-Bold.ttf")):
            p = _os.path.join(base, reg)
            if _os.path.exists(p):
                out.append((p, _os.path.join(base, bold)))
    except Exception:
        pass
    # 2) Polices système (macOS puis Windows).
    sys_fonts = [
        ("/Library/Fonts/Arial Unicode.ttf", "/Library/Fonts/Arial Unicode.ttf"),
        ("/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
         "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
        ("/System/Library/Fonts/Supplemental/Arial.ttf",
         "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
        ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("C:/Windows/Fonts/tahoma.ttf", "C:/Windows/Fonts/tahomabd.ttf"),
        ("C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/segoeuib.ttf"),
    ]
    out.extend(sys_fonts)
    return out


def _register_unicode_font(pdf):
    """Enregistre une police Unicode sous le nom 'Uni'. Renvoie 'Uni' ou 'Times'."""
    import os as _os
    for reg, bold in _unicode_font_candidates():
        if not (reg and _os.path.exists(reg)):
            continue
        try:
            pdf.add_font("Uni", "", reg)
            if bold and _os.path.exists(bold):
                pdf.add_font("Uni", "B", bold)
            else:
                pdf.add_font("Uni", "B", reg)  # faute de gras : régulier
            return "Uni"
        except Exception:
            continue
    return "Times"


def _draw_letter_page(pdf, client, company=None, date_override=""):
    """Dessine UNE lettre (une page) dans un objet FPDF existant.

    Utilisé à la fois pour une lettre seule et pour le PDF combiné (toutes les
    lettres dans un seul document, une par page). Reproduit fidèlement le
    modèle : logo, encadré adresse (nom + adresse), date, réf., corps, pied.
    """
    company = company or {}
    v = {k: fn(client) for k, fn in MERGE_MAP.items()}
    date_txt = date_override or v["DateCourrier"] or "........................"

    pdf.set_auto_page_break(auto=True, margin=22)
    # Marges gauche/droite un peu plus serrées : le corps (et les numéros de
    # téléphone) tient mieux sur une ligne.
    pdf.set_margins(18, 20, 18)
    pdf.add_page()

    # Police Unicode pour accents + ponctuation typographique.
    # fpdf2 récent ne fournit plus DejaVu : on cherche d'abord une police
    # fournie avec l'app (static/), puis les polices système, sinon Times.
    FONT = _register_unicode_font(pdf)

    def _ascii_punct(s):
        # Remplace la ponctuation hors latin-1 par des équivalents sûrs
        # (utilisé seulement si aucune police Unicode n'est disponible).
        return (s.replace("–", "-").replace("—", "-")
                 .replace("‘", "'").replace("’", "'")
                 .replace("“", '"').replace("”", '"')
                 .replace("…", "...").replace(" ", " ")
                 .replace(" ", " ").replace("•", "-"))

    def txt(s):
        if FONT == "Uni":
            return s
        return _ascii_punct(s).encode("latin-1", "replace").decode("latin-1")

    # En-tête = logo CORAIL DEVELOPPEMENT (promoteur du projet), comme le modèle.
    # (Romana est l'agence ; elle figure uniquement en pied de page.)
    import tempfile, base64
    corail = ""
    try:
        import paths
        p = paths.resource("static", "corail_logo.png")
        if os.path.exists(p):
            corail = p
    except Exception:
        pass
    if corail:
        try:
            # Logo centré horizontalement en haut (A4 = 210 mm de large).
            logo_w = 58
            pdf.image(corail, x=(210 - logo_w) / 2, y=10, w=logo_w)
        except Exception:
            pass
    pdf.set_text_color(17, 17, 17)

    # Encadré adresse (à droite) — logo plus haut, encadré nettement plus bas.
    # 1re ligne = NOM du destinataire (en gras), puis l'adresse.
    nom = (v.get("NomPrenom") or "").strip()
    addr_lines = address_lines(client)
    pdf.set_draw_color(200, 200, 200)
    box_y = 48            # un peu plus bas : espace sous le logo
    box_x, box_w = 110, 85
    # Nom en gras sur la 1re ligne, le reste en normal, dans le même encadré.
    body = (("**" + txt(nom) + "**\n") if nom else "") + txt("\n".join(addr_lines) or " ")
    pdf.set_xy(box_x, box_y)
    pdf.set_font(FONT, "", 11)
    pdf.multi_cell(box_w, 6, body, border=1, align="L", markdown=True)

    pdf.set_y(max(pdf.get_y(), 76))
    pdf.set_font(FONT, "", 11)
    pdf.cell(0, 7, txt("Rabat, le " + date_txt), ln=1, align="R")
    pdf.ln(4)

    # Bloc Référence — sur UNE seule ligne :
    #   Réf. : <TypeBien> — Îlot <Ilot> — Immeuble <Immeuble> — N° <NumeroBien>
    ref_parts = [
        v["TypeBien"] if v["TypeBien"] else "",
        f"Îlot {v['Ilot']}" if v["Ilot"] else "",
        f"Immeuble {v['Immeuble']}" if v["Immeuble"] else "",
        f"N° {v['NumeroBien']}" if v["NumeroBien"] else "",
    ]
    ref_parts = [p for p in ref_parts if p]
    ref_line = "Réf. : " + " – ".join(ref_parts) if ref_parts else "Réf. :"
    pdf.set_font(FONT, "B", 11)
    pdf.multi_cell(0, 6, txt(ref_line))
    pdf.ln(3)

    # Corps — certaines portions sont mises en forme (Objet en gras, phrase
    # soulignée) via le mode markdown de fpdf2 : **gras**, --souligné--.
    def nbsp_phone(s):
        # Colle les 4 derniers chiffres de chaque n° de tél. pour éviter
        # qu'un « 13 » isolé ne saute à la ligne. 06 66 33 09 13 -> 06 66 33 0913
        # Colle UNIQUEMENT le dernier groupe « NN NN » (les 4 derniers
        # chiffres) de chaque numéro, p.ex. « 66 20 » -> « 6620 ».
        # Le lookahead (?!\s*\d) garantit qu'on ne touche que la paire finale.
        return re.sub(r"(\d\d)\s(\d\d)(?!\s*\d)", r"\1\2", s)

    pdf.set_font(FONT, "", 11)
    # Interligne ~1,3 : hauteur de ligne ~1,3× le corps (11 pt ≈ 5,5 mm).
    LH = 7.2
    for para in _BODY_PARAS:
        gap = 5            # espace sous le paragraphe (par défaut)
        if para.startswith("Objet :"):
            # Toute la ligne « Objet : … » en gras (mot « Objet » compris).
            pdf.set_font(FONT, "B", 11)
            pdf.multi_cell(0, LH, txt(para), align="L")
            pdf.set_font(FONT, "", 11)
            # Ligne vide supplémentaire après l'objet (avant « Madame, Monsieur »).
            pdf.ln(LH)
        elif para.startswith("Par conséquent"):
            UNDER = ("dans un délai de deux semaines à compter de la réception "
                     "de la présente lettre")
            marked = txt(para).replace(txt(UNDER), "--" + txt(UNDER) + "--")
            pdf.multi_cell(0, LH, marked, align="J", markdown=True)
        elif "numéros suivants" in para:
            pdf.multi_cell(0, LH, txt(nbsp_phone(para)), align="J")
        else:
            pdf.multi_cell(0, LH, txt(para), align="J")
            # Espace réduit juste après « Madame, Monsieur, ».
            if para.startswith("Madame, Monsieur"):
                gap = 2
        # Ligne vide entre les paragraphes (aération, comme le modèle).
        pdf.ln(gap)

    pdf.ln(8)
    pdf.set_font(FONT, "B", 11)
    pdf.cell(0, 6, txt("La Direction Commerciale"), ln=1, align="R")

    # Pied de page ROMANA (agence) — reproduit EXACTEMENT celui du modèle.
    footer_l1 = ("295, Avenue Med VI, OLM Soussi, Rabat – Tél : 0537371045 – "
                 "Fax : 0537375970 – E-mail : direction@romana.ma")
    footer_l2 = ("Patente : 25950420 – RC : 65025 Rabat – I.F : 3334972 – "
                 "CNSS : 7338339 – ICE : 000165779000032")
    pdf.set_auto_page_break(auto=False)          # éviter un saut de page parasite
    pdf.set_y(-22)
    pdf.set_x(14)
    pdf.set_draw_color(20, 58, 130); pdf.set_line_width(0.4)
    pdf.cell(0, 0, "", border="T", ln=1)
    pdf.ln(1.5)
    pdf.set_font(FONT, "", 7.5); pdf.set_text_color(70, 70, 70)
    pdf.set_x(14); pdf.cell(0, 3.6, txt(footer_l1), align="C", ln=1)
    pdf.set_x(14); pdf.cell(0, 3.6, txt(footer_l2), align="C", ln=1)


def _new_pdf():
    from fpdf import FPDF
    pdf = FPDF(format="A4", unit="mm")
    return pdf


def _pdf_via_fpdf(client, out_pdf, logo_uri="", company=None, date_override=""):
    """Rend UNE lettre en PDF (pur Python, hors-ligne, sans installation)."""
    try:
        pdf = _new_pdf()
    except ImportError:
        return False
    try:
        _draw_letter_page(pdf, client, company, date_override)
        pdf.output(out_pdf)
        return os.path.exists(out_pdf)
    except Exception:
        import traceback; traceback.print_exc()
        return False


def combined_pdf(clients, out_pdf, company=None, date_override="", progress=None):
    """Toutes les lettres dans UN seul PDF (une lettre par page)."""
    try:
        pdf = _new_pdf()
    except ImportError:
        return False
    try:
        total = len(clients)
        for i, c in enumerate(clients):
            _draw_letter_page(pdf, c, company, date_override)
            if progress:
                progress(c.get("nom_prenom") or str(c.get("numero", "")),
                         int(100 * (i + 1) / max(1, total)))
        pdf.output(out_pdf)
        return os.path.exists(out_pdf)
    except Exception:
        import traceback; traceback.print_exc()
        return False


def html_to_pdf(html_path, out_pdf, client=None, logo_uri="", company=None, date_override=""):
    """PDF hors-ligne. Priorité : LibreOffice (fidèle au HTML) sinon fpdf2 (intégré)."""
    import subprocess, shutil
    soffice = _find_soffice()
    out_dir = os.path.dirname(out_pdf)
    if soffice:
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                            "--outdir", out_dir, html_path],
                           capture_output=True, timeout=90)
            produced = os.path.join(out_dir, os.path.splitext(os.path.basename(html_path))[0] + ".pdf")
            if os.path.exists(produced):
                if os.path.abspath(produced) != os.path.abspath(out_pdf):
                    shutil.move(produced, out_pdf)
                return True
        except Exception:
            pass
    if shutil.which("wkhtmltopdf"):
        try:
            subprocess.run(["wkhtmltopdf", "--quiet", html_path, out_pdf],
                           capture_output=True, timeout=90)
            if os.path.exists(out_pdf):
                return True
        except Exception:
            pass
    # Repli intégré : fpdf2 (toujours disponible, hors-ligne).
    if client is not None:
        return _pdf_via_fpdf(client, out_pdf, logo_uri, company, date_override)
    return False


def generate_batch(clients, out_dir, formats=("pdf",), date_override="",
                   company=None, naming="", logo_uri="", template=None,
                   progress=None):
    """Génère des fichiers plats (PDF et/ou DOCX) dans `out_dir`.

    Retourne une liste de résultats : {client, name, files:[...], warnings:[...]}.
    """
    os.makedirs(out_dir, exist_ok=True)
    if not logo_uri:
        logo_uri = logo_data_uri()
    results = []
    used = set()
    for i, c in enumerate(clients):
        if progress:
            progress(c.get("nom_prenom") or str(c.get("numero", "")),
                     int(100 * (i + 1) / max(1, len(clients))))
        base = output_filename(c, naming)
        name = base
        n = 2
        while name.lower() in used:
            name = f"{base}_{n}"; n += 1
        used.add(name.lower())

        files, warns = [], []
        want_pdf = "pdf" in formats
        want_docx = "docx" in formats

        html_tmp = os.path.join(out_dir, name + ".html")
        keep_html = False

        if want_pdf:
            pdf_path = os.path.join(out_dir, name + ".pdf")
            # LibreOffice si dispo (fidèle au HTML), sinon fpdf2 intégré.
            if _find_soffice() or __import__("shutil").which("wkhtmltopdf"):
                fill_html(c, html_tmp, logo_uri, company, date_override)
                ok = html_to_pdf(html_tmp, pdf_path, client=c, logo_uri=logo_uri,
                                 company=company, date_override=date_override)
            else:
                ok = _pdf_via_fpdf(c, pdf_path, logo_uri, company, date_override)
            if ok:
                files.append(os.path.basename(pdf_path))
            else:
                fill_html(c, html_tmp, logo_uri, company, date_override)
                warns.append("PDF indisponible — fichier HTML imprimable fourni.")
                files.append(name + ".html"); keep_html = True

        if want_docx:
            docx_path = os.path.join(out_dir, name + ".docx")
            cc = dict(c)
            if date_override:
                cc["date_courrier"] = date_override
            if fill_docx(template, cc, docx_path):
                files.append(os.path.basename(docx_path))
            else:
                warns.append("DOCX indisponible (modèle de lettre introuvable).")

        # Nettoyer le HTML temporaire sauf si on doit le garder.
        if not keep_html and os.path.exists(html_tmp):
            try:
                os.remove(html_tmp)
            except OSError:
                pass

        results.append({"id": c.get("id"), "name": name,
                        "client": c.get("nom_prenom", ""), "files": files,
                        "warnings": warns})
    return results


def preview_html(client, logo_uri="", company=None, date_override=""):
    """HTML d'aperçu (sans bouton d'impression) pour l'affichage dans l'app."""
    if not logo_uri:
        logo_uri = logo_data_uri()
    return build_html(client, logo_uri, company, date_override, show_print=False)


def generate_all(clients, root, template=None, logo_uri="", force=False, progress=None):
    """Génère toutes les lettres (option « auto-générer »)."""
    if template is None:
        template = find_template(root)
    if not logo_uri:
        logo_uri = logo_data_uri()
    for i, c in enumerate(clients):
        if progress:
            progress(f"Lettre : {c.get('nom_prenom') or c['numero']}",
                     int(100 * (i + 1) / max(1, len(clients))))
        paths = ensure_letter(c, root, template, logo_uri, force=force)
        c.update(paths)
    return clients


if __name__ == "__main__":
    import sys
    root = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.dirname(HERE)
    data = json.load(open(os.path.join(HERE, "data.json"), encoding="utf-8"))
    generate_all(data["clients"], root, force="--force" in sys.argv,
                 progress=lambda m, p: print(f"[{p:3d}%] {m}"))
    json.dump(data, open(os.path.join(HERE, "data.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    print(f"[+] Lettres générées dans {os.path.join(root, 'Letters')}")
