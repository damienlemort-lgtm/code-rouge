# -*- coding: utf-8 -*-
"""Génère le guide de RENDEZ-VOUS COMMERCIAL ANDRAGOPS x COUSIN GROUP."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BLEU = RGBColor(0x1F, 0x3A, 0x5F)
ROUGE = RGBColor(0xB3, 0x1B, 0x1B)
GRIS = RGBColor(0x44, 0x44, 0x44)
VERT = RGBColor(0x1E, 0x6B, 0x3A)


def set_base_style(doc):
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    style.font.color.rgb = GRIS


def add_cover(doc, title, subtitle, meta_lines):
    for _ in range(4):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title); r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = BLEU
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle); r.font.size = Pt(15); r.font.color.rgb = ROUGE
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('—' * 20); r.font.color.rgb = BLEU
    for line in meta_lines:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(line); r.font.size = Pt(11); r.font.color.rgb = GRIS
    doc.add_page_break()


def h1(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text); r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = BLEU
    pPr = p._p.get_or_add_pPr(); pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for k, v in (('w:val', 'single'), ('w:sz', '6'), ('w:space', '4'), ('w:color', 'B31B1B')):
        bottom.set(qn(k), v)
    pbdr.append(bottom); pPr.append(pbdr)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text); r.font.size = Pt(12.5); r.font.bold = True; r.font.color.rgb = ROUGE
    return p


def para(doc, text, bold=False):
    p = doc.add_paragraph(); r = p.add_run(text); r.font.bold = bold
    return p


def bullet(doc, text, sub=False):
    p = doc.add_paragraph(style='List Bullet 2' if sub else 'List Bullet')
    for i, part in enumerate(text.split('**')):
        r = p.add_run(part)
        if i % 2 == 1:
            r.font.bold = True
    return p


def script_line(doc, label, text):
    p = doc.add_paragraph()
    r = p.add_run(label + ' '); r.font.bold = True; r.font.color.rgb = ROUGE
    r2 = p.add_run('« ' + text + ' »'); r2.font.italic = True
    return p


def add_info_table(doc, rows, c0=2.3, c1=4.0):
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Light Grid Accent 1'; table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for k, v in rows:
        cells = table.add_row().cells
        cells[0].width = Inches(c0); cells[1].width = Inches(c1)
        rk = cells[0].paragraphs[0].add_run(k); rk.font.bold = True; rk.font.size = Pt(10.5); rk.font.color.rgb = BLEU
        rv = cells[1].paragraphs[0].add_run(v); rv.font.size = Pt(10.5)
    doc.add_paragraph()
    return table


def add_objection(doc, objection, reponse):
    p = doc.add_paragraph()
    r = p.add_run('« ' + objection + ' »'); r.font.bold = True; r.font.color.rgb = BLEU
    p2 = doc.add_paragraph()
    r2 = p2.add_run('→ Réponse : '); r2.font.bold = True; r2.font.color.rgb = VERT
    p2.add_run(reponse)
    p2.paragraph_format.space_after = Pt(10)


def fill_lines(doc, n=3):
    for _ in range(n):
        doc.add_paragraph("……………………………………………………………………………………………………")


doc = Document()
set_base_style(doc)

add_cover(
    doc,
    "Guide de rendez-vous commercial",
    "ANDRAGOPS Académie  ✕  COUSIN GROUP",
    [
        "Échange téléphonique avec la Responsable RH",
        "Objectif : présenter nos formations & identifier les besoins des équipes",
        "Angle : sur-mesure & ancrage local (Hauts-de-France)",
        "Contact : +33 (0)6 35 25 89 22 — Wervicq-Sud (59)",
        "",
        "Guide préparé le 19 juin 2026",
    ],
)

# 1. Objectif & état d'esprit
h1(doc, "1. Objectif & état d'esprit du rendez-vous")
para(doc, "Ce n'est pas un argumentaire « descendant ». C'est un échange de découverte : on présente "
          "brièvement ANDRAGOPS, puis on fait parler la RH pour cerner les besoins réels des équipes, "
          "et on co-construit une piste de formation sur-mesure.")
add_info_table(doc, [
    ("But de l'appel", "Créer le lien, qualifier les besoins, décrocher un rendez-vous de cadrage"),
    ("Ce que je vends", "Des formations sur-mesure, ancrées localement, adaptées à leurs équipes"),
    ("Posture", "Expert-conseil à l'écoute (80 % écoute / 20 % parole), pas de monologue"),
    ("Résultat visé", "Un prochain RDV (audit des besoins / proposition) + un contact identifié"),
])
para(doc, "Règle d'or : la RH doit parler plus que moi. Mon rôle = poser les bonnes questions et "
          "reformuler leurs besoins en solutions.")

# 2. Le prospect en bref
h1(doc, "2. Le prospect en 30 secondes (rappel)")
bullet(doc, "**COUSIN GROUP** — groupe industriel familial de Wervicq-Sud (59), fondé en **1848**.")
bullet(doc, "3 pôles : **Cousin Trestec** (cordages techniques), **Cousin Composites** (joncs/profilés), "
            "**Cousin Surgery** (implants médicaux).")
bullet(doc, "≈ **150 collaborateurs** ; valeurs affichées : **exigence, qualité, innovation** (~5 % du CA en R&D).")
bullet(doc, "**175 ans de savoir-faire** dans la transformation des fibres → fort enjeu de transmission.")
bullet(doc, "Recrute des profils de production/encadrement (managers de production, responsables atelier, qualité, IT).")

# 3. Lecture des besoins probables -> opportunités de formation
h1(doc, "3. Besoins probables → pistes de formation (à valider en RDV)")
para(doc, "Hypothèses à confirmer par les questions de découverte. Ne pas les asséner : les amener "
          "sous forme de questions.")

h2(doc, "A. Transmission du savoir-faire (votre angle fort)")
bullet(doc, "Enjeu : une maison de 175 ans = des experts métier dont le savoir doit se transmettre.")
bullet(doc, "Pistes : **formation de formateurs / tuteurs internes**, AFEST (formation en situation de travail), "
            "ingénierie de transmission, mentorat des nouveaux.")

h2(doc, "B. Management & encadrement de proximité")
bullet(doc, "Ils recrutent des managers de production / responsables d'atelier.")
bullet(doc, "Pistes : management d'équipe, communication, gestion des conflits, conduite du changement, "
            "intégration des nouveaux managers.")

h2(doc, "C. IA & outils numériques (productivité)")
bullet(doc, "Une entreprise très orientée R&D et innovation → gains possibles au bureau d'études, "
            "au commercial, à l'administratif.")
bullet(doc, "Pistes : IA générative au quotidien, outils bureautiques avancés, digitalisation des process.")

h2(doc, "D. Qualité, sécurité & compétences métier")
bullet(doc, "Milieu industriel exigeant (et salle blanche côté médical).")
bullet(doc, "Pistes : qualité, sécurité au poste, habilitations, montée en compétences techniques.")

para(doc, "→ Adaptez ces pistes à VOTRE catalogue ANDRAGOPS réel. Notez vos 2-3 offres phares ci-dessous :")
fill_lines(doc, 3)

# 4. Pitch ANDRAGOPS
h1(doc, "4. Pitch ANDRAGOPS (30–60 secondes)")
para(doc, "Trame à personnaliser avec vos mots :")
script_line(doc, "Accroche :", "Bonjour, [Prénom Nom], d'ANDRAGOPS Académie. Merci de prendre le temps. "
            "On accompagne les entreprises industrielles comme la vôtre sur la montée en compétences de leurs équipes.")
script_line(doc, "Qui :", "ANDRAGOPS, c'est un organisme de formation [Qualiopi le cas échéant], "
            "spécialisé dans des formations sur-mesure, conçues à partir de vos besoins réels, et de proximité.")
script_line(doc, "Pourquoi vous :", "J'ai vu que Cousin est une maison de 175 ans, très attachée à la qualité "
            "et à l'innovation — c'est exactement le type d'entreprise pour qui le sur-mesure fait la différence.")
script_line(doc, "Transition :", "Avant de vous présenter ce qu'on fait, j'aimerais d'abord comprendre vos "
            "enjeux RH du moment. Je peux vous poser quelques questions ?")
para(doc, "Votre version personnalisée :")
fill_lines(doc, 3)

# 5. Questions de découverte
h1(doc, "5. Questions de découverte (le cœur de l'appel)")
para(doc, "Méthode : Situation → Besoins → Impact → Décision. Laissez des silences, notez tout.")

h2(doc, "Situation / contexte")
bullet(doc, "Comment est organisée la formation chez vous aujourd'hui (interne, externe, plan de développement des compétences) ?")
bullet(doc, "Quels sont vos principaux enjeux RH cette année (recrutement, fidélisation, montée en compétences) ?")
bullet(doc, "Sur quels métiers avez-vous le plus de tension ou de turn-over ?")

h2(doc, "Besoins")
bullet(doc, "Y a-t-il des compétences clés détenues par quelques experts que vous craignez de perdre (départs, retraites) ?")
bullet(doc, "Comment intégrez-vous et formez-vous vos nouveaux arrivants aujourd'hui ?")
bullet(doc, "Vos managers de proximité sont-ils accompagnés/formés à l'encadrement ?")
bullet(doc, "Où aimeriez-vous que vos équipes progressent en priorité dans les 12 prochains mois ?")

h2(doc, "Impact & décision")
bullet(doc, "Qu'est-ce que ça vous coûterait de ne rien faire sur ce sujet ?")
bullet(doc, "Qui est impliqué dans le choix d'un organisme de formation ?")
bullet(doc, "Avez-vous un budget formation déjà engagé via votre OPCO pour cette année ?")

# 6. Arguments différenciants
h1(doc, "6. Vos arguments différenciants (sur-mesure & local)")
bullet(doc, "**Sur-mesure** : on ne vend pas un catalogue figé, on construit la formation à partir de VOS process, "
            "VOTRE vocabulaire métier, VOS cas concrets.")
bullet(doc, "**Local / proximité** : ancrage Hauts-de-France, intervenants qui se déplacent sur site, réactivité, "
            "relation humaine durable (en résonance avec leur ADN familial et territorial).")
bullet(doc, "**Andragogie** : pédagogie pensée pour des adultes en activité — concret, opérationnel, applicable tout de suite.")
bullet(doc, "**Souplesse** : formats courts, en présentiel sur site, adaptés aux contraintes de production.")
bullet(doc, "**Financement facilité** : [si Qualiopi] prise en charge possible via l'OPCO et le plan de développement des compétences.")

# 7. Objections
h1(doc, "7. Objections fréquentes & réponses")
add_objection(doc, "On a déjà notre organisme de formation habituel.",
              "« Très bien, ce n'est pas pour le remplacer. Justement, sur les besoins très spécifiques à "
              "vos métiers, le sur-mesure peut compléter ce qui existe. On peut commencer par un sujet précis. »")
add_objection(doc, "On n'a pas de budget en ce moment.",
              "« Je comprends. Beaucoup de nos formations sont finançables par votre OPCO via le plan de "
              "développement des compétences — on peut regarder ensemble ce qui est mobilisable. »")
add_objection(doc, "On n'a pas le temps, on est en pleine production.",
              "« C'est exactement pour ça qu'on construit du sur-mesure : formats courts, sur site, calés sur "
              "vos contraintes d'atelier, sans casser la production. »")
add_objection(doc, "Envoyez-moi une plaquette, je regarderai.",
              "« Avec plaisir, je vous l'envoie. Mais une plaquette générique ne dira rien de VOS besoins. "
              "Accordez-moi 30 minutes pour cadrer un ou deux sujets, et je reviens avec une proposition ciblée. »")
add_objection(doc, "Pourquoi vous plutôt qu'un gros organisme national ?",
              "« Parce qu'on est sur-mesure et local : on vient sur site, on parle votre métier, et vous avez "
              "un interlocuteur unique et réactif, pas un numéro de dossier. »")

# 8. Financement
h1(doc, "8. Financement (à maîtriser)")
bullet(doc, "**Qualiopi** : certification qui ouvre les financements publics/mutualisés — mentionnez-la si vous l'avez.")
bullet(doc, "**OPCO** : Cousin (textile technique / plasturgie / chimie) dépend a priori de l'**OPCO 2i** "
            "(industrie) — à confirmer ; c'est lui qui peut financer le plan de développement des compétences.")
bullet(doc, "**Plan de développement des compétences** : le levier principal côté entreprise.")
bullet(doc, "Objectif : montrer que le frein budgétaire est souvent surmontable.")

# 9. Closing
h1(doc, "9. Conclure : la prochaine étape")
para(doc, "Ne jamais raccrocher sans une étape suivante datée.")
script_line(doc, "Proposition :", "Ce que je vous propose : un rendez-vous de 30–45 min, sur site ou en visio, "
            "pour cadrer 1 ou 2 sujets prioritaires. Je reviens ensuite avec une proposition sur-mesure et chiffrée.")
bullet(doc, "Proposez **2 créneaux précis** plutôt qu'un vague « quand vous voulez ».")
bullet(doc, "Validez : interlocuteur, email, prochain RDV, ce que vous envoyez d'ici là.")
bullet(doc, "Reformulez les besoins entendus avant de raccrocher (« si j'ai bien compris, vos priorités sont… »).")

# 10. Check-list + notes
h1(doc, "10. Check-list avant l'appel")
bullet(doc, "Dossier de recherche COUSIN GROUP + ce guide sous les yeux.")
bullet(doc, "Mes 2-3 offres phares et un exemple de réussite client prêts à citer.")
bullet(doc, "Mes créneaux de RDV de cadrage disponibles.")
bullet(doc, "Endroit calme, bon réseau, écouteurs, eau, sourire (s'entend au téléphone).")
bullet(doc, "De quoi noter : besoins exprimés, budget, décideurs, prochaine étape.")

h1(doc, "11. Notes de l'appel")
para(doc, "Interlocuteur / fonction :"); fill_lines(doc, 1)
para(doc, "Besoins exprimés :"); fill_lines(doc, 3)
para(doc, "Budget / OPCO :"); fill_lines(doc, 1)
para(doc, "Décideurs impliqués :"); fill_lines(doc, 1)
para(doc, "Prochaine étape & date :"); fill_lines(doc, 2)

doc.save('/home/user/code-rouge/Guide_RDV_Commercial_COUSIN_GROUP.docx')
print("Guide commercial OK")
