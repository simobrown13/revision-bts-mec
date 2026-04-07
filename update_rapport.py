#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Update RAPPORT U62: Add 2 new pages with SVG donut charts,
renumber all pages, update TOC.
Result: 30 content pages + 25 annexes = 55 pages total.
"""

FILE_PATH = r'D:\PREPA BTS MEC\08_U62_Rapport_Activites\Rapport_Final\RAPPORT_U62_BAHAFID_Mohamed.html'

# ===========================
# NEW PAGE A: Missions Projet 2 (insert before original PAGE 20)
# ===========================
NEW_PAGE_A = '''

<!-- ================================================================
     PAGE NEW_A : PROJET 2 - Missions r&eacute;alis&eacute;es
     ================================================================ -->
<div class="page content">
    <div class="page-header">
        <div class="page-header__left">Rapport U62 &mdash; BTS MEC 2026</div>
        <div class="page-header__center"><div class="page-header__center-diamond"><span>&#9670;</span></div></div>
        <div class="page-header__right">BAHAFID Mohamed</div>
    </div>
    <div class="corner-tl"></div><div class="corner-tr"></div><div class="corner-bl"></div><div class="corner-br"></div>
    <div class="content__topbar">
        <span class="content__topbar-section">Partie 2 &bull; Projets r&eacute;alis&eacute;s</span>
        <span class="content__topbar-title">2.2 Missions r&eacute;alis&eacute;es</span>
    </div>

    <h3 class="ssub-title">Missions r&eacute;alis&eacute;es sur le Projet 2</h3>

    <p>Sur ce projet routier de <strong>25 km en zone montagneuse</strong>, j'ai assur&eacute; quatre missions principales couvrant l'ensemble du cycle du march&eacute; public, de l'estimation initiale &agrave; la r&eacute;ception des travaux.</p>

    <div style="display:flex;align-items:flex-start;gap:20px;margin:10px 0 14px;">
        <svg viewBox="0 0 42 42" width="120" height="120" style="flex-shrink:0;">
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#1B2A4A" stroke-width="6" stroke-dasharray="30.1 69.9" stroke-dashoffset="25"/>
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#E63946" stroke-width="6" stroke-dasharray="28.1 71.9" stroke-dashoffset="-5.1"/>
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#4A90D9" stroke-width="6" stroke-dasharray="17.9 82.1" stroke-dashoffset="-33.2"/>
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#24375E" stroke-width="6" stroke-dasharray="11.3 88.7" stroke-dashoffset="-51.1"/>
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#F47D86" stroke-width="6" stroke-dasharray="10.7 89.3" stroke-dashoffset="-62.4"/>
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#C0C1CE" stroke-width="6" stroke-dasharray="1.9 98.1" stroke-dashoffset="-73.1"/>
            <text x="21" y="20" text-anchor="middle" dominant-baseline="central" font-family="Montserrat,sans-serif" font-weight="700" font-size="4.5" fill="#1B2A4A">24,2 M</text>
            <text x="21" y="24.5" text-anchor="middle" font-family="Inter,sans-serif" font-size="2.8" fill="#5A5D6E">DH HT</text>
        </svg>
        <div style="font-size:8.5pt;line-height:1.8;">
            <div><span style="display:inline-block;width:10px;height:10px;background:#1B2A4A;margin-right:6px;vertical-align:middle;border-radius:2px;"></span>Corps de chauss&eacute;e &mdash; 30,1%</div>
            <div><span style="display:inline-block;width:10px;height:10px;background:#E63946;margin-right:6px;vertical-align:middle;border-radius:2px;"></span>Ouvrages hydrauliques &mdash; 28,1%</div>
            <div><span style="display:inline-block;width:10px;height:10px;background:#4A90D9;margin-right:6px;vertical-align:middle;border-radius:2px;"></span>Terrassement &mdash; 17,9%</div>
            <div><span style="display:inline-block;width:10px;height:10px;background:#24375E;margin-right:6px;vertical-align:middle;border-radius:2px;"></span>Bretelles &mdash; 11,3%</div>
            <div><span style="display:inline-block;width:10px;height:10px;background:#F47D86;margin-right:6px;vertical-align:middle;border-radius:2px;"></span>Rev&ecirc;tement &mdash; 10,7%</div>
        </div>
    </div>

    <h3 class="ssub-title">a) Estimation et m&eacute;tr&eacute;s</h3>
    <p>J'ai r&eacute;alis&eacute; les avant-m&eacute;tr&eacute;s &agrave; partir des profils en long et en travers fournis par le bureau d'&eacute;tudes topographiques. Les cubatures de terrassement (<strong>120 334 m&sup3;</strong> de d&eacute;blais, <strong>76 735 m&sup3;</strong> de remblais) ont &eacute;t&eacute; calcul&eacute;es par la m&eacute;thode des profils en travers successifs.</p>

    <h3 class="ssub-title">b) R&eacute;daction du CPS et BPDE</h3>
    <p>J'ai contribu&eacute; &agrave; la r&eacute;daction des clauses techniques du CPS et &agrave; l'&eacute;tablissement du BPDE structur&eacute; en 3 parties : section lin&eacute;aire (23 prix), carrefour PK 0+000 (11 prix) et bretelles (19 prix). Chaque prix unitaire a &eacute;t&eacute; d&eacute;compos&eacute; selon le sous-d&eacute;tail des prix.</p>

    <h3 class="ssub-title">c) Suivi technique et r&eacute;ception</h3>
    <p>Pendant l'ex&eacute;cution, j'ai assur&eacute; le suivi des attachements contradictoires, le contr&ocirc;le des situations de travaux et la pr&eacute;paration des d&eacute;comptes provisoires et d&eacute;finitifs.</p>

    <div class="hl-box">
        <h4>&#9670; Sp&eacute;cificit&eacute; montagne</h4>
        <p>Le relief du Moyen Atlas a impos&eacute; des adaptations co&ucirc;teuses : terrassements massifs, murs de sout&egrave;nement en gabions (789 m&sup3;) et 23 ouvrages hydrauliques de travers&eacute;e. Cette complexit&eacute; a renforc&eacute; mes comp&eacute;tences en estimation d'ouvrages sp&eacute;ciaux.</p>
    </div>

    <div class="page-footer">
        <div class="page-footer__left">BAHAFID Mohamed &bull; BTS MEC 2026</div>
        <div class="page-footer__center">Partie 2 &mdash; Projets R&eacute;alis&eacute;s</div>
        <div class="page-footer__right"><div class="page-footer__right-num"><span>NEW_A</span></div></div>
    </div>
</div>

'''

# ===========================
# NEW PAGE B: Comparative detaillee (insert before original PAGE 26)
# ===========================
NEW_PAGE_B = '''

<!-- ================================================================
     PAGE NEW_B : COMPARATIVE D&Eacute;TAILL&Eacute;E
     ================================================================ -->
<div class="page content">
    <div class="page-header">
        <div class="page-header__left">Rapport U62 &mdash; BTS MEC 2026</div>
        <div class="page-header__center"><div class="page-header__center-diamond"><span>&#9670;</span></div></div>
        <div class="page-header__right">BAHAFID Mohamed</div>
    </div>
    <div class="corner-tl"></div><div class="corner-tr"></div><div class="corner-bl"></div><div class="corner-br"></div>
    <div class="content__topbar">
        <span class="content__topbar-section">Partie 3 &bull; Analyse et comp&eacute;tences</span>
        <span class="content__topbar-title">3.2 Comparative d&eacute;taill&eacute;e &mdash; M&eacute;thodologies</span>
    </div>

    <h3 class="ssub-title">Comparaison d&eacute;taill&eacute;e des m&eacute;thodologies</h3>

    <div class="table-wrap">
        <table class="cmp">
            <thead><tr><th>&Eacute;tape du projet</th><th>Pratique au Maroc</th><th>Pratique en France</th></tr></thead>
            <tbody>
                <tr><td><strong>Programmation</strong></td><td>PDR r&eacute;gional, programmes INDH</td><td>PPI, contrats de plan &Eacute;tat-R&eacute;gion</td></tr>
                <tr><td><strong>Estimation</strong></td><td>Estimation confidentielle de l'administration</td><td>Estimation pr&eacute;visionnelle du MOA</td></tr>
                <tr><td><strong>Consultation</strong></td><td>Appel d'offres ouvert (d&eacute;cret 2-12-349)</td><td>MAPA ou proc&eacute;dure formalis&eacute;e (CCP)</td></tr>
                <tr><td><strong>S&eacute;lection</strong></td><td>Commission AO, jugement technique + financier</td><td>CAO, crit&egrave;res pond&eacute;r&eacute;s (prix, technique, d&eacute;lai)</td></tr>
                <tr><td><strong>Ex&eacute;cution</strong></td><td>OS, attachements contradictoires, d&eacute;comptes</td><td>OS, situations mensuelles, avenants</td></tr>
                <tr><td><strong>R&eacute;ception</strong></td><td>R&eacute;ception provisoire + d&eacute;finitive (1 an)</td><td>OPR, r&eacute;ception + GPA (1 an)</td></tr>
            </tbody>
        </table>
    </div>

    <h3 class="ssub-title">R&eacute;partition de mon exp&eacute;rience par domaine</h3>

    <div style="display:flex;align-items:flex-start;gap:20px;margin:10px 0 14px;">
        <svg viewBox="0 0 42 42" width="120" height="120" style="flex-shrink:0;">
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#1B2A4A" stroke-width="6" stroke-dasharray="40 60" stroke-dashoffset="25"/>
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#E63946" stroke-width="6" stroke-dasharray="35 65" stroke-dashoffset="-15"/>
            <circle cx="21" cy="21" r="15.9155" fill="transparent" stroke="#4A90D9" stroke-width="6" stroke-dasharray="25 75" stroke-dashoffset="-50"/>
            <text x="21" y="20" text-anchor="middle" dominant-baseline="central" font-family="Montserrat,sans-serif" font-weight="700" font-size="5" fill="#1B2A4A">8+</text>
            <text x="21" y="24.5" text-anchor="middle" font-family="Inter,sans-serif" font-size="2.8" fill="#5A5D6E">ann&eacute;es</text>
        </svg>
        <div style="font-size:8.5pt;line-height:1.8;">
            <div><span style="display:inline-block;width:10px;height:10px;background:#1B2A4A;margin-right:6px;vertical-align:middle;border-radius:2px;"></span>MOA publique (Maroc) &mdash; 4,5 ans &mdash; 40%</div>
            <div><span style="display:inline-block;width:10px;height:10px;background:#E63946;margin-right:6px;vertical-align:middle;border-radius:2px;"></span>Ex&eacute;cution terrain (France) &mdash; 3 ans &mdash; 35%</div>
            <div><span style="display:inline-block;width:10px;height:10px;background:#4A90D9;margin-right:6px;vertical-align:middle;border-radius:2px;"></span>BIM &amp; Num&eacute;rique &mdash; 2 ans &mdash; 25%</div>
        </div>
    </div>

    <div class="hl-box">
        <h4>&#9670; Enrichissement mutuel des deux syst&egrave;mes</h4>
        <p>Les m&eacute;thodes marocaines (attachements contradictoires, BPDE d&eacute;taill&eacute;s, estimation confidentielle) et fran&ccedil;aises (situations de travaux, DQE, crit&egrave;res pond&eacute;r&eacute;s) se compl&egrave;tent. Cette double exp&eacute;rience enrichit ma capacit&eacute; &agrave; analyser les co&ucirc;ts et &agrave; piloter les budgets dans des contextes r&eacute;glementaires vari&eacute;s.</p>
    </div>

    <p>Cette analyse comparative illustre comment les <strong>principes fondamentaux</strong> de la commande publique (transparence, &eacute;galit&eacute; de traitement, mise en concurrence) se d&eacute;clinent diff&eacute;remment selon les deux syst&egrave;mes juridiques, tout en poursuivant les m&ecirc;mes objectifs de qualit&eacute; et d'efficience.</p>

    <div class="page-footer">
        <div class="page-footer__left">BAHAFID Mohamed &bull; BTS MEC 2026</div>
        <div class="page-footer__center">Partie 3 &mdash; Analyse et Comp&eacute;tences</div>
        <div class="page-footer__right"><div class="page-footer__right-num"><span>NEW_B</span></div></div>
    </div>
</div>

'''


def main():
    with open(FILE_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    original_len = len(html)
    print(f"Original file: {original_len} chars")

    # ===========================
    # STEP 1: Insert page A before PAGE 20
    # ===========================
    marker_A = 'PAGE 20 :'
    pos_A = html.find(marker_A)
    if pos_A == -1:
        print("ERROR: Could not find PAGE 20 marker")
        return
    # Go back to find the <!-- that starts this comment
    comment_start_A = html.rfind('<!--', 0, pos_A)
    if comment_start_A == -1:
        print("ERROR: Could not find comment start for PAGE 20")
        return

    html = html[:comment_start_A] + NEW_PAGE_A + html[comment_start_A:]
    print("Inserted new page A (Missions Projet 2) before PAGE 20")

    # ===========================
    # STEP 2: Insert page B before PAGE 26 (original numbering)
    # PAGE 26 still exists with original numbering since we haven't renumbered yet
    # ===========================
    marker_B = 'PAGE 26 :'
    pos_B = html.find(marker_B)
    if pos_B == -1:
        print("ERROR: Could not find PAGE 26 marker")
        return
    comment_start_B = html.rfind('<!--', 0, pos_B)
    if comment_start_B == -1:
        print("ERROR: Could not find comment start for PAGE 26")
        return

    html = html[:comment_start_B] + NEW_PAGE_B + html[comment_start_B:]
    print("Inserted new page B (Comparative detaillee) before PAGE 26")

    # ===========================
    # STEP 3: Renumber pages
    # After insertions, the structure is:
    #   Pages 1-19: unchanged
    #   NEW_A: should be page 20
    #   Original 20-25: should be 21-26
    #   NEW_B: should be page 27
    #   Original 26-28: should be 28-30
    #   Original 29-53: should be 31-55
    #
    # Strategy: pages 26-53 shift by +2, pages 20-25 shift by +1
    # Process from highest to lowest to avoid conflicts
    # ===========================

    # Group 1: pages 26-53 shift by +2 (process high to low)
    for old_num in range(53, 25, -1):
        new_num = old_num + 2

        # Update page comments
        old_comment = f'PAGE {old_num} :'
        new_comment = f'PAGE {new_num} :'
        count = html.count(old_comment)
        if count > 0:
            html = html.replace(old_comment, new_comment)
            # print(f"  Comment: PAGE {old_num} -> PAGE {new_num} ({count}x)")

        # Update footer page numbers
        old_footer = f'page-footer__right-num"><span>{old_num}</span>'
        new_footer = f'page-footer__right-num"><span>{new_num}</span>'
        count = html.count(old_footer)
        if count > 0:
            html = html.replace(old_footer, new_footer)
            # print(f"  Footer: {old_num} -> {new_num} ({count}x)")

    print("Renumbered pages 26-53 -> 28-55 (+2)")

    # Group 2: pages 20-25 shift by +1 (process high to low)
    for old_num in range(25, 19, -1):
        new_num = old_num + 1

        old_comment = f'PAGE {old_num} :'
        new_comment = f'PAGE {new_num} :'
        count = html.count(old_comment)
        if count > 0:
            html = html.replace(old_comment, new_comment)

        old_footer = f'page-footer__right-num"><span>{old_num}</span>'
        new_footer = f'page-footer__right-num"><span>{new_num}</span>'
        count = html.count(old_footer)
        if count > 0:
            html = html.replace(old_footer, new_footer)

    print("Renumbered pages 20-25 -> 21-26 (+1)")

    # ===========================
    # STEP 4: Replace placeholder numbers for new pages
    # ===========================
    html = html.replace('PAGE NEW_A :', 'PAGE 20 :')
    html = html.replace('<span>NEW_A</span>', '<span>20</span>')
    html = html.replace('PAGE NEW_B :', 'PAGE 27 :')
    html = html.replace('<span>NEW_B</span>', '<span>27</span>')
    print("Replaced placeholder numbers: NEW_A -> 20, NEW_B -> 27")

    # ===========================
    # STEP 5: Update TOC page numbers
    # Original TOC:
    #   Introduction: 3 (no change)
    #   Partie 1: 6 (no change)
    #   1.1: 6, 1.2: 9, 1.3: 10 (no change)
    #   Partie 2: 13 (no change)
    #   2.1: 13, 2.2: 18 (no change)
    #   2.3: 22 -> 23 (+1)
    #   Partie 3: 24 -> 25 (+1)
    #   3.1: 24 -> 25 (+1)
    #   3.2: 25 -> 26 (+1)
    #   3.3: 26 -> 28 (+2)
    #   Conclusion: 28 -> 30 (+2)
    #   Annexes: 29 -> 31 (+2)
    # Process high to low within the TOC context
    # ===========================

    toc_replacements = [
        (29, 31),
        (28, 30),
        (26, 28),
        (25, 26),
        (24, 25),
        (22, 23),
    ]

    for old_pg, new_pg in toc_replacements:
        old_toc = f'toc__pg">{old_pg}</span>'
        new_toc = f'toc__pg">{new_pg}</span>'
        count = html.count(old_toc)
        if count > 0:
            html = html.replace(old_toc, new_toc)
            print(f"  TOC: {old_pg} -> {new_pg} ({count}x)")

    print("Updated TOC page numbers")

    # ===========================
    # STEP 6: Verify
    # ===========================
    # Count pages
    page_count = html.count('<div class="page ')
    print(f"\nTotal pages in document: {page_count}")
    print(f"File size: {original_len} -> {len(html)} chars")

    # Check for any remaining placeholders
    if 'NEW_A' in html or 'NEW_B' in html:
        print("WARNING: Placeholder markers still present!")
    else:
        print("All placeholders replaced successfully")

    # Check page number sequence in footers
    import re
    footer_nums = re.findall(r'page-footer__right-num"><span>(\d+)</span>', html)
    footer_nums = [int(x) for x in footer_nums]
    print(f"Footer page numbers: {footer_nums[:5]}...{footer_nums[-5:]}")
    print(f"Min: {min(footer_nums)}, Max: {max(footer_nums)}, Count: {len(footer_nums)}")

    # Check for gaps
    expected = list(range(2, max(footer_nums) + 1))
    actual = sorted(footer_nums)
    missing = set(expected) - set(actual)
    dupes = [x for x in actual if actual.count(x) > 1]
    if missing:
        print(f"WARNING: Missing page numbers: {sorted(missing)}")
    if dupes:
        print(f"WARNING: Duplicate page numbers: {sorted(set(dupes))}")
    if not missing and not dupes:
        print("Page number sequence OK (no gaps, no duplicates)")

    # ===========================
    # STEP 7: Write output
    # ===========================
    with open(FILE_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\nFile written: {FILE_PATH}")
    print("Done!")


if __name__ == '__main__':
    main()
