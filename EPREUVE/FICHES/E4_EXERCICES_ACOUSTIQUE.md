# EXERCICES CORRIGÉS — ACOUSTIQUE (E4)
> Basés sur les sujets BTS MEC 2023-2025 | Compétence C2 | Savoir S4 + S5

---

## RAPPEL DES FORMULES

```
TR = 0,16 × V / A              temps de réverbération [secondes]
A  = Σ (Si × αi)               surface d'absorption équivalente [m²]
A_nécessaire = 0,16 × V / TR_cible
ΔA = A_nécessaire - A_existante

  V  = volume du local [m³]
  Si = surface de l'élément i [m²]
  αi = coefficient d'absorption (0 à 1, sans unité)
```

---

## EXERCICE 1 — Temps de réverbération d'une salle de classe (type sujet 2023)

### Énoncé
Une salle de classe de dimensions 9 m × 7 m × 3 m est composée des surfaces suivantes :

| Élément | Surface (m²) | Coefficient α |
|---|---|---|
| Plafond (dalle béton brut) | 63 | 0,02 |
| Sol (carrelage) | 63 | 0,02 |
| Murs béton enduit | 66 | 0,03 |
| Fenêtres simple vitrage | 12 | 0,06 |
| Tableau blanc | 6 | 0,04 |
| 30 élèves assis (siège + personne) | — | A = 30 × 0,45 m² |

**Questions :**
1. Calculer le volume V de la salle.
2. Calculer la surface d'absorption totale A.
3. Calculer le temps de réverbération TR.
4. La valeur de TR est-elle acceptable pour une salle de classe (TR ≤ 0,6 s) ?
5. Que faut-il faire pour corriger ?

---

### CORRIGÉ

**1. Volume :**

```
V = 9 × 7 × 3 = 189 m³
```

**2. Surface d'absorption totale A :**

```
A_plafond  = 63  × 0,02 = 1,26 m²
A_sol      = 63  × 0,02 = 1,26 m²
A_murs     = 66  × 0,03 = 1,98 m²
A_fenêtres = 12  × 0,06 = 0,72 m²
A_tableau  = 6   × 0,04 = 0,24 m²
A_élèves   = 30  × 0,45 = 13,50 m²

A_total = 1,26 + 1,26 + 1,98 + 0,72 + 0,24 + 13,50
A_total = 18,96 m²
```

**3. Temps de réverbération :**

```
TR = 0,16 × V / A
TR = 0,16 × 189 / 18,96
TR = 30,24 / 18,96
TR = 1,59 secondes
```

**4. Conformité :**

```
TR calculé = 1,59 s
Exigence   = TR ≤ 0,6 s (salle de classe)

1,59 > 0,6  →  NON CONFORME
→ La salle est trop réverbérante (écho important, mauvaise intelligibilité de la parole).
```

**5. Correction nécessaire :**

```
A_nécessaire = 0,16 × V / TR_cible
A_nécessaire = 0,16 × 189 / 0,6
A_nécessaire = 30,24 / 0,6
A_nécessaire = 50,4 m²

ΔA = A_nécessaire - A_existante
ΔA = 50,4 - 18,96
ΔA = 31,44 m² d'absorption supplémentaire

Solution : Poser un faux-plafond absorbant (α ≈ 0,70 à 0,90)
  S_faux-plafond nécessaire = ΔA / α = 31,44 / 0,75 ≈ 42 m²
  → Recouvrir le plafond entier (63 m²) d'un revêtement absorbant
     suffit largement.
```

---

## EXERCICE 2 — Isolement acoustique des façades (type sujet 2023-2025)

### Énoncé
Un bâtiment de logements est situé à proximité d'une voie routière.
Le bureau d'études indique que la voie est classée **catégorie 3** selon l'arrêté préfectoral.

**Questions :**
1. Quelle est la largeur de secteur affecté par le bruit pour une voie de catégorie 3 ?
2. Quel est l'isolement acoustique minimum DnT,A,tr exigé pour les chambres ?
3. Pour les pièces de vie (séjour, salon) ?
4. L'entrée d'air de ventilation doit-elle être acoustique ? Justifier.

---

### CORRIGÉ

**Classement des voies et secteurs affectés :**

```
Catégorie 1 : L > 81 dB(A)  →  bande de 300 m
Catégorie 2 : L > 76 dB(A)  →  bande de 250 m
Catégorie 3 : L > 70 dB(A)  →  bande de 100 m
Catégorie 4 : L > 65 dB(A)  →  bande de  30 m
Catégorie 5 : L > 60 dB(A)  →  bande de  10 m

→ Voie catégorie 3 : secteur affecté = 100 m de part et d'autre
```

**Isolements exigés (NRA — arrêté du 30/05/1996) :**

```
Pièces principales (chambres, séjour) façades exposées :
  Catégorie 1 : DnT,A,tr ≥ 45 dB
  Catégorie 2 : DnT,A,tr ≥ 42 dB
  Catégorie 3 : DnT,A,tr ≥ 38 dB  ← notre cas
  Catégorie 4 : DnT,A,tr ≥ 35 dB
  Catégorie 5 : DnT,A,tr ≥ 30 dB

→ Chambres ET séjour : DnT,A,tr ≥ 38 dB
```

**Entrées d'air :**

```
OUI — les entrées d'air font partie de la façade acoustique.
Elles constituent le maillon faible de l'isolement.
→ Il faut des entrées d'air acoustiques (avec déflecteur et absorbant)
  adaptées à la catégorie de la voie.
→ Pour catégorie 3 : entrée d'air de type "acoustique renforcé"
  (performance ≥ 38 dB).
```

---

## EXERCICE 3 — Calcul complet (type sujet 2024 — Gymnase)

### Énoncé
Un gymnase municipal a les caractéristiques suivantes :
- Dimensions : 44 m × 24 m × 8 m
- Surfaces et matériaux :

| Élément | Surface (m²) | α |
|---|---|---|
| Sol (parquet) | 1 056 | 0,06 |
| Plafond (béton) | 1 056 | 0,02 |
| Murs longs (béton enduit) | 704 | 0,03 |
| Murs courts (béton enduit) | 384 | 0,03 |
| Tribunes (gradins bois vides) | 80 | 0,04 |

On veut obtenir un TR ≤ 1,5 s (exigence ERP sportif).

**Questions :**
1. Calculer V et A_existante.
2. Calculer le TR actuel.
3. Calculer ΔA nécessaire.
4. On pose des panneaux absorbants sur le plafond (α = 0,80). Quelle surface couvrir ?

---

### CORRIGÉ

**1. Volume et absorption existante :**

```
V = 44 × 24 × 8 = 8 448 m³

A_sol      = 1 056 × 0,06 = 63,36 m²
A_plafond  = 1 056 × 0,02 = 21,12 m²
A_murs L   =   704 × 0,03 = 21,12 m²
A_murs C   =   384 × 0,03 = 11,52 m²
A_tribunes =    80 × 0,04 =  3,20 m²

A_existante = 63,36 + 21,12 + 21,12 + 11,52 + 3,20
A_existante = 120,32 m²
```

**2. TR actuel :**

```
TR = 0,16 × 8 448 / 120,32
TR = 1 351,68 / 120,32
TR = 11,23 secondes  →  très fortement réverbérant
```

**3. ΔA nécessaire pour TR = 1,5 s :**

```
A_nécessaire = 0,16 × 8 448 / 1,5
A_nécessaire = 1 351,68 / 1,5
A_nécessaire = 901,12 m²

ΔA = 901,12 - 120,32 = 780,80 m²
```

**4. Surface de panneaux absorbants (α = 0,80) :**

```
S_panneaux = ΔA / α = 780,80 / 0,80 = 976 m²

Le plafond fait 1 056 m² → couvrir 976 / 1 056 = 92,4% du plafond.
→ Solution réaliste : couvrir la totalité du plafond de panneaux absorbants.
```

---

## POINTS DE VIGILANCE

- **Distinguer TR (réverbération locale)** et **DnT,A,tr (isolement de façade)**
- La formule de Sabine TR = 0,16 × V / A s'applique à l'acoustique **interne** d'un local
- Le classement de voie (catégorie 1 à 5) détermine l'isolement **de façade**
- Les entrées d'air VMC sont toujours un point faible de la façade acoustique
- **Unités** : V en m³, A en m², TR en secondes
