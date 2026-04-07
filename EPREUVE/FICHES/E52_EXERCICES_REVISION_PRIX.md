# EXERCICES CORRIGÉS — RÉVISION DES PRIX (E52)
> Basés sur les sujets BTS MEC 2023-2025 | Compétence C17-1 | Savoir S7

---

## RAPPEL DES FORMULES

```
Cn = 0,15 + 0,85 × (Im / I0)
  Cn = coefficient de révision (arrondi au millième SUPÉRIEUR)
  Im = indice BT du mois de révision (parfois Im-3 ou Im-4 selon CCAP)
  I0 = indice BT du mois zéro (dépôt de l'offre)

Révision HT = Travaux HT du mois × (Cn - 1)

Formule à deux indices :
  Cn = 0,15 + 0,85 × [(Im_BT1/I0_BT1) × p1% + (Im_BT2/I0_BT2) × p2%]

P actualisé = P0 × (Im-3 / I0)
```

---

## EXERCICE 1 — Révision simple avec un seul indice (type sujet 2023)

### Énoncé
Un marché de gros-œuvre est passé avec les données suivantes :
- Indice de référence (I0, mois 0 = dépôt offre) : **BT01 = 112,4**
- Indice du mois de réalisation (Im) : **BT01 = 121,7**
- Travaux HT réalisés le mois M : **38 500 € HT**

Le CCAP prévoit la formule : `Cn = 0,15 + 0,85 × (Im / I0)`

**Questions :**
1. Calculer le coefficient de révision Cn (arrondir au millième supérieur).
2. Calculer le montant de la révision HT pour le mois M.
3. Quel est le montant HT total facturé ce mois (travaux + révision) ?

---

### CORRIGÉ

**1. Coefficient Cn :**

```
Cn = 0,15 + 0,85 × (Im / I0)
Cn = 0,15 + 0,85 × (121,7 / 112,4)
Cn = 0,15 + 0,85 × 1,08274...
Cn = 0,15 + 0,92033...
Cn = 1,07033...

→ Arrondi au millième SUPÉRIEUR : Cn = 1,071
```

> **Attention** : l'arrondi est TOUJOURS au millième supérieur (jamais inférieur).
> 1,07033 → on monte à 1,071 (et non 1,070).

**2. Révision HT :**

```
Révision HT = Travaux HT × (Cn - 1)
Révision HT = 38 500 × (1,071 - 1)
Révision HT = 38 500 × 0,071
Révision HT = 2 733,50 €
```

**3. Montant total facturé :**

```
Montant HT = Travaux HT + Révision HT
Montant HT = 38 500 + 2 733,50
Montant HT = 41 233,50 € HT
```

---

## EXERCICE 2 — Révision avec décalage temporel (type sujet 2024-2025)

### Énoncé
Un marché de charpente métallique prévoit dans le CCAP :
> *"La révision est calculée avec l'indice du mois M-3 (3 mois avant le mois des travaux)"*

Données :
| Mois | Indice BT50 |
|---|---|
| M0 (offre) | 108,2 |
| M-3 (3 mois avant M) | 115,8 |
| M (mois travaux) | 118,3 |

Travaux HT du mois M = **52 000 € HT**

**Questions :**
1. Quel indice doit-on utiliser pour Im ? Pourquoi ?
2. Calculer Cn et la révision HT.

---

### CORRIGÉ

**1. Indice à utiliser :**

```
Le CCAP précise Im-3 → on utilise l'indice du mois M-3, soit BT50 = 115,8
(et non l'indice du mois M = 118,3)

Raison : les indices BT sont publiés avec retard.
Le CCAP anticipe ce délai en imposant un décalage de 3 mois.
```

**2. Calcul Cn et révision :**

```
Cn = 0,15 + 0,85 × (Im-3 / I0)
Cn = 0,15 + 0,85 × (115,8 / 108,2)
Cn = 0,15 + 0,85 × 1,07024...
Cn = 0,15 + 0,90970...
Cn = 1,05970...
→ Cn = 1,060  (millième supérieur)

Révision HT = 52 000 × (1,060 - 1)
Révision HT = 52 000 × 0,060
Révision HT = 3 120,00 € HT
```

---

## EXERCICE 3 — Formule à deux indices (type sujet 2023 — lot toiture)

### Énoncé
Un marché de couverture-étanchéité a la formule de révision suivante :
```
Cn = 0,15 + 0,85 × [(Im_BT30 / I0_BT30) × 10% + (Im_BT34 / I0_BT34) × 90%]
```

Données :
| Indice | I0 | Im |
|---|---|---|
| BT30 (couverture) | 105,1 | 109,4 |
| BT34 (étanchéité) | 118,7 | 127,9 |

Travaux HT du mois = **76 200 € HT**

**Questions :**
1. Calculer Cn (arrondi au millième supérieur).
2. Calculer la révision HT.

---

### CORRIGÉ

**1. Coefficient Cn :**

```
Terme BT30 : (109,4 / 105,1) × 10% = 1,04091 × 0,10 = 0,10409
Terme BT34 : (127,9 / 118,7) × 90% = 1,07750 × 0,90 = 0,96975

Somme pondérée = 0,10409 + 0,96975 = 1,07384

Cn = 0,15 + 0,85 × 1,07384
Cn = 0,15 + 0,91276
Cn = 1,06276...
→ Cn = 1,063  (millième supérieur)
```

**2. Révision HT :**

```
Révision HT = 76 200 × (1,063 - 1)
Révision HT = 76 200 × 0,063
Révision HT = 4 800,60 € HT
```

---

## EXERCICE 4 — Révision sur plusieurs mois consécutifs (type sujet 2025)

### Énoncé
Un marché de maçonnerie : montant initial = **180 000 € HT**, durée = 6 mois.
Formule : `Cn = 0,15 + 0,85 × (Im / I0)`, indice BT01.

| Mois | Indice BT01 | % avancement cumulé | % avancement mois |
|---|---|---|---|
| I0 (offre) | 110,0 | — | — |
| M1 | 111,2 | 10% | 10% |
| M2 | 112,5 | 25% | 15% |
| M3 | 114,8 | 45% | 20% |
| M4 | 116,0 | 70% | 25% |
| M5 | 117,3 | 90% | 20% |
| M6 | 118,1 | 100% | 10% |

**Questions :**
1. Calculer Cn pour chaque mois.
2. Calculer les travaux HT et la révision HT pour chaque mois.
3. Calculer la révision totale HT sur le marché.

---

### CORRIGÉ

**Travaux HT du mois = Montant marché × % avancement mois**

| Mois | Im | Im/I0 | Cn (millième sup.) | Travaux HT | Révision HT |
|---|---|---|---|---|---|
| M1 | 111,2 | 1,01091 | 1,010 | 18 000 | 180,00 |
| M2 | 112,5 | 1,02273 | 1,020 | 27 000 | 540,00 |
| M3 | 114,8 | 1,04364 | 1,038 | 36 000 | 1 368,00 |
| M4 | 116,0 | 1,05455 | 1,047 | 45 000 | 2 115,00 |
| M5 | 117,3 | 1,06636 | 1,057 | 36 000 | 2 052,00 |
| M6 | 118,1 | 1,07364 | 1,074 | 18 000 | 1 332,00 |

> Détail M3 : Cn = 0,15 + 0,85 × 1,04364 = 0,15 + 0,88709 = 1,03709 → 1,038

```
Révision totale = 180 + 540 + 1 368 + 2 115 + 2 052 + 1 332
Révision totale = 7 587,00 € HT
```

---

## POINTS DE VIGILANCE

- **Cn arrondi au millième SUPÉRIEUR** : jamais inférieur, jamais au plus proche
- **Vérifier quel indice utiliser** : Im, Im-3, ou Im-4 selon le CCAP
- **Révision = Travaux × (Cn - 1)** et non Travaux × Cn
- **La part fixe de 0,15 (15%) n'est pas révisée** (protège le maître d'ouvrage)
- Ne jamais appliquer de révision sur l'avance forfaitaire
