# EXERCICES CORRIGÉS — ANALYSE ET NOTATION DES OFFRES (E52)
> Basés sur les sujets BTS MEC 2023-2025 | Compétence C16-1 | Savoir S7

---

## RAPPEL DES FORMULES

```
Note prix  = Note_max × (Prix_min / Prix_i)
Note délai = Note_max × (Délai_min / Délai_i)

  Prix_min  = offre la moins chère APRÈS négociation
  Délai_min = délai le plus court proposé
  Prix_i    = montant de l'entreprise i

Écart (%) = (Offre - Estimation) / Estimation × 100
  Positif = dépassement / Négatif = économie

Montant après négociation = Offre de base - Moins-values + Plus-values - Remise

Score total = Note prix + Note valeur technique + Note délai + ...
```

---

## EXERCICE 1 — Analyse comparative d'offres lot gros œuvre (type sujet 2023)

### Énoncé
Marché d'une école maternelle. Estimation maître d'ouvrage : **248 500 € HT**

Critères de sélection :
- Prix : **50 points** sur 100
- Valeur technique (méthodologie + références) : **35 points** sur 100
- Délai : **15 points** sur 100

Offres reçues :

| Entreprise | Offre HT | Négociation | Délai (semaines) | Note valeur tech. |
|---|---|---|---|---|
| SOGEBA | 261 200 € | -3 500 € | 14 | 28 |
| BATIPLUS | 244 800 € | +1 200 € | 16 | 31 |
| CMEC | 252 000 € | -8 000 € | 12 | 27 |
| LABAT TP | 238 500 € | +2 000 € | 18 | 33 |

**Questions :**
1. Calculer le montant final de chaque offre après négociation.
2. Calculer l'écart en % entre chaque offre finale et l'estimation.
3. Calculer la note prix de chaque entreprise.
4. Calculer la note délai de chaque entreprise.
5. Calculer le score total et classer les entreprises.

---

### CORRIGÉ

**1. Montants après négociation :**

```
SOGEBA   : 261 200 - 3 500 = 257 700 € HT
BATIPLUS : 244 800 + 1 200 = 246 000 € HT
CMEC     : 252 000 - 8 000 = 244 000 € HT  ← moins chère
LABAT TP : 238 500 + 2 000 = 240 500 € HT
```

> **Attention** : Prix_min = 244 000 € (CMEC, après négociation)

**2. Écarts par rapport à l'estimation (248 500 €) :**

```
SOGEBA   : (257 700 - 248 500) / 248 500 × 100 = +3,70%  (dépassement)
BATIPLUS : (246 000 - 248 500) / 248 500 × 100 = -1,00%  (économie)
CMEC     : (244 000 - 248 500) / 248 500 × 100 = -1,81%  (économie)
LABAT TP : (240 500 - 248 500) / 248 500 × 100 = -3,22%  (économie)
```

**3. Notes prix (sur 50 points) :**

```
Note prix = 50 × (Prix_min / Prix_i) = 50 × (244 000 / Prix_i)

SOGEBA   : 50 × (244 000 / 257 700) = 50 × 0,9468 = 47,34 pts
BATIPLUS : 50 × (244 000 / 246 000) = 50 × 0,9919 = 49,59 pts
CMEC     : 50 × (244 000 / 244 000) = 50 × 1,0000 = 50,00 pts  ← max
LABAT TP : 50 × (244 000 / 240 500) = 50 × 1,0146 → plafonné à 50,00 pts
```

> **Remarque LABAT TP** : Le résultat donne > 50 car son prix est inférieur à Prix_min.
> Cela arrive quand l'offre la moins chère n'est pas celle retenue comme Prix_min
> (ici, LABAT 240 500 < CMEC 244 000). En pratique, **Prix_min = le moins cher = 240 500 €**.
> **Correction :**

```
Prix_min réel = 240 500 € (LABAT TP)
SOGEBA   : 50 × (240 500 / 257 700) = 50 × 0,9332 = 46,66 pts
BATIPLUS : 50 × (240 500 / 246 000) = 50 × 0,9776 = 48,88 pts
CMEC     : 50 × (240 500 / 244 000) = 50 × 0,9857 = 49,29 pts
LABAT TP : 50 × (240 500 / 240 500) = 50 × 1,0000 = 50,00 pts
```

**4. Notes délai (sur 15 points) :**

```
Délai_min = 12 semaines (CMEC)
Note délai = 15 × (12 / Délai_i)

SOGEBA   : 15 × (12 / 14) = 15 × 0,857 = 12,86 pts
BATIPLUS : 15 × (12 / 16) = 15 × 0,750 = 11,25 pts
CMEC     : 15 × (12 / 12) = 15 × 1,000 = 15,00 pts
LABAT TP : 15 × (12 / 18) = 15 × 0,667 = 10,00 pts
```

**5. Score total et classement :**

| Entreprise | Note prix | Note tech. | Note délai | **Total** | **Rang** |
|---|---|---|---|---|---|
| SOGEBA | 46,66 | 28 | 12,86 | **87,52** | 4 |
| BATIPLUS | 48,88 | 31 | 11,25 | **91,13** | 2 |
| CMEC | 49,29 | 27 | 15,00 | **91,29** | 1 |
| LABAT TP | 50,00 | 33 | 10,00 | **93,00** | — |

> **Attention** : LABAT TP a le meilleur prix et la meilleure valeur technique,
> mais son délai le plus long pénalise. Résultat final à recalculer avec Prix_min = 240 500.

**Classement final :**
1. **LABAT TP** : 50,00 + 33 + 10,00 = **93,00 pts** → RETENU
2. CMEC : 49,29 + 27 + 15,00 = 91,29 pts
3. BATIPLUS : 48,88 + 31 + 11,25 = 91,13 pts
4. SOGEBA : 46,66 + 28 + 12,86 = 87,52 pts

---

## EXERCICE 2 — Négociation et actualisation (type sujet 2025)

### Énoncé
Lot charpente bois. Estimation : **89 000 € HT**. Note max prix = 40 pts.

| Entreprise | Offre initiale HT | Moins-values | Plus-values | Remise |
|---|---|---|---|---|
| CHARPBOIS | 94 500 € | 1 200 € | 800 € | 1,5% |
| LIGNOBAT | 87 300 € | 0 € | 1 500 € | 2,0% |
| STRUCTBOIS | 91 000 € | 2 500 € | 0 € | 1,0% |

**Questions :**
1. Calculer le montant final HT de chaque offre.
2. Identifier Prix_min.
3. Calculer la note prix de chaque entreprise.

---

### CORRIGÉ

**1. Montants finaux :**

```
CHARPBOIS  : Base = 94 500 - 1 200 + 800 = 94 100 €
             Remise = 94 100 × 1,5% = 1 411,50 €
             Final = 94 100 - 1 411,50 = 92 688,50 € HT

LIGNOBAT   : Base = 87 300 + 1 500 = 88 800 €
             Remise = 88 800 × 2,0% = 1 776,00 €
             Final = 88 800 - 1 776 = 87 024,00 € HT

STRUCTBOIS : Base = 91 000 - 2 500 = 88 500 €
             Remise = 88 500 × 1,0% = 885,00 €
             Final = 88 500 - 885 = 87 615,00 € HT
```

**2. Prix_min = 87 024 € (LIGNOBAT)**

**3. Notes prix (sur 40 pts) :**

```
CHARPBOIS  : 40 × (87 024 / 92 688,50) = 40 × 0,9389 = 37,56 pts
LIGNOBAT   : 40 × (87 024 / 87 024,00) = 40 × 1,0000 = 40,00 pts
STRUCTBOIS : 40 × (87 024 / 87 615,00) = 40 × 0,9933 = 39,73 pts
```

---

## POINTS DE VIGILANCE

- Toujours utiliser le montant **APRÈS négociation** pour calculer Prix_min et les notes prix
- Si le sujet précise une remise **en %**, l'appliquer sur le montant après moins/plus-values
- La note ne peut **pas dépasser Note_max** (plafonner à 50 si résultat > 50)
- Les notes techniques et délai sont données par le jury ou calculées selon les barèmes
- Lire attentivement les **critères et leurs pondérations** dans chaque sujet
