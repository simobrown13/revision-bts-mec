# EXERCICES CORRIGÉS — THERMIQUE (E4)
> Basés sur les sujets BTS MEC 2023-2025 | Compétence C2 | Savoir S4 + S5

---

## RAPPEL DES FORMULES

```
R = e / λ             résistance thermique d'une couche [m².K/W]
RT = Rsi + ΣR + Rse   résistance totale de la paroi
U = 1 / RT            coefficient de transmission [W/m².K]
a = λ / (ρ × Cp)      diffusivité thermique [m²/s]
φ = 0,023 × e / √a    déphasage thermique [heures]

  Rsi = 0,13 m².K/W  (paroi verticale, flux horizontal)
  Rse = 0,04 m².K/W
  Exigence déphasage : φ > 12 heures
```

---

## EXERCICE 1 — Résistance thermique d'une paroi simple

### Énoncé
Une paroi extérieure est composée des couches suivantes (de l'intérieur vers l'extérieur) :

| Couche | Épaisseur e (m) | Conductivité λ (W/m.K) |
|---|---|---|
| Plâtre | 0,013 | 0,35 |
| Béton banché | 0,20 | 1,75 |
| Isolant laine de verre | 0,12 | 0,035 |
| Enduit extérieur | 0,015 | 0,87 |

**Questions :**
1. Calculer la résistance thermique R de chaque couche.
2. Calculer la résistance totale RT de la paroi.
3. Calculer le coefficient U de la paroi.
4. La réglementation RE2020 impose U ≤ 0,20 W/m².K pour un mur. La paroi est-elle conforme ?

---

### CORRIGÉ

**1. Résistances de chaque couche :**

```
R_plâtre   = 0,013 / 0,35  = 0,037 m².K/W
R_béton    = 0,20  / 1,75  = 0,114 m².K/W
R_isolant  = 0,12  / 0,035 = 3,429 m².K/W
R_enduit   = 0,015 / 0,87  = 0,017 m².K/W
```

**2. Résistance totale :**

```
RT = Rsi + R_plâtre + R_béton + R_isolant + R_enduit + Rse
RT = 0,13 + 0,037 + 0,114 + 3,429 + 0,017 + 0,04
RT = 3,767 m².K/W
```

**3. Coefficient U :**

```
U = 1 / RT = 1 / 3,767 = 0,27 W/m².K
```

**4. Conformité RE2020 :**

```
U calculé = 0,27 W/m².K
U exigé   = 0,20 W/m².K (maxi)

0,27 > 0,20  →  NON CONFORME
→ Il faut augmenter l'épaisseur de l'isolant.
```

**Épaisseur d'isolant nécessaire pour U = 0,20 :**
```
RT_cible = 1 / 0,20 = 5,00 m².K/W
R_isolant_nécessaire = 5,00 - (0,13 + 0,037 + 0,114 + 0,017 + 0,04)
                     = 5,00 - 0,338 = 4,662 m².K/W
e_isolant = R × λ = 4,662 × 0,035 = 0,163 m  →  soit 16,5 cm minimum
```

---

## EXERCICE 2 — Déphasage thermique (inertie d'été)

### Énoncé
Une toiture-terrasse est constituée d'une dalle béton de :
- Épaisseur : e = 0,22 m
- Masse volumique : ρ = 2 300 kg/m³
- Chaleur spécifique : Cp = 1 000 J/kg.K
- Conductivité thermique : λ = 1,75 W/m.K

**Questions :**
1. Calculer la diffusivité thermique a de la dalle.
2. Calculer le déphasage thermique φ.
3. La toiture est-elle conforme à l'exigence de confort d'été (φ > 12 heures) ?
4. Quelle épaisseur minimale faudrait-il pour atteindre φ = 12 heures ?

---

### CORRIGÉ

**1. Diffusivité thermique :**

```
a = λ / (ρ × Cp)
a = 1,75 / (2 300 × 1 000)
a = 1,75 / 2 300 000
a = 7,609 × 10⁻⁷ m²/s
```

**2. Déphasage thermique :**

```
φ = 0,023 × e / √a
φ = 0,023 × 0,22 / √(7,609 × 10⁻⁷)
φ = 0,00506 / 8,723 × 10⁻⁴
φ = 5,80 heures
```

**3. Conformité :**

```
φ calculé = 5,80 heures
Exigence  = φ > 12 heures

5,80 < 12  →  NON CONFORME
→ La dalle seule est insuffisante pour le confort d'été.
→ Solution : ajouter une couche d'isolant à forte inertie ou augmenter l'épaisseur.
```

**4. Épaisseur minimale pour φ = 12 heures :**

```
12 = 0,023 × e / √(7,609 × 10⁻⁷)
12 = 0,023 × e / 8,723 × 10⁻⁴
e  = 12 × 8,723 × 10⁻⁴ / 0,023
e  = 0,01047 / 0,023
e  = 0,455 m  →  soit 45,5 cm de béton
```

> **Conclusion** : Pour le béton seul, l'épaisseur nécessaire est irréaliste (45 cm). Il faut des **matériaux à forte inertie** (béton lourd, terre cuite, béton de chanvre) ou une **combinaison isolant + masse thermique**.

---

## EXERCICE 3 — Paroi composite : calcul complet (type sujet 2024)

### Énoncé
Mur d'un centre de loisirs CLSH, constitution (int → ext) :

| Couche | e (m) | λ (W/m.K) | ρ (kg/m³) | Cp (J/kg.K) |
|---|---|---|---|---|
| Enduit plâtre | 0,012 | 0,35 | 1 050 | 840 |
| Béton cellulaire | 0,30 | 0,11 | 500 | 1 000 |
| Laine de roche | 0,08 | 0,036 | 50 | 840 |
| Bardage bois | 0,022 | 0,13 | 600 | 1 600 |

**Questions :**
1. Calculer U de la paroi. Est-elle conforme RE2020 (U ≤ 0,20) ?
2. Calculer le déphasage φ pour la couche de béton cellulaire seule.
3. Si φ_béton = 5,2 h et φ_isolant = 2,1 h, peut-on additionner les déphasages ?

---

### CORRIGÉ

**1. Coefficient U :**

```
R_enduit   = 0,012 / 0,35  = 0,034 m².K/W
R_béton C. = 0,30  / 0,11  = 2,727 m².K/W
R_laine    = 0,08  / 0,036 = 2,222 m².K/W
R_bardage  = 0,022 / 0,13  = 0,169 m².K/W

RT = 0,13 + 0,034 + 2,727 + 2,222 + 0,169 + 0,04
RT = 5,322 m².K/W

U = 1 / 5,322 = 0,188 W/m².K

0,188 < 0,20  →  CONFORME RE2020
```

**2. Déphasage couche béton cellulaire :**

```
a = λ / (ρ × Cp) = 0,11 / (500 × 1 000) = 2,2 × 10⁻⁷ m²/s
φ = 0,023 × 0,30 / √(2,2 × 10⁻⁷)
φ = 0,0069 / 4,690 × 10⁻⁴
φ = 14,7 heures  →  CONFORME (> 12 h)
```

**3. Additionner les déphasages :**

```
NON — on ne peut pas additionner les déphasages de chaque couche.
La formule φ = 0,023 × e / √a s'applique à UNE SEULE couche homogène.
Pour une paroi composite, on utilise le déphasage de la couche
la plus inerte (ou on applique la formule à l'ensemble si le
matériau est homogène).
→ Ici, le béton cellulaire seul donne φ = 14,7 h : CONFORME.
```

---

## POINTS DE VIGILANCE

- **e en mètres** (jamais en cm dans les formules)
- **√a** : calculer d'abord a, puis prendre la racine carrée
- **φ > 12 h** : exigence de confort d'été (inertie thermique)
- **U ≤ 0,20 W/m².K** pour les murs en RE2020 (neuf)
- **Rsi = 0,13** pour paroi verticale flux horizontal
- **Rse = 0,04** en conditions normales
