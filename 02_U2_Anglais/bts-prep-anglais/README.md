# BTS MEC - Prep Anglais E2-A

Application d'entrainement a la comprehension orale en anglais pour le BTS Management Economique de la Construction, session 2026.

## Epreuve E2-A
- **Date** : 27 avril 2026, 13h30
- **Lieu** : Lycee Condorcet, Saint-Priest
- **Duree** : 30 minutes
- **Format** : Ecoute d'un document audio en anglais, compte-rendu en francais

## Contenu
5 textes thematiques BTP, niveau B1 a B2+ :

| # | Titre | Niveau | Theme |
|---|-------|--------|-------|
| 1 | Project Progress Update | B1+ | Suivi de chantier |
| 2 | BIM Implementation Meeting | B2 | Maquette numerique |
| 3 | Site Safety Briefing | B1 | Securite chantier |
| 4 | Cost Estimation Presentation | B2 | Estimation des couts |
| 5 | Sustainability in Construction | B2+ | Developpement durable |

## Fonctionnalites
- Lecteur audio avec controle de vitesse
- Simulation complete de l'epreuve en 5 etapes
- Chronometre pour les phases de notes et redaction
- Correction automatique avec score sur les points cles
- Vocabulaire BTP anglais/francais par theme

## Structure
```
audio/          MP3 generes (voix MBROLA, fallback)
src/            Textes bruts anglais
data.json       Donnees structurees (textes, points cles, vocabulaire)
CLAUDE.md       Instructions pour Claude Code
index.html      Application (a generer)
```

## TODO
- [ ] Generer index.html avec Web Speech API (voix naturelle navigateur)
- [ ] Ajouter plus de textes (appels d'offres, reunion de chantier, etc.)
- [ ] Mode quiz vocabulaire
- [ ] Historique des scores
