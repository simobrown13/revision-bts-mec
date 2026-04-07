# BTS MEC - Prep Anglais E2-A

## Contexte
Application d'entrainement a la comprehension orale anglais pour le BTS Management Economique de la Construction (session 2026). Candidat libre, epreuve E2-A le 27 avril 2026.

## Objectif
Application web (HTML/CSS/JS single page) qui simule l'epreuve E2-A :
1. L'utilisateur choisit un texte parmi 5 thematiques BTP
2. 1ere ecoute audio (MP3 ou voix navigateur) sans prise de notes
3. Pause chronometree pour noter les idees retenues
4. 2eme ecoute pour completer les notes
5. Redaction d'un compte-rendu en francais
6. Correction automatique : comparaison avec les points cles attendus + score
7. Section vocabulaire BTP anglais/francais avec cartes retournables

## Structure du projet
```
bts-prep-anglais/
  CLAUDE.md          # Ce fichier
  README.md          # Doc du projet
  data.json          # Textes, points cles, vocabulaire
  index.html         # Application (a creer)
  audio/             # Fichiers MP3 (voix MBROLA, qualite moyenne)
    text1.mp3 a text5.mp3
  src/               # Textes bruts anglais (pour Web Speech API)
    text1.txt a text5.txt
```

## Contraintes techniques
- **Single page HTML** : tout dans un seul fichier index.html (CSS + JS inline)
- **Audio** : preferer la Web Speech API du navigateur (voix naturelle du systeme) avec fallback sur les MP3 embarques
- **Pas de framework** : vanilla JS uniquement
- **Mobile first** : l'utilisateur est sur telephone
- **Offline capable** : les MP3 peuvent etre embarques en base64 si necessaire
- **Pas de serveur** : fichier statique ouvrable directement

## Audio - Priorite
Le probleme principal est la qualite audio. Les MP3 generes avec espeak-ng/MBROLA sont robotiques.
**Solution preferee** : utiliser `window.speechSynthesis` (Web Speech API) qui utilise les voix systeme du telephone (Siri sur iPhone, Google TTS sur Android). Bien plus naturel.
- Precharger les voix avec `speechSynthesis.getVoices()`
- Selectionner une voix en-GB ou en-US
- Vitesse configurable (0.7 / 0.85 / 1.0)
- Fallback sur les MP3 si speechSynthesis indisponible

## Donnees
Voir `data.json` pour les 5 textes :
1. Project Progress Update (B1+) - Suivi de chantier
2. BIM Implementation Meeting (B2) - Maquette numerique
3. Site Safety Briefing (B1) - Securite chantier
4. Cost Estimation Presentation (B2) - Estimation couts
5. Sustainability in Construction (B2+) - Developpement durable

## Design
- Theme sombre (fond #0f1117, cartes #1a1d27)
- Accent dore #e09f3e, vert #2a9d8f, rouge #e76f51
- Font system-ui
- Stepper visuel pour les etapes
- Chronometre pour pause et redaction
