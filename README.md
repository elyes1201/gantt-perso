# Gantt PQR – visuel personnalisé Power BI

Diagramme de Gantt **Site > Produit > Réunions (Livrable CPV)** avec progression et jalons.
Construit avec `pbiviz` (powerbi-visuals-tools), TypeScript et des modules D3 minimaux
(`d3-selection`, `d3-scale`, `d3-time`, `d3-time-format`).

- Livrable : `dist/ganttPQR4F1C2B7E9A3D45C8B6E0D1A2F3C4B5D6.1.0.0.0.pbiviz`
- Données d'exemple (fictives) : `sample/donnees_exemple.csv`

## Fonctionnalités

- Hiérarchie **Site > Produit > Réunion** (une barre par ligne), sites et produits repliables
  (chevron ▸/▾, double-clic, boutons ⊞ « Tout déplier » / ⊟ « Tout replier »).
  Mode « Compact » disponible (réunions sur la ligne Produit).
- Bouton **⚙** : panneau de paramétrage dans le visuel (mode d'affichage, hiérarchie de dates,
  hauteurs, symbole/taille/couleur du jalon, labels, échelle de progression, couleurs par site).
  Les choix sont enregistrés **dans le rapport** (`persistProperties`), jamais dans le navigateur.
  En mode lecture du service, ils ne s'appliquent qu'à la session en cours.
- Hiérarchie de dates : Automatique, Année > Trimestre, Année > Mois (défaut), Trimestre > Mois,
  Mois > Semaine (lundi), Mois > Jour, Semaine > Jour.
- Symboles de jalon : losange (défaut), cercle, carré, triangle, étoile, drapeau.

## Construire

```bash
npm ci
npm run package      # = pbiviz package --no-minify (bundle lisible, non minifié)
```

## Sources

| Fichier | Rôle |
|---|---|
| `capabilities.json` | Rôles de données (table, sans agrégation, fenêtre 30000), objets du volet Format, `"privileges": []` |
| `src/dataModel.ts` | Parsing / nettoyage / tri / regroupement (module pur, sans DOM) |
| `src/settings.ts` | Modèle du volet Format (formattingModel) |
| `src/settingsPanel.ts` | Panneau ⚙ intégré (DOM via createElement / textContent) |
| `src/timeLevels.ts` | Niveaux de la timeline (locale fr locale) et symboles de jalon |
| `src/visual.ts` | Rendu D3, défilement, zoom, sélection, infobulles, menu contextuel |
| `style/visual.less` | Styles (polices système uniquement) |

## Sécurité

- `"privileges": []` : aucun accès réseau, stockage ou fenêtre externe.
- Aucun `fetch`, `XMLHttpRequest`, `WebSocket`, police ou script distant ; aucune télémétrie,
  aucun `console.log`, aucun `localStorage` / cookie / IndexedDB.
- Tout texte issu des données passe par `textContent` ou `d3 .text()` ; les infobulles passent par
  `ITooltipService` (texte brut). Note : la méthode `.html()` de `d3-selection` est présente dans le
  bundle (code de la bibliothèque) mais n'est jamais appelée.
- Dépendances épinglées (versions exactes) et `package-lock.json` commité.
- `npm audit` : 0 faille en dépendances d'exécution ; 0 haute / critique au total. 11 failles
  **modérées** subsistent uniquement dans l'outillage de développement (`powerbi-visuals-tools` →
  eslint/minimatch, webpack-dev-server/sockjs/uuid) ; non embarquées dans le `.pbiviz`, pas de
  correctif sans régression majeure de l'outil.
