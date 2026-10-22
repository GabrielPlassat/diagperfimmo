# Extraction Diag Perf'Immo → Excel (API Albert)

Application web (Streamlit) qui lit des rapports d'audit **Diag Perf'Immo** (PDF) et remplit
automatiquement le modèle Excel **« modèle indicateurs perfimmo - enrichis »** grâce à
l'IA souveraine **Albert** (DINUM).

Chaque valeur extraite est :
- **colorée** selon le niveau de confiance (vert / jaune / rouge / gris / bleu = validé à la main) ;
- **justifiée** par un commentaire Excel : page du PDF, extrait cité, calcul éventuel ;
- **contrôlée** automatiquement : la valeur doit se retrouver dans le texte du PDF, et des règles
  de cohérence sont vérifiées (SIRET à 14 chiffres, Crelat = 60/50/40 % de Cref, somme des énergies ≈ Cref…).

Avant de télécharger, vous relisez et corrigez les valeurs dans l'application, et vous pouvez
afficher la page du PDF correspondante.

---

## Mise en ligne (sans rien installer sur votre ordinateur)

### 1. Obtenir une clé API Albert
Connectez-vous sur l'espace Albert API (accès ProConnect pour les agents publics), créez une
clé et **copiez-la tout de suite** : elle ne s'affiche qu'une fois.
Documentation : https://guides.ia.numerique.gouv.fr/albert-api

### 2. Créer le dépôt GitHub
1. Sur github.com : **New repository** → nom `perfimmo-albert` → *Private* conseillé → **Create**.
2. Cliquez **uploading an existing file** et glissez **tout le contenu** du dossier
   (y compris les dossiers `modele/` et `.streamlit/`). → **Commit changes**.

> Le dossier `.streamlit` est caché sur certains ordinateurs. S'il ne part pas, créez-le dans GitHub :
> *Add file → Create new file* → nom `.streamlit/config.toml` → collez le contenu du fichier.

### 3. Déployer sur Streamlit Community Cloud
1. Allez sur https://share.streamlit.io → **Create app** → *Deploy a public app from GitHub*.
2. Repository : `votre-compte/perfimmo-albert`, Branch : `main`, Main file : `app.py`.
3. **Advanced settings → Secrets**, collez :
   ```toml
   ALBERT_API_KEY = "votre-cle-albert"
   ```
4. **Deploy**. Premier démarrage : 2–3 minutes.

Sans secret, l'application demande la clé dans la barre latérale (utile pour partager l'outil
sans partager votre clé).

---

## Utilisation
1. Déposez un ou plusieurs PDF → **Lancer l'extraction** (≈ 1 à 3 min par rapport).
2. Onglet par rapport : activez *« Afficher seulement les valeurs à vérifier »*, corrigez la
   colonne **Valeur** (la cellule passe en bleu « validé »), ouvrez *« Voir une page du PDF »* pour contrôler.
3. **Télécharger l'Excel rempli** : une colonne par rapport dans la feuille du modèle, plus une feuille
   *Détail extraction* (format long : une ligne par rapport × indicateur, pratique pour consolider
   des dizaines de rapports) et une feuille *Légende*.

### Réglages (barre latérale)
| Réglage | Conseil |
|---|---|
| Modèle | `gpt-oss-120b` (alias `openweight-large`) par défaut. Essayez `mistral-small-…` ou `gemma-4-31b-it` pour comparer. |
| Texte envoyé | *Rapport complet* = le plus fiable (~22 000 tokens × 7 appels par rapport). *Pages ciblées* = moitié moins de tokens, si vous atteignez vos quotas. |
| Requêtes simultanées | Passez à 1 en cas d'erreur 429 (quota). Les relances sont automatiques. |
| Modèle Excel | Facultatif : tout fichier dont la colonne A liste les indicateurs. |

---

## Améliorer l'extraction sans coder
- **`consignes.csv`** (modifiable directement dans GitHub, séparateur `;`) : une consigne par
  indicateur (« où chercher », « quelle unité », « quel synonyme »). La colonne `indicateur` doit
  reprendre exactement le libellé de la colonne A du modèle. C'est le levier principal pour corriger
  une erreur récurrente.
- **Prompt général** : variable `PROMPT_SYSTEME` dans `perfimmo.py` (glossaire Cref, Cabs, Crelat…).
- **Groupes d'indicateurs** : l'outil envoie 7 requêtes par rapport (identification, référence,
  déperditions, gains rapides, 2030, 2040, 2050), détectées automatiquement à partir des libellés.

## Fichiers
| Fichier | Rôle |
|---|---|
| `app.py` | Interface Streamlit |
| `perfimmo.py` | Lecture PDF, appels Albert, contrôles, écriture Excel |
| `consignes.csv` | Consignes par indicateur |
| `modele/modele_indicateurs_perfimmo_enrichis.xlsx` | Modèle Excel par défaut |
| `requirements.txt` | Bibliothèques Python (installées automatiquement par Streamlit Cloud) |
| `.streamlit/config.toml` | Taille max des PDF (100 Mo), couleur |
| `.streamlit/secrets.toml.example` | Exemple de secret (ne jamais publier la vraie clé) |

## Limites connues
- Les valeurs présentes **uniquement dans des graphiques** (camemberts, histogrammes) ne sont lues que
  si le PDF contient leur texte. Piste : envoyer l'image de ces pages à un modèle vision d'Albert.
- Les tableaux complexes sur plusieurs colonnes (rapport Santé) peuvent être mal reconstruits :
  ces valeurs ressortent en général en jaune ou rouge.
- Les rapports anonymisés (« [adresse] », « NC ») donnent logiquement des cases grises.
- Confidentialité : les rapports sont envoyés à l'API Albert (hébergement souverain) ; Streamlit
  Community Cloud héberge l'application aux États-Unis mais ne conserve pas les fichiers déposés.
