# gh-bench-bot

Bot qui compare les librairies GitHub sur un sujet et propose lesquelles utiliser.
Il fonctionne avec des règles simples, sans IA ni dépendance (Python 3.10+).

## Utilisation locale

```bash
export GITHUB_TOKEN=ghp_xxx   # facultatif, mais 60 requêtes/heure sans
python gh_bench.py "http client" -l python -n 10 -o rapport.md
```

Options : `-l` langage, `-n` nombre de dépôts, `--min-stars`, `-o` fichier Markdown, `--json` export brut.

## Automatisation

Le workflow `.github/workflows/gh-bench.yml` tourne chaque lundi :
1. il publie le rapport dans une **issue**, avec un lien de téléchargement ;
2. il génère une **présentation PowerPoint** et l'enregistre dans le dossier `rapports/`.

Pour générer la présentation à la main :

```bash
pip install python-pptx
python gh_bench.py "http client" -l python --json data.json
python gh_slides.py data.json -o rapport.pptx
```
Pour un autre sujet : onglet **Actions → Benchmark de librairies → Run workflow**.

## Notation (/100)

| Critère | Poids | Mesure |
|---|---|---|
| Popularité | 25 | stars (log) |
| Activité | 30 | dernier push + commits 90 j |
| Communauté | 20 | contributeurs |
| Maintenance | 15 | dernière release + issues ouvertes / stars |
| Licence | 10 | permissive > copyleft > autre > aucune |

Verdict : ≥ 75 recommandé, ≥ 55 acceptable, sinon à éviter. Seuils et poids modifiables en tête de `gh_bench.py`.
