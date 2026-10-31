#!/usr/bin/env python3
"""
gh_bench — Bot de benchmark de librairies GitHub (sans IA, règles simples).

Recherche les dépôts GitHub sur un sujet, collecte des métriques, calcule un
score et produit un rapport Markdown avec des recommandations.

Usage :
    python gh_bench.py "http client" --language python --top 10 --out rapport.md

Un jeton (variable GITHUB_TOKEN) est facultatif mais conseillé :
sans lui, l'API GitHub limite à 60 requêtes par heure.
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone

API = "https://api.github.com"
NOW = datetime.now(timezone.utc)

# --------------------------------------------------------------------------- #
# Règles — modifiez ces seuils pour adapter le bot
# --------------------------------------------------------------------------- #
WEIGHTS = {            # poids de chaque critère dans le score final (/100)
    "popularite": 25,
    "activite": 30,
    "communaute": 20,
    "maintenance": 15,
    "licence": 10,
}
STALE_DAYS = 365            # pas de push depuis X jours -> abandon probable
WARN_DAYS = 180             # pas de push depuis X jours -> ralentissement
MIN_CONTRIBUTORS = 3        # en dessous : dépendance à une seule personne
ISSUES_RATIO_WARN = 0.10    # issues ouvertes / stars au-dessus : support débordé
PERMISSIVE = {"mit", "apache-2.0", "bsd-2-clause", "bsd-3-clause", "isc",
              "mpl-2.0", "unlicense", "0bsd", "zlib"}
COPYLEFT = {"gpl-2.0", "gpl-3.0", "agpl-3.0", "lgpl-2.1", "lgpl-3.0"}


# --------------------------------------------------------------------------- #
# Accès à l'API GitHub
# --------------------------------------------------------------------------- #
class GitHub:
    def __init__(self, token: str | None):
        self.token = token
        self.calls = 0

    def get(self, path: str, params: dict | None = None):
        url = path if path.startswith("http") else API + path
        if params:
            url += "?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "gh-bench-bot",
            "X-GitHub-Api-Version": "2022-11-28",
            **({"Authorization": f"Bearer {self.token}"} if self.token else {}),
        })
        self.calls += 1
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                body = r.read().decode()
                return (json.loads(body) if body else None), dict(r.headers)
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None, dict(e.headers)
            if e.code in (403, 429) and e.headers.get("X-RateLimit-Remaining") == "0":
                reset = int(e.headers.get("X-RateLimit-Reset", "0"))
                when = datetime.fromtimestamp(reset).strftime("%H:%M")
                sys.exit(f"Limite de l'API GitHub atteinte (réinitialisation à {when}). "
                         "Définissez GITHUB_TOKEN pour passer à 5000 requêtes/heure.")
            raise

    def count(self, path: str, params: dict | None = None) -> int:
        """Compte les éléments d'une liste paginée en 1 requête (per_page=1 + en-tête Link)."""
        data, headers = self.get(path, {**(params or {}), "per_page": 1})
        link = headers.get("Link") or headers.get("link") or ""
        m = re.search(r'[?&]page=(\d+)>; rel="last"', link)
        if m:
            return int(m.group(1))
        return len(data) if isinstance(data, list) else 0


# --------------------------------------------------------------------------- #
# Modèle
# --------------------------------------------------------------------------- #
@dataclass
class Repo:
    name: str
    url: str
    description: str
    stars: int
    forks: int
    open_issues: int
    pushed_at: datetime
    created_at: datetime
    license: str | None
    topics: list[str]
    contributors: int | None = None   # None = donnée inaccessible
    commits_90d: int | None = None
    partial: bool = False
    last_release: datetime | None = None
    release_tag: str | None = None
    scores: dict = field(default_factory=dict)
    total: float = 0.0
    alerts: list[str] = field(default_factory=list)
    strengths: list[str] = field(default_factory=list)

    @property
    def days_since_push(self) -> int:
        return (NOW - self.pushed_at).days

    @property
    def age_years(self) -> float:
        return (NOW - self.created_at).days / 365


def parse_dt(s: str | None) -> datetime | None:
    return datetime.fromisoformat(s.replace("Z", "+00:00")) if s else None


# --------------------------------------------------------------------------- #
# Collecte
# --------------------------------------------------------------------------- #
def search(gh: GitHub, query: str, language: str | None, min_stars: int, n: int) -> list[Repo]:
    q = f"{query} archived:false stars:>={min_stars}"
    if language:
        q += f" language:{language}"
    data, _ = gh.get("/search/repositories", {"q": q, "sort": "stars", "order": "desc",
                                              "per_page": min(n, 100)})
    repos = []
    for it in (data or {}).get("items", []):
        if it.get("fork"):
            continue
        repos.append(Repo(
            name=it["full_name"], url=it["html_url"],
            description=(it.get("description") or "").strip(),
            stars=it["stargazers_count"], forks=it["forks_count"],
            open_issues=it["open_issues_count"],
            pushed_at=parse_dt(it["pushed_at"]), created_at=parse_dt(it["created_at"]),
            license=(it.get("license") or {}).get("spdx_id") if it.get("license") else None,
            topics=it.get("topics", []),
        ))
    return repos


def enrich(gh: GitHub, r: Repo) -> None:
    since = (NOW - timedelta(days=90)).strftime("%Y-%m-%dT%H:%M:%SZ")
    def safe(fn):
        try:
            return fn()
        except urllib.error.HTTPError:
            # 403 possible : dépôt trop gros pour lister les contributeurs, accès restreint…
            r.partial = True
            return None
    r.contributors = safe(lambda: gh.count(f"/repos/{r.name}/contributors", {"anon": "1"}))
    r.commits_90d = safe(lambda: gh.count(f"/repos/{r.name}/commits", {"since": since}))
    res = safe(lambda: gh.get(f"/repos/{r.name}/releases/latest"))
    rel = res[0] if res else None
    if rel:
        r.last_release = parse_dt(rel.get("published_at"))
        r.release_tag = rel.get("tag_name")


# --------------------------------------------------------------------------- #
# Notation (chaque critère ramené entre 0 et 1, puis pondéré)
# --------------------------------------------------------------------------- #
def score(repos: list[Repo]) -> None:
    max_log_stars = max(math.log10(r.stars + 1) for r in repos) or 1
    for r in repos:
        s = {}
        s["popularite"] = math.log10(r.stars + 1) / max_log_stars

        push = max(0.0, 1 - r.days_since_push / STALE_DAYS)
        if r.commits_90d is None:
            s["activite"] = push                       # repli : fraîcheur du push seule
        else:
            commits = min(1.0, math.log10(r.commits_90d + 1) / 2)   # 100 commits/90 j = 1
            s["activite"] = 0.5 * push + 0.5 * commits

        if r.contributors is None:                     # repli : forks comme indicateur
            s["communaute"] = min(1.0, math.log10(r.forks + 1) / 3.5)
        else:
            s["communaute"] = min(1.0, math.log10(r.contributors + 1) / 2)  # 100 contributeurs = 1

        if r.last_release:
            days = (NOW - r.last_release).days
            rel = max(0.0, 1 - days / (2 * STALE_DAYS))
        else:
            rel = 0.3   # pas de release formelle : on ne pénalise pas totalement
        ratio = r.open_issues / max(r.stars, 1)
        issues = max(0.0, 1 - ratio / (2 * ISSUES_RATIO_WARN))
        s["maintenance"] = 0.6 * rel + 0.4 * issues

        lic = (r.license or "").lower()
        s["licence"] = 1.0 if lic in PERMISSIVE else 0.6 if lic in COPYLEFT else \
                       0.4 if lic and lic != "noassertion" else 0.0

        r.scores = s
        r.total = round(sum(WEIGHTS[k] * v for k, v in s.items()), 1)
        apply_rules(r)


def apply_rules(r: Repo) -> None:
    lic = (r.license or "").lower()
    # Alertes
    if r.days_since_push > STALE_DAYS:
        r.alerts.append(f"Aucun push depuis {r.days_since_push} jours : projet probablement abandonné.")
    elif r.days_since_push > WARN_DAYS:
        r.alerts.append(f"Activité ralentie (dernier push il y a {r.days_since_push} jours).")
    if r.partial:
        r.alerts.append("Certaines métriques étaient inaccessibles : score partiellement estimé.")
    if r.contributors is not None and r.contributors < MIN_CONTRIBUTORS:
        r.alerts.append(f"Seulement {r.contributors} contributeur(s) : dépend d'une seule personne.")
    if not lic or lic == "noassertion":
        r.alerts.append("Pas de licence claire : usage légalement risqué.")
    elif lic in COPYLEFT:
        r.alerts.append(f"Licence copyleft ({r.license}) : vérifiez la compatibilité avec votre projet.")
    if r.open_issues / max(r.stars, 1) > ISSUES_RATIO_WARN:
        r.alerts.append(f"{r.open_issues} issues/PR ouvertes pour {r.stars} stars : support débordé.")
    if not r.last_release and not r.partial:
        r.alerts.append("Aucune release publiée : versions moins faciles à épingler.")
    # Points forts
    if (r.commits_90d or 0) >= 30:
        r.strengths.append(f"Très actif ({r.commits_90d} commits sur 90 jours).")
    if (r.contributors or 0) >= 50:
        r.strengths.append(f"Large communauté ({r.contributors} contributeurs).")
    if r.last_release and (NOW - r.last_release).days <= 90:
        r.strengths.append(f"Release récente ({r.release_tag}).")
    if r.age_years >= 5 and r.days_since_push <= 60:
        r.strengths.append(f"Mature ({r.age_years:.0f} ans) et toujours maintenu.")
    if lic in PERMISSIVE:
        r.strengths.append(f"Licence permissive ({r.license}).")


def verdict(r: Repo) -> str:
    if r.total >= 75 and not any("abandonné" in a or "licence claire" in a for a in r.alerts):
        return "✅ Recommandé"
    if r.total >= 55:
        return "🟡 Acceptable"
    return "🔴 À éviter"


# --------------------------------------------------------------------------- #
# Propositions globales
# --------------------------------------------------------------------------- #
def proposals(ranked: list[Repo]) -> list[str]:
    out = []
    best = ranked[0]
    out.append(f"**Choix par défaut : [{best.name}]({best.url})** — meilleur score global "
               f"({best.total}/100).")

    safe = [r for r in ranked if verdict(r) == "✅ Recommandé"]
    if len(safe) > 1:
        alt = safe[1]
        out.append(f"**Alternative solide : [{alt.name}]({alt.url})** ({alt.total}/100), "
                   "à garder en plan B.")

    most_active = max(ranked, key=lambda r: r.commits_90d or 0)
    if most_active is not best and (most_active.commits_90d or 0) > 0:
        out.append(f"**Le plus dynamique : [{most_active.name}]({most_active.url})** "
                   f"({most_active.commits_90d} commits en 90 jours), intéressant si vous avez "
                   "besoin de fonctionnalités récentes.")

    rising = [r for r in ranked if r.age_years < 2 and r.scores["activite"] > 0.6]
    if rising:
        r = max(rising, key=lambda x: x.stars)
        out.append(f"**Étoile montante : [{r.name}]({r.url})** — moins de 2 ans, "
                   f"{r.stars} stars, à surveiller mais à adopter avec prudence.")

    popular_risky = [r for r in ranked if r.scores["popularite"] > 0.8 and verdict(r) == "🔴 À éviter"]
    for r in popular_risky:
        out.append(f"**Piège : [{r.name}]({r.url})** est très populaire mais mal noté "
                   f"({r.total}/100) : {r.alerts[0] if r.alerts else 'voir détails'}")

    if not safe:
        out.append("Aucun dépôt n'atteint le seuil « Recommandé » : élargissez la recherche "
                   "(autre mot-clé, langage, `--min-stars` plus bas) ou évaluez manuellement.")
    return out


# --------------------------------------------------------------------------- #
# Rapport
# --------------------------------------------------------------------------- #
def na(v) -> str:
    return "n/d" if v is None else str(v)


def bar(v: float) -> str:
    n = round(v * 5)
    return "█" * n + "░" * (5 - n)


def report(query: str, language: str | None, ranked: list[Repo]) -> str:
    L = [f"# Benchmark GitHub : « {query} »" + (f" ({language})" if language else ""),
         "", f"_Généré le {NOW.strftime('%d/%m/%Y à %H:%M UTC')} — {len(ranked)} dépôts analysés._",
         "", "## Propositions", ""]
    L += [f"{i}. {p}" for i, p in enumerate(proposals(ranked), 1)]
    L += ["", "## Classement", "",
          "| # | Dépôt | Score | Verdict | ⭐ Stars | Contrib. | Commits 90 j | Dernier push | Licence |",
          "|---|---|---|---|---|---|---|---|---|"]
    for i, r in enumerate(ranked, 1):
        L.append(f"| {i} | [{r.name}]({r.url}) | **{r.total}** | {verdict(r)} | {r.stars:,} | "
                 f"{na(r.contributors)} | {na(r.commits_90d)} | il y a {r.days_since_push} j | "
                 f"{r.license or '—'} |".replace(",", " "))
    L += ["", "## Détails", ""]
    for i, r in enumerate(ranked, 1):
        L += [f"### {i}. {r.name} — {r.total}/100 {verdict(r)}", ""]
        if r.description:
            L += [f"> {r.description}", ""]
        L += ["| Critère | Note | Poids |", "|---|---|---|"]
        L += [f"| {k.capitalize()} | `{bar(v)}` {v:.2f} | {WEIGHTS[k]} |" for k, v in r.scores.items()]
        L.append("")
        if r.strengths:
            L += ["**Points forts**"] + [f"- {s}" for s in r.strengths] + [""]
        if r.alerts:
            L += ["**Points de vigilance**"] + [f"- ⚠️ {a}" for a in r.alerts] + [""]
    L += ["---", "### Méthode", "",
          "Score sur 100 = " + " + ".join(f"{k} ({w})" for k, w in WEIGHTS.items()) + ".",
          "- **Popularité** : stars (échelle logarithmique, relative au meilleur du lot).",
          "- **Activité** : fraîcheur du dernier push + commits sur 90 jours.",
          "- **Communauté** : nombre de contributeurs.",
          "- **Maintenance** : fraîcheur de la dernière release + ratio issues ouvertes / stars.",
          "- Si une métrique est inaccessible (n/d), l'activité se base sur le dernier push "
          "et la communauté sur les forks.",
          "- **Licence** : permissive = 1, copyleft = 0,6, autre = 0,4, absente = 0.",
          "", "Seuils et poids modifiables en tête de `gh_bench.py`."]
    return "\n".join(L) + "\n"


# --------------------------------------------------------------------------- #
def main() -> None:
    p = argparse.ArgumentParser(description="Benchmark de librairies GitHub avec recommandations.")
    p.add_argument("query", help='Sujet recherché, ex. "http client"')
    p.add_argument("-l", "--language", help="Langage, ex. python, javascript, go")
    p.add_argument("-n", "--top", type=int, default=10, help="Nombre de dépôts analysés (défaut 10)")
    p.add_argument("--min-stars", type=int, default=100, help="Stars minimum (défaut 100)")
    p.add_argument("-o", "--out", default="rapport.md", help="Fichier de sortie Markdown")
    p.add_argument("--json", help="Exporter aussi les données brutes en JSON")
    a = p.parse_args()

    gh = GitHub(os.environ.get("GITHUB_TOKEN"))
    print(f"🔍 Recherche « {a.query} »…", file=sys.stderr)
    repos = search(gh, a.query, a.language, a.min_stars, a.top)
    if not repos:
        sys.exit("Aucun dépôt trouvé. Essayez un autre mot-clé ou --min-stars plus bas.")
    for i, r in enumerate(repos, 1):
        print(f"  [{i}/{len(repos)}] {r.name}", file=sys.stderr)
        enrich(gh, r)
    score(repos)
    ranked = sorted(repos, key=lambda r: r.total, reverse=True)

    with open(a.out, "w", encoding="utf-8") as f:
        f.write(report(a.query, a.language, ranked))
    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump({
                "query": a.query, "language": a.language, "generated": NOW.isoformat(),
                "weights": WEIGHTS,
                "proposals": proposals(ranked),
                "repos": [{**{k: v for k, v in vars(r).items()
                              if k not in ("pushed_at", "created_at", "last_release")},
                           "days_since_push": r.days_since_push,
                           "age_years": round(r.age_years, 1),
                           "pushed_at": r.pushed_at.isoformat(),
                           "last_release": r.last_release.isoformat() if r.last_release else None,
                           "verdict": verdict(r)} for r in ranked],
            }, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Rapport écrit dans {a.out} ({gh.calls} appels API)", file=sys.stderr)
    for i, r in enumerate(ranked[:3], 1):
        print(f"  {i}. {r.name} — {r.total}/100 {verdict(r)}", file=sys.stderr)


if __name__ == "__main__":
    main()
