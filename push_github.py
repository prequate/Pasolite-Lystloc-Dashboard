#!/usr/bin/env python3
"""
Push the weekly dashboard update to GitHub, or check that a manual upload landed.

  python push_github.py "Week 28 Sep to 4 Oct 2026"   # commit and push
  python push_github.py --verify                       # compare GitHub with this folder

Repo: https://github.com/prequate/Pasolite-Lystloc-Dashboard (branch main).
Only the repo layout is sent: index.html, build_dashboard.py, README.md,
requirements.txt, vendor/, data/cleaned/. The raw LystLoc files,
inbox/, the runbook and the pipeline scripts stay out of GitHub.

Exit codes: 0 done or nothing to do, 3 no push access (manual upload needed),
4 GitHub does not match after push or on --verify.
"""
import filecmp
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = "https://github.com/prequate/Pasolite-Lystloc-Dashboard.git"
TOP_FILES = ["index.html", "build_dashboard.py", "README.md", "requirements.txt"]
# assets/ (the logo) never changes and is left out of the sync on purpose.
DIRS = ["vendor", os.path.join("data", "cleaned")]
TRAILER = ("\n\nCo-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>\n"
           "Claude-Session: https://claude.ai/code/session_01JwtuwzJAfiaREpYHnn8PLp")


def git(args, cwd, check=True):
    r = subprocess.run(["git"] + args, cwd=cwd, capture_output=True, text=True, timeout=180)
    if check and r.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} failed:\n{r.stdout}\n{r.stderr}")
    return r


def clone():
    d = tempfile.mkdtemp(prefix="pasolite_gh_")
    git(["clone", "-q", "--depth", "1", REPO, d], cwd=HERE)
    return d


def compare(clone_dir):
    """Return a list of differences between GitHub (clone) and this folder."""
    diffs = []
    for f in TOP_FILES:
        a, b = os.path.join(HERE, f), os.path.join(clone_dir, f)
        if not os.path.exists(b):
            diffs.append(f"missing on GitHub: {f}")
        elif not filecmp.cmp(a, b, shallow=False):
            diffs.append(f"different on GitHub: {f}")
    for d in DIRS:
        for f in sorted(os.listdir(os.path.join(HERE, d))):
            a, b = os.path.join(HERE, d, f), os.path.join(clone_dir, d, f)
            if not os.path.exists(b):
                diffs.append(f"missing on GitHub: {d}/{f}")
            elif not filecmp.cmp(a, b, shallow=False):
                diffs.append(f"different on GitHub: {d}/{f}")
    return diffs


def sync_into(clone_dir):
    for f in TOP_FILES:
        shutil.copy2(os.path.join(HERE, f), os.path.join(clone_dir, f))
    for d in DIRS:
        os.makedirs(os.path.join(clone_dir, d), exist_ok=True)
        for f in os.listdir(os.path.join(HERE, d)):
            shutil.copy2(os.path.join(HERE, d, f), os.path.join(clone_dir, d, f))


def main():
    if "--verify" in sys.argv:
        c = clone()
        diffs = compare(c)
        shutil.rmtree(c, ignore_errors=True)
        if diffs:
            print("GITHUB_MISMATCH\n" + "\n".join(diffs))
            sys.exit(4)
        print("GITHUB_MATCHES")
        return

    msg = " ".join(a for a in sys.argv[1:] if not a.startswith("--")) or "Weekly dashboard update"
    c = clone()
    before = compare(c)
    if not before:
        print("NOTHING_TO_PUSH: GitHub already matches.")
        return
    sync_into(c)
    git(["add", "-A"], c)
    git(["-c", "user.name=Harsh Mundra (Prequate weekly run)", "-c", "user.email=mundra@prequate.one",
         "commit", "-q", "-m", msg + TRAILER], c)
    p = git(["push", "-q", "origin", "HEAD:main"], c, check=False)
    if p.returncode != 0:
        print("NO_PUSH_ACCESS. Upload these by hand (GitHub > Add file > Upload files, keep the folder):")
        print("\n".join(" - " + x.split(": ", 1)[1] for x in before))
        print("\n" + (p.stderr or p.stdout).strip()[-400:])
        shutil.rmtree(c, ignore_errors=True)
        sys.exit(3)
    shutil.rmtree(c, ignore_errors=True)
    c2 = clone()
    after = compare(c2)
    shutil.rmtree(c2, ignore_errors=True)
    if after:
        print("GITHUB_MISMATCH_AFTER_PUSH\n" + "\n".join(after))
        sys.exit(4)
    print("PUSHED_AND_VERIFIED: " + ", ".join(x.split(": ", 1)[1] for x in before))


if __name__ == "__main__":
    main()
