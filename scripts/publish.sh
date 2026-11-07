#!/usr/bin/env bash
#
# Wait for GitHub auth, then create the repo and push.
#
#   ./scripts/publish.sh
#
# Environment:
#   REPO_NAME       repo name (default: govt-comment-resolution)
#   REPO_VISIBILITY  public or private (default: public)
#   GIT_AUTHOR_NAME   commit author name
#   GIT_AUTHOR_EMAIL  commit author email
#
# The script polls for auth, then runs the publish steps. It is safe to
# re-run: if the repo already exists, it pushes to it instead of failing.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

REPO_NAME="${REPO_NAME:-govt-comment-resolution}"
REPO_VISIBILITY="${REPO_VISIBILITY:-public}"
TIMEOUT_SECONDS="${TIMEOUT_SECONDS:-900}"

echo "waiting for GitHub auth (up to ${TIMEOUT_SECONDS}s)"
deadline=$(( $(date +%s) + TIMEOUT_SECONDS ))
while ! gh auth status >/dev/null 2>&1; do
	if [ "$(date +%s)" -ge "$deadline" ]; then
		echo "timed out waiting for auth" >&2
		exit 1
	fi
	sleep 5
done
echo "authenticated as $(gh api user --jq .login)"

if [ -n "${GIT_AUTHOR_NAME:-}" ]; then
	git config user.name "$GIT_AUTHOR_NAME"
fi
if [ -n "${GIT_AUTHOR_EMAIL:-}" ]; then
	git config user.email "$GIT_AUTHOR_EMAIL"
fi

echo
echo "checking the working tree"
if [ -n "$(git status --porcelain)" ]; then
	echo "uncommitted changes present, committing them" >&2
	git add -A
	git commit -q -m "chore: save work in progress"
fi

echo
echo "creating the repo"
if gh repo view "$REPO_NAME" >/dev/null 2>&1; then
	echo "repo $REPO_NAME already exists, reusing it"
else
	gh repo create "$REPO_NAME" \
		--public --description "Resolve review comments left by a government department on a Word document" \
		--source . --remote origin
fi

if ! git remote get-url origin >/dev/null 2>&1; then
	git remote add origin "https://github.com/$(gh api user --jq .login)/${REPO_NAME}.git"
fi

echo
echo "pushing"
git push -u origin main

echo
echo "done: https://github.com/$(gh api user --jq .login)/${REPO_NAME}"
