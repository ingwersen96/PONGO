#!/bin/sh -l

python3 /update_readme.py > "${INPUT_PROJECTBASEDIR}"

set -e
sh -c "ls"

sh -c "git config --global user.name '${GITHUB_ACTOR}' \
      && git config --global user.email '${GITHUB_ACTOR}@users.noreply.github.com' \
      && git add -A && git commit -m 'updated readme file' \
      && git push -u origin HEAD"