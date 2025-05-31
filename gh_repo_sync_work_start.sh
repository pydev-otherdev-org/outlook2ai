#!/bin/bash
# sync_start.sh
# Fetches and merges the latest changes from the remote 'origin'
# for the current branch. Run this from within your project repository.
set -o pipefail

# Define colors for output
readonly GREEN='\033[0;32m'
readonly RED='\033[0;31m'
readonly YELLOW='\033[0;33m'
readonly NC='\033[0m' # No Color

echo -e "${YELLOW}--- Starting Sync: Fetching updates from GitHub ---${NC}"

# 1. Check if inside a git repository
if ! git rev-parse --is-inside-work-tree > /dev/null 2>&1; then
  echo -e "${RED}Error: Not inside a git repository. Please 'cd' into your project directory.${NC}" >&2
  exit 1
fi
echo "Verified: Inside a git repository ($(pwd))"

# 2. Check if 'origin' remote exists
if ! git remote | grep -q '^origin$'; then
  echo -e "${RED}Error: Remote 'origin' not found. Ensure the repository is linked to GitHub.${NC}" >&2
  exit 1
fi
echo "Verified: Remote 'origin' exists."

# 3. Get current branch name
current_branch=$(git rev-parse --abbrev-ref HEAD)
if [[ -z "$current_branch" || "$current_branch" == "HEAD" ]]; then
  echo -e "${RED}Error: Could not determine current branch name. Are you in a detached HEAD state?${NC}" >&2
  exit 1
fi
echo "Current branch: ${current_branch}"

# 4. Pull changes (fetch + merge)
echo "Attempting to pull changes from origin/${current_branch}..."
if git pull origin "${current_branch}"; then
  echo -e "${GREEN}Successfully pulled latest changes from origin/${current_branch}.${NC}"
else
  echo -e "${RED}Error: Failed to pull changes. Check for conflicts or connection issues.${NC}" >&2
  # Consider adding conflict resolution instructions here if needed
  exit 1
fi

echo -e "${GREEN}--- Sync Start Complete ---${NC}"
exit 0