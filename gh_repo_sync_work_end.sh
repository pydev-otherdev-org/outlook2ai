#!/bin/bash
# gh_repo_sync_work_end.sh
# Stages all changes, accepts a commit message as an argument or prompts for one,
# commits, and pushes to the remote 'origin' for the current branch.

set -o pipefail

# Define colors for output
readonly GREEN='\033[0;32m'
readonly RED='\033[0;31m'
readonly YELLOW='\033[0;33m'
readonly NC='\033[0m' # No Color

echo -e "${YELLOW}--- Ending Sync: Pushing changes to GitHub ---${NC}"

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

	# 3. Check for changes to commit
	if git diff-index --quiet HEAD --; then
		    echo "No changes to commit. Working tree clean."
		        echo -e "${GREEN}--- Sync End Complete (No changes) ---${NC}"
			    exit 0
		    fi
		    echo "Changes detected."

		    # 4. Stage all changes
		    echo "Staging all changes..."
		    if ! git add .; then
			      echo -e "${RED}Error: Failed to stage changes.${NC}" >&2
			        exit 1
			fi
			echo "Changes staged."
			git status --short # Show staged changes briefly

			# 5. Get commit message (either from argument or prompt)
			commit_message="$1"
			if [[ -z "$commit_message" ]]; then
				  while [[ -z "$commit_message" ]]; do
					      read -p "Enter commit message: " commit_message
					          if [[ -z "$commit_message" ]]; then
							        echo -e "${YELLOW}Commit message cannot be empty. Please try again.${NC}"
								    fi
								      done
							      else
								        echo "Using provided commit message: \"$commit_message\""
								fi

								# 6. Commit changes
								echo "Committing changes..."
								if ! git commit -m "$commit_message"; then
									  # Check if commit failed because there was nothing to commit (e.g., only whitespace changes ignored)
									    if git diff-index --quiet HEAD --; then
										          echo "No effective changes were committed (perhaps only whitespace?)."
											        echo -e "${GREEN}--- Sync End Complete (No effective changes) ---${NC}"
												      exit 0
												        else
														      echo -e "${RED}Error: Failed to commit changes.${NC}" >&2
														            exit 1
															      fi
														      fi
														      echo "Changes committed."

														      # 7. Get current branch name
														      current_branch=$(git rev-parse --abbrev-ref HEAD)
														      if [[ -z "$current_branch" || "$current_branch" == "HEAD" ]]; then
															        echo -e "${RED}Error: Could not determine current branch name.${NC}" >&2
																  exit 1
															  fi
															  echo "Current branch: ${current_branch}"

															  # 8. Push changes
															  echo "Attempting to push changes to origin/${current_branch}..."
															  if git push origin "${current_branch}"; then
																    echo -e "${GREEN}Successfully pushed changes to origin/${current_branch}.${NC}"
															    else
																      echo -e "${RED}Error: Failed to push changes. Check connection or if remote has new changes (try running sync_start.sh first).${NC}" >&2
																        exit 1
																fi

																echo -e "${GREEN}--- Sync End Complete ---${NC}"

