#!/bin/bash
# deploy.sh — Deploy PDG Translate library and update all stubs
# Usage: ./deploy.sh

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# Script ID of the main library project (from .clasp.json)
SCRIPT_ID=$(jq -r '.scriptId' .clasp.json)

# Stub Script IDs to update (add yours here)
STUB_SCRIPT_IDS=(
  # "1xavof4et8zzdz3tdbnPmS0_fNGJ-L9pJCATg39e0wlYgVKCiBE6DZ9Vq"
  # "1xbHo2zpxLmXVAST-3fupf4b0gs4Qdy7aQmkZZS-4n-9JxGYflcP6G92N"
  # Add your 5 stub script IDs here
)

# ── Step 1: Push files to library ─────────────────────────────────────────────
echo ""
echo -e "${CYAN}── Step 1: Pushing files to library ─────────────────────────────${NC}"
clasp push

if [ $? -ne 0 ]; then
    echo -e "${RED}Failed to push files${NC}"
    exit 1
fi

# ── Step 2: Get version number ────────────────────────────────────────────────
echo ""
echo -e "${CYAN}── Step 2: Publishing new library version ────────────────────────${NC}"

# Create new deployment
DEPLOY_OUTPUT=$(clasp deploy -d "Library deployment" 2>&1) || true

# Extract version number
VERSION=$(echo "$DEPLOY_OUTPUT" | grep -oP '@\K[0-9]+' | tail -1)

if [ -z "$VERSION" ]; then
    echo -e "${YELLOW}Warning: Could not extract version number${NC}"
    echo "Creating new deployment..."
    clasp deploy -d "Library deployment"
    exit 0
fi

echo -e "${GREEN}Published as version $VERSION${NC}"

# ── Step 3: Update stubs ─────────────────────────────────────────────────────
echo ""
echo -e "${CYAN}── Step 3: Updating stubs to library v$VERSION ───────────────────${NC}"

# Get access token from clasp
ACCESS_TOKEN=$(cat ~/.clasprc.json 2>/dev/null | jq -r '.tokens.default.access_token' 2>/dev/null)

if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
    echo -e "${RED}Not logged in. Run: clasp login${NC}"
    exit 1
fi

SUCCESS_COUNT=0
API_BASE="https://script.googleapis.com/v1/projects"

for stub_id in "${STUB_SCRIPT_IDS[@]}"; do
    echo -n "  $stub_id ..."
    
    # Get current content
    CONTENT=$(curl -s -H "Authorization: Bearer $ACCESS_TOKEN" \
        "$API_BASE/$stub_id/content" 2>/dev/null)
    
    # Check if we got valid content
    if [ -z "$CONTENT" ] || [[ "$CONTENT" == *"error"* ]]; then
        echo -e " ${YELLOW}failed to fetch, skipping${NC}"
        continue
    fi
    
    # Find and update manifest
    UPDATED=$(echo "$CONTENT" | jq --arg ver "$VERSION" '
        .files = (.files | map(
            if .name == "appsscript" then
                .source = (.source | fromjson | 
                    .dependencies.libraries = (.dependencies.libraries // []) |
                    .dependencies.libraries = (.dependencies.libraries | map(
                        if .userSymbol == "PDGLib" then
                            .version = ($ver | tonumber)
                        else . end
                    )) |
                    tojson
                )
            else . end
        ))
    ')
    
    # Push updated content back
    RESULT=$(echo "$UPDATED" | curl -s -X PUT \
        -H "Authorization: Bearer $ACCESS_TOKEN" \
        -H "Content-Type: application/json" \
        -d @- \
        "$API_BASE/$stub_id/content" 2>/dev/null)
    
    if [[ "$RESULT" == *"error"* ]] || [ -z "$RESULT" ]; then
        echo -e " ${RED}failed${NC}"
    else
        echo -e " ${GREEN}v$VERSION${NC}"
        ((SUCCESS_COUNT++))
    fi
done

echo ""
echo "═══════════════════════════════════════════════════════════════════"
echo -e "${GREEN}Deployed library v$VERSION — $SUCCESS_COUNT / ${#STUB_SCRIPT_IDS[@]} stubs updated.${NC}"
echo "═══════════════════════════════════════════════════════════════════"
