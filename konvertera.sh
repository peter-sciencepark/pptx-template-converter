#!/bin/bash
#
# konvertera.sh — Enkel startfil för att konvertera presentationer
#
# Användning:
#   Dubbelklicka på filen, eller kör i Terminal:
#   ./konvertera.sh min_presentation.pptx
#

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TEMPLATE="$SCRIPT_DIR/mall/Ny mall.potx"
CONVERT="$SCRIPT_DIR/convert.py"

# --- Färger för terminalutskrift ---
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║   PowerPoint-mallkonverterare — Science Park     ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""

# --- Kontrollera Python ---
if command -v python3 &>/dev/null; then
    PYTHON=python3
elif command -v python &>/dev/null; then
    PYTHON=python
else
    echo -e "${RED}Fel: Python hittades inte.${NC}"
    echo "Installera Python 3 från https://www.python.org/downloads/"
    echo ""
    read -p "Tryck Enter för att avsluta..."
    exit 1
fi

# --- Installera beroenden om det behövs ---
if ! $PYTHON -c "import pptx" 2>/dev/null; then
    echo -e "${YELLOW}Installerar nödvändiga paket (första gången)...${NC}"
    $PYTHON -m pip install -r "$SCRIPT_DIR/requirements.txt" --quiet
    if [ $? -ne 0 ]; then
        echo -e "${RED}Kunde inte installera paket. Kör manuellt:${NC}"
        echo "  $PYTHON -m pip install -r $SCRIPT_DIR/requirements.txt"
        read -p "Tryck Enter för att avsluta..."
        exit 1
    fi
    echo -e "${GREEN}Paket installerade!${NC}"
    echo ""
fi

# --- Kontrollera att mallen finns ---
if [ ! -f "$TEMPLATE" ]; then
    echo -e "${RED}Fel: Mallen hittades inte.${NC}"
    echo "Förväntad plats: $TEMPLATE"
    echo "Lägg filen 'Ny mall.potx' i mappen 'mall/'."
    read -p "Tryck Enter för att avsluta..."
    exit 1
fi

# --- Hantera input ---
if [ $# -eq 0 ]; then
    # Inget argument — fråga efter fil
    echo "Ange sökväg till presentationen som ska konverteras."
    echo "Tips: dra och släpp filen hit från Finder."
    echo ""
    read -p "Fil: " INPUT_FILE
    # Ta bort eventuella omgivande citattecken och trailing spaces
    INPUT_FILE=$(echo "$INPUT_FILE" | sed "s/^['\"]//;s/['\"]$//;s/ *$//")
else
    INPUT_FILE="$1"
fi

if [ ! -f "$INPUT_FILE" ]; then
    echo -e "${RED}Fel: Filen hittades inte: $INPUT_FILE${NC}"
    read -p "Tryck Enter för att avsluta..."
    exit 1
fi

# --- Skapa utfilnamn ---
DIR=$(dirname "$INPUT_FILE")
BASENAME=$(basename "$INPUT_FILE" .pptx)
OUTPUT="$DIR/${BASENAME}_ny_mall.pptx"

echo -e "Indatafil:  ${YELLOW}$INPUT_FILE${NC}"
echo -e "Utdatafil:  ${YELLOW}$OUTPUT${NC}"
echo ""

# --- Kör konverteringen ---
$PYTHON "$CONVERT" "$INPUT_FILE" --template "$TEMPLATE" -o "$OUTPUT"

if [ $? -eq 0 ]; then
    echo ""
    echo -e "${GREEN}═══════════════════════════════════════════════════${NC}"
    echo -e "${GREEN}  Klart! Den konverterade filen sparades som:${NC}"
    echo -e "${GREEN}  $OUTPUT${NC}"
    echo -e "${GREEN}═══════════════════════════════════════════════════${NC}"
    echo ""
    echo "Öppna filen i PowerPoint och gör eventuella justeringar."
else
    echo ""
    echo -e "${RED}Något gick fel vid konverteringen.${NC}"
    echo "Kontrollera att indatafilen är en giltig .pptx-fil."
fi

echo ""
read -p "Tryck Enter för att avsluta..."
