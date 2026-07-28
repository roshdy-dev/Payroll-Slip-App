#!/bin/bash
# ============================================================================
#  Payroll Slip Generator - Publish Script (macOS)
#  Creates a self-contained, single-file executable that runs without .NET installed.
# ============================================================================

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║   Payroll Slip Generator - Publish (macOS)       ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""

PROJECT_DIR="src/PayrollSlipApp"
OUTPUT_DIR="publish/osx-x64"

echo "[1/2] Restoring packages..."
dotnet restore "$PROJECT_DIR/PayrollSlipApp.csproj"
if [ $? -ne 0 ]; then
    echo "ERROR: Restore failed."
    exit 1
fi

echo ""
echo "[2/2] Publishing self-contained single-file executable..."
echo "      This may take a few minutes on first run..."
echo ""

dotnet publish "$PROJECT_DIR/PayrollSlipApp.csproj" \
    -c Release \
    -r osx-x64 \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:PublishTrimmed=true \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -o "$OUTPUT_DIR"

if [ $? -ne 0 ]; then
    echo "ERROR: Publish failed."
    exit 1
fi

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║   PUBLISH SUCCESSFUL                             ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""
echo "   Output: $OUTPUT_DIR/PayrollSlipGenerator"
echo ""
echo "   This file can be distributed and run on ANY Mac"
echo "   machine WITHOUT .NET installed!"
echo ""

chmod +x "$OUTPUT_DIR/PayrollSlipGenerator"
