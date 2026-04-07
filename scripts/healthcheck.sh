#!/bin/bash
# Script to run INSIDE the container to verify environment health

echo "🏥 Running Health Checks..."

# 1. Check Python
if command -v python3 &> /dev/null; then
    PY_VER=$(python3 --version)
    echo "✅ Python: $PY_VER"
else
    echo "❌ Python not found"
    exit 1
fi

# 2. Check LibreOffice (headless)
if command -v soffice &> /dev/null; then
    LO_VER=$(soffice --version)
    echo "✅ LibreOffice: $LO_VER"
else
    echo "❌ LibreOffice not found (PDF export will fail)"
    exit 1
fi

# 3. Check Core Library Import
echo "🔍 Checking PowerPoint Agent Core..."
python3 -c "from core.powerpoint_agent_core import PowerPointAgent; print('✅ Core Library Loaded Successfully')"

# 4. Check Write Permissions
echo "📝 Checking Write Permissions..."
touch .healthcheck_tmp
if [ -f .healthcheck_tmp ]; then
    rm .healthcheck_tmp
    echo "✅ Write Access Confirmed"
else
    echo "❌ Cannot write to current directory"
    exit 1
fi

echo "🎉 Environment is HEALTHY and ready for Agent."
