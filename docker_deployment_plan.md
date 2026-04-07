# Comprehensive plan to containerize the PowerPoint Agent Tools

This ensures a consistent, reproducible environment across Windows, macOS, and Linux, solving the complex dependency chain of `libreoffice` (for PDF/Image export) and `fonts` which often plague local Python setups.

## 1. Architecture
*   **Base Image:** `python:3.13-trixie` (Stable and includes modern LibreOffice repositories).
*   **System Dependencies:** `libreoffice-impress` (headless), `fonts-liberation` (for rendering), `uv` (for fast Python management).
*   **Volume Strategy:** The host repository is mounted to `/app` inside the container. This allows the AI agent to modify code and generate presentations that persist on the host machine.
*   **Permissions:** The container will execute as `root` internally to simplify operations, but file ownership on Linux will be managed via standard Docker volume behavior (or simple `chown` if needed in specific environments).

## 2. Docker Configuration

-  **`docker/Dockerfile`**: Multi-stage build definition installing OS deps, Fonts, and Python environment.
-  **`docker/.dockerignore`**: Prevents temp files (`__pycache__`, `.git`) from bloating the build context.

## 3. Control Scripts

-  **`scripts/build_docker.sh`** (Linux/Mac) & **`.bat`** (Windows): Compiles the image `ppt-agent-tools:latest`.
-  **`scripts/run_docker.sh`** (Linux/Mac) & **`.bat`** (Windows): Launches an interactive shell (`/bin/bash`) with the current directory mounted.
-  **`scripts/healthcheck.sh`**: Internal script to verify `python`, `pptx`, and `soffice` (LibreOffice) are operational.

## 4. Documentation

-  **`DOCKER_GUIDE.md`**: The "Handbook" explaining how to spin up the environment and interact with it.

---

### 1. `docker/Dockerfile`

```dockerfile
# Use official Python slim image based on Debian Bookworm
# Bookworm is required for recent LibreOffice versions
FROM python:3.13-trixie

# Prevent Python from writing pyc files and buffering stdout
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set the working directory
WORKDIR /app

# Install System Dependencies
# - curl: for installing uv
# - libreoffice-impress: for PDF/Image export
# - fonts-liberation: standard fonts to prevent layout shifting
# - procps: for process monitoring
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    libreoffice-impress \
    fonts-liberation \
    procps \
    && rm -rf /var/lib/apt/lists/*

# Install uv for fast package management
RUN curl -LsSf https://astral.sh/uv/install.sh | sh
ENV PATH="/root/.cargo/bin:$PATH"

# Copy requirements first for caching
COPY requirements.txt .

# Install Python dependencies via uv
# --system installs into the system python environment, reducing complexity for the container
RUN uv pip install --system -r requirements.txt

# Copy the rest of the application
# (This is a fallback; usually we mount the volume at runtime)
COPY . .

# Set PYTHONPATH so tools can import 'core' without path hacking issues
ENV PYTHONPATH="${PYTHONPATH}:/app"

# Default command is a shell, ready for the Agent
CMD ["/bin/bash"]
```

### 2. `docker/.dockerignore`

```text
# Version Control
.git
.gitignore

# Python
__pycache__
*.pyc
*.pyo
*.pyd
.Python
env/
venv/
.env
.venv

# PowerPoint Files (Don't bake generated files into the image)
*.pptx
*.pdf
*.png
*.jpg

# Docker
docker/
scripts/
```

### 3. `scripts/build_docker.sh` (Linux/Mac)

```bash
#!/bin/bash
set -e

# Get the root directory of the project
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )/.." && pwd )"

echo "🐳 Building PowerPoint Agent Tools Docker Image..."
docker build -t ppt-agent-tools:latest -f "$DIR/docker/Dockerfile" "$DIR"

echo "✅ Build Complete. Image: ppt-agent-tools:latest"
```

### 4. `scripts/build_docker.bat` (Windows)

```batch
@echo off
echo 🐳 Building PowerPoint Agent Tools Docker Image...

:: Get parent directory
pushd %~dp0..
docker build -t ppt-agent-tools:latest -f docker\Dockerfile .
popd

echo ✅ Build Complete. Image: ppt-agent-tools:latest
pause
```

### 5. `scripts/run_docker.sh` (Linux/Mac)

```bash
#!/bin/bash
set -e

# Get the root directory of the project
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )/.." && pwd )"

echo "🚀 Starting PowerPoint Agent Container..."
echo "📂 Mounting: $DIR -> /app"

# Run interactive container
# --rm: Remove container after exit
# -it: Interactive TTY
# -v: Mount project root to /app
docker run --rm -it \
    -v "$DIR:/app" \
    ppt-agent-tools:latest
```

### 6. `scripts/run_docker.bat` (Windows)

```batch
@echo off
echo 🚀 Starting PowerPoint Agent Container...

:: Get parent directory for mounting
pushd %~dp0..
set HOST_DIR=%CD%
popd

echo 📂 Mounting: %HOST_DIR% -^> /app

:: Run interactive container
docker run --rm -it -v "%HOST_DIR%:/app" ppt-agent-tools:latest
```

### 7. `scripts/healthcheck.sh`

```bash
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
```

### 8. `DOCKER_GUIDE.md`

```markdown
# 🐳 Docker Environment Handbook

This guide explains how to use the containerized environment for **PowerPoint Agent Tools**. This environment is the "Gold Standard" for running the agent, ensuring all dependencies (especially LibreOffice and Fonts) are correctly configured on any OS.

---

## 1. Why Docker?

*   **PDF/Image Export:** Requires `LibreOffice` which is heavy and hard to configure on Lambda/Cloud functions. The container has it pre-installed.
*   **Fonts:** PowerPoint layouts break if fonts are missing. The container includes `fonts-liberation` for stability.
*   **Consistency:** Guarantees the code runs exactly the same on Windows, Mac, and Linux.

---

## 2. Quick Start

### Step 1: Build the Image
Run the build script once to create the `ppt-agent-tools` image.

**Mac/Linux:**
```bash
./scripts/build_docker.sh
```

**Windows:**
```cmd
scripts\build_docker.bat
```

### Step 2: Enter the Environment
This drops you (or the Agent) into a `bash` shell inside the container. Your local files are mounted to `/app`, so any files you generate will appear on your host machine immediately.

**Mac/Linux:**
```bash
./scripts/run_docker.sh
```

**Windows:**
```cmd
scripts\run_docker.bat
```

---

## 3. Verifying the Setup

Once inside the container, run the health check to ensure all systems are go.

```bash
root@container:/app# ./scripts/healthcheck.sh
```

**Expected Output:**
```text
🏥 Running Health Checks...
✅ Python: Python 3.11.x
✅ LibreOffice: LibreOffice 7.x
✅ Core Library Loaded Successfully
✅ Write Access Confirmed
🎉 Environment is HEALTHY and ready for Agent.
```

---

## 4. Agent Workflow

When an AI Agent is using this environment, it should perform all operations inside the interactive shell session.

**Example Session:**

```bash
# 1. Create a deck
uv python tools/ppt_create_new.py --output my_deck.pptx --json

# 2. Add a chart (uses internal python-pptx)
uv python tools/ppt_add_chart.py --file my_deck.pptx ... --json

# 3. Export to PDF (uses internal LibreOffice)
uv python tools/ppt_export_pdf.py --file my_deck.pptx --output my_deck.pdf --json
```

---

## 5. Troubleshooting

### **Permissions Issues (Linux)**
If generated files are owned by `root` and you cannot edit them on your host:
1.  This is expected as the container runs as root.
2.  Fix ownership on host: `sudo chown -R $USER:$USER .`
3.  *Advanced:* Edit `Dockerfile` to create a user matching your host UID.

### **PDF Export Fails**
If `ppt_export_pdf.py` fails inside Docker:
1.  Ensure `libreoffice-impress` is installed (run healthcheck).
2.  Ensure the container has enough memory (Docker Desktop default is sometimes too low). Increase to 4GB if processing large decks.

### **"Module not found"**
If `core` cannot be imported:
1.  The `Dockerfile` sets `ENV PYTHONPATH="${PYTHONPATH}:/app"`.
2.  Ensure you are running scripts from the `/app` directory.
```

---

https://aistudio.google.com/app/prompts?state=%7B%22ids%22:%5B%2214JUfPU8LM7JCTGdhVdXcBtn23M6s1hd1%22%5D,%22action%22:%22open%22,%22userId%22:%22108686197475781557359%22,%22resourceKeys%22:%7B%7D%7D&usp=sharing, https://drive.google.com/file/d/17-v8tsynq-IOaTWVjTjKaElNVuJhvx-A/view?usp=sharing, https://drive.google.com/file/d/1GZYAmIfN1SVokiYIFVYaL0Q2wc_eQDxO/view?usp=sharing, https://drive.google.com/file/d/1RrrW5UxvpEvUufytMunLUCY6eFbGd-4b/view?usp=sharing, https://drive.google.com/file/d/1luvKEdUJzqEhmqTUfRjRbRIftuPL6y7N/view?usp=sharing, https://drive.google.com/file/d/1vJp24mGfmsCPvdVmfakuDCgsrsQr81w7/view?usp=sharing
