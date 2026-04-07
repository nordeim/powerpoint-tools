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

