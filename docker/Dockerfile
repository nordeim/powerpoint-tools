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
#RUN curl -LsSf https://astral.sh/uv/install.sh | sh
# Ensure uv is available in the image before we call it
# Copy prebuilt uv binary into /app/bin and make executable
COPY bin/uv /app/bin/uv
# optionally copy uvx or other helper binaries
COPY bin/uvx /app/bin/uvx
RUN chmod +x /app/bin/uv /app/bin/uvx
ENV PATH="/app/bin:/root/.cargo/bin:$PATH"

# Copy requirements first for caching
COPY requirements.txt .

# Install Python dependencies via uv
# --system installs into the system python environment, reducing complexity for the container
RUN uv pip install --system -r requirements.txt

# Copy the rest of the application
# (This is a fallback; usually we mount the volume at runtime)
COPY . .

# Set PYTHONPATH so tools can import 'core' without path hacking issues
#ENV PYTHONPATH="${PYTHONPATH}:/app"
# Add above ENV line (preferably before COPY . .)
ARG PYTHONPATH=""
ENV PYTHONPATH="${PYTHONPATH}:/app"

# Default command is a shell, ready for the Agent
CMD ["/bin/bash"]
