# NanoClaw Agent Container
# Runs Claude Agent SDK in isolated Linux VM with browser automation

FROM node:22-slim

# Install system dependencies for Chromium and PDF/OCR processing
RUN apt-get update && apt-get install -y \
    chromium \
    fonts-liberation \
    fonts-noto-cjk \
    fonts-noto-color-emoji \
    libgbm1 \
    libnss3 \
    libatk-bridge2.0-0 \
    libgtk-3-0 \
    libx11-xcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    libasound2 \
    libpangocairo-1.0-0 \
    libcups2 \
    libdrm2 \
    libxshmfence1 \
    curl \
    git \
    python3.11-venv \
    poppler-utils \
    tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

# Set Chromium path for agent-browser
ENV AGENT_BROWSER_EXECUTABLE_PATH=/usr/bin/chromium
ENV PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH=/usr/bin/chromium

# Install agent-browser and claude-code globally
RUN npm install -g agent-browser @anthropic-ai/claude-code

# Create app directory
WORKDIR /app

# Copy package files first for better caching
COPY agent-runner/package*.json ./

# Install dependencies
RUN npm install

# Copy source code
COPY agent-runner/ ./

# Build TypeScript
RUN npm run build

# Clone PubMed MCP server for literature search (requirements installed after venv creation)
RUN git clone https://github.com/JackKuo666/PubMed-MCP-Server.git /app/mcp-servers/pubmed-mcp

# Create workspace directories
RUN mkdir -p /workspace/group /workspace/global /workspace/extra /workspace/ipc/messages /workspace/ipc/tasks /workspace/ipc/input

# Create entrypoint script
# Secrets are passed via stdin JSON — temp file is deleted immediately after Node reads it
# Follow-up messages arrive via IPC files in /workspace/ipc/input/
RUN printf '#!/bin/bash\nset -e\ncd /app && npx tsc --outDir /tmp/dist 2>&1 >&2\nln -s /app/node_modules /tmp/dist/node_modules\nchmod -R a-w /tmp/dist\ncat > /tmp/input.json\nnode /tmp/dist/index.js < /tmp/input.json\n' > /app/entrypoint.sh && chmod +x /app/entrypoint.sh

# Set ownership to node user (non-root) for writable directories
RUN chown -R node:node /workspace && chmod 777 /home/node

# Switch to non-root user (required for --dangerously-skip-permissions)
USER node

# Install Python packages for the node user
# Create venv, install python-pptx, pdfplumber, openpyxl, pytesseract
# Also install PubMed MCP server requirements
RUN mkdir -p /home/node/bin \
    && python3 -m venv /home/node/.venv \
    && . /home/node/.venv/bin/activate \
    && pip install --no-cache-dir \
        python-pptx \
        pdfplumber \
        PyPDF2 \
        openpyxl \
        pytesseract \
        matplotlib \
        Pillow \
    && pip install --no-cache-dir -r /app/mcp-servers/pubmed-mcp/requirements.txt \
    && ln -sf /home/node/.venv/bin/pip /home/node/bin/pip \
    && ln -sf /home/node/.venv/bin/pip /home/node/bin/pip3

# Create wrapper scripts that use venv python
RUN printf '#!/bin/bash\nexec /home/node/.venv/bin/python "$@"' > /home/node/bin/python \
    && printf '#!/bin/bash\nexec /home/node/.venv/bin/python "$@"' > /home/node/bin/python3 \
    && chmod +x /home/node/bin/python /home/node/bin/python3

# Add /home/node/bin to PATH for all users
ENV PATH="/home/node/bin:${PATH}"

# Set working directory to group workspace
WORKDIR /workspace/group

# Entry point reads JSON from stdin, outputs JSON to stdout
ENTRYPOINT ["/app/entrypoint.sh"]
