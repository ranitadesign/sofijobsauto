FROM node:20-bookworm

RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  python3 \
  python3-pip \
  fonts-dejavu \
  fonts-liberation \
  fonts-noto \
  fonts-noto-cjk \
  fonts-noto-color-emoji \
  fontconfig \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app
# Dependencies (incl. archiver) – change this line to force cache invalidation if needed
COPY package.json package-lock.json ./
RUN npm ci --omit=dev \
  && node -e "require('archiver')" || (echo "FATAL: archiver not installed" && exit 1)

# Postproceso PPTX (python-pptx). En Render setear PYTHON_BIN=python3
COPY requirements.txt ./
RUN pip3 install --no-cache-dir -r requirements.txt

COPY . .

# (Opcional) si tenés fuentes .ttf propias:
# COPY fonts/ /usr/local/share/fonts/
# RUN fc-cache -fv

ENV PORT=3000
EXPOSE 3000
CMD ["node", "server.js"]
