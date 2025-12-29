FROM node:20-bookworm

RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  fonts-noto \
  fonts-noto-cjk \
  fonts-noto-color-emoji \
  fontconfig \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev

COPY . .

# (Opcional) si tenés fuentes .ttf propias:
# COPY fonts/ /usr/local/share/fonts/
# RUN fc-cache -fv

ENV PORT=3000
EXPOSE 3000
CMD ["node", "server.js"]
