FROM node:20-bullseye

# LibreOffice para convertir PPTX->PDF
RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm ci --only=production

COPY . .

ENV NODE_ENV=production
ENV PORT=3000
# En Linux, el bin suele ser "soffice"
ENV SOFFICE_PATH=/usr/bin/soffice

EXPOSE 3000

CMD ["node", "server.js"]
