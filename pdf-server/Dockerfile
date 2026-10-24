# PDF化サーバー（LibreOffice入り）。Render では、このフォルダを Docker で動かす
FROM node:20-bookworm-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer libreoffice-calc libreoffice-impress fonts-noto-cjk fonts-ipafont \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY server.js ./
ENV NODE_ENV=production
CMD ["node", "server.js"]
