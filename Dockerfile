FROM node:22-alpine
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
COPY server.js ./
ENV NODE_ENV=production PORT=3000
EXPOSE 3000
USER node
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s CMD wget -qO- http://127.0.0.1:3000/health || exit 1
CMD ["node", "server.js"]
