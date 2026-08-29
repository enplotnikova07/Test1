FROM node:24-bookworm-slim
WORKDIR /app
COPY package.json pnpm-lock.yaml server.mjs ./
COPY public ./public
COPY scripts ./scripts
RUN corepack enable && pnpm install --frozen-lockfile --prod && mkdir -p /app/.data
EXPOSE 3000
CMD ["node", "server.mjs"]
