# Build stage
FROM node:22-alpine AS builder

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install all dependencies (including devDependencies for build)
RUN npm ci

# Copy source code
COPY tsconfig.json ./
COPY src/ ./src/

# Build TypeScript
RUN npm run build

# Production stage
FROM node:22-alpine

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install production dependencies only, then drop npm and yarn.
# The server only runs node. Those tools ship their own copies of tar and other packages.
RUN npm ci --omit=dev && npm cache clean --force \
 && rm -rf /usr/local/lib/node_modules/npm \
           /usr/local/bin/npm \
           /usr/local/bin/npx \
           /usr/local/lib/node_modules/corepack \
           /usr/local/bin/corepack \
           /opt/yarn-v* \
           /usr/local/bin/yarn \
           /usr/local/bin/yarnpkg

# Copy built code from builder
COPY --from=builder /app/dist ./dist/

# Non-root user for security
RUN addgroup -g 1001 -S nodejs && \
    adduser -S mcp -u 1001
USER mcp

# Environment
ENV NODE_ENV=production
ENV MS365_MCP_PORT=3000
ENV MS365_MCP_HOST=0.0.0.0

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD wget -q --spider http://localhost:3000/health || exit 1

CMD ["node", "dist/index.js"]
