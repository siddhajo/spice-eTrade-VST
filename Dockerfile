# Spice Config — Railway Dockerfile (sql.js variant)
# No native compilation required — sql.js is pure JavaScript with WASM.
# Much simpler and more reliable than the better-sqlite3 path.

FROM node:20-bookworm-slim

WORKDIR /app

# Copy package.json (no package-lock — Railway gets a clean install)
COPY package.json ./

# Install only runtime deps. better-sqlite3 is in optionalDependencies
# and will be skipped if compilation fails (or via the flag below).
ENV NPM_CONFIG_OPTIONAL=false
RUN npm install --omit=dev

# Verify sql.js loads
RUN node -e "require('sql.js')().then(SQL => { new SQL.Database(); console.log('sql.js OK'); });"

# Copy app source (after deps so Docker caches deps separately)
COPY . .

# Strip build artifacts
RUN rm -rf data/ electron/ *.zip BUILD.md RELEASE.md MIGRATION.md \
    && rm -f recover-isp.js

EXPOSE 3001

CMD ["node", "server.js"]
