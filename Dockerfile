# Use Apify's Playwright base image (includes all browser dependencies)
FROM apify/actor-node-playwright-chrome:18

# Copy package files and install dependencies
COPY package*.json ./
RUN npm --quiet set progress=false \
    && npm install --omit=dev --omit=optional \
    && echo "NPM install succeeded"

# Copy source
COPY . ./

# Run the actor
CMD ["node", "main.js"]
