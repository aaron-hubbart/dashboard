FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm install --omit=dev
COPY server.js .
ARG APP_VERSION=dev
ENV APP_VERSION=$APP_VERSION
EXPOSE 3001
CMD ["node", "server.js"]