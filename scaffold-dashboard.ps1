# ============================================================
# scaffold-dashboard.ps1
# Run from the directory where you want the project created:
#   cd C:\Projects
#   .\scaffold-dashboard.ps1
# ============================================================

$root = "dashboard"

function Write-File($path, $content) {
  $full = Join-Path $root $path
  $dir  = Split-Path $full -Parent
  if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  Set-Content -Path $full -Value $content -Encoding UTF8
  Write-Host "  wrote $path"
}

Write-Host "`nScaffolding TAM Dashboard project...`n"

# ============================================================
# index.html
# ============================================================
Write-File "index.html" @'
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>TAM Dashboard</title>
    <link rel="stylesheet"
      href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/tabler-icons.min.css" />
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.jsx"></script>
  </body>
</html>
'@

# ============================================================
# package.json
# ============================================================
Write-File "package.json" @'
{
  "name": "tam-dashboard",
  "private": true,
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "vite build",
    "preview": "vite preview"
  },
  "dependencies": {
    "react": "^18.3.1",
    "react-dom": "^18.3.1"
  },
  "devDependencies": {
    "@vitejs/plugin-react": "^4.3.1",
    "vite": "^5.4.1"
  }
}
'@

# ============================================================
# vite.config.js
# ============================================================
Write-File "vite.config.js" @'
import { defineConfig } from "vite"
import react from "@vitejs/plugin-react"

export default defineConfig({
  plugins: [react()],
  build: { outDir: "dist" },
  server: {
    proxy: {
      "/api": "http://localhost:3001",
      "/auth": "http://localhost:3001",
    }
  }
})
'@

# ============================================================
# Dockerfile  (frontend)
# ============================================================
Write-File "Dockerfile" @'
FROM node:20-alpine AS build
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
RUN npm run build

FROM nginx:alpine
COPY --from=build /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
'@

# ============================================================
# nginx.conf
# ============================================================
Write-File "nginx.conf" @'
server {
  listen 80;
  root /usr/share/nginx/html;
  index index.html;

  location / {
    try_files $uri $uri/ /index.html;
  }

  gzip on;
  gzip_types text/plain text/css application/javascript application/json;
}
'@

# ============================================================
# cloudbuild.yaml
# ============================================================
Write-File "cloudbuild.yaml" @'
steps:
  - name: gcr.io/cloud-builders/docker
    args: [build, -t, "gcr.io/tam-aaron-hubbart/dashboard-frontend:$COMMIT_SHA",
                  -t, "gcr.io/tam-aaron-hubbart/dashboard-frontend:latest", -f, Dockerfile, .]
  - name: gcr.io/cloud-builders/docker
    args: [push, "gcr.io/tam-aaron-hubbart/dashboard-frontend:$COMMIT_SHA"]
  - name: gcr.io/cloud-builders/docker
    args: [push, "gcr.io/tam-aaron-hubbart/dashboard-frontend:latest"]

  - name: gcr.io/cloud-builders/docker
    args: [build, -t, "gcr.io/tam-aaron-hubbart/dashboard-proxy:$COMMIT_SHA",
                  -t, "gcr.io/tam-aaron-hubbart/dashboard-proxy:latest", -f, proxy/Dockerfile, proxy]
  - name: gcr.io/cloud-builders/docker
    args: [push, "gcr.io/tam-aaron-hubbart/dashboard-proxy:$COMMIT_SHA"]
  - name: gcr.io/cloud-builders/docker
    args: [push, "gcr.io/tam-aaron-hubbart/dashboard-proxy:latest"]

  - name: gcr.io/cloud-builders/kubectl
    args: [set, image, deployment/dashboard,
           "frontend=gcr.io/tam-aaron-hubbart/dashboard-frontend:$COMMIT_SHA",
           "proxy=gcr.io/tam-aaron-hubbart/dashboard-proxy:$COMMIT_SHA",
           -n, tam-dashboard]
    env:
      - CLOUDSDK_COMPUTE_REGION=us-central1
      - CLOUDSDK_CONTAINER_CLUSTER=tam-ah-admin-cluster

images:
  - "gcr.io/tam-aaron-hubbart/dashboard-frontend:$COMMIT_SHA"
  - "gcr.io/tam-aaron-hubbart/dashboard-frontend:latest"
  - "gcr.io/tam-aaron-hubbart/dashboard-proxy:$COMMIT_SHA"
  - "gcr.io/tam-aaron-hubbart/dashboard-proxy:latest"

options:
  logging: CLOUD_LOGGING_ONLY
'@

# ============================================================
# src/main.jsx
# ============================================================
Write-File "src/main.jsx" @'
import React from "react"
import ReactDOM from "react-dom/client"
import App from "./App.jsx"

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
)
'@

# ============================================================
# src/App.jsx  — paste full dashboard artifact code here
# ============================================================
Write-File "src/App.jsx" @'
// !! PASTE THE FULL CONTENTS OF THE "Personal Dashboard v3" ARTIFACT HERE !!
// The artifact is the React component exported as default from App.jsx.
// Make sure to remove the "export default function Dashboard()" wrapper if
// you want to rename it to App, or just keep it as-is and re-export:
//
//   export { Dashboard as default } from "./Dashboard.jsx"
//
// Or paste directly and rename the function to App.
'@

# ============================================================
# proxy/package.json
# ============================================================
Write-File "proxy/package.json" @'
{
  "name": "tam-dashboard-proxy",
  "version": "1.0.0",
  "type": "module",
  "main": "server.js",
  "scripts": {
    "start": "node server.js",
    "dev": "node --watch server.js"
  },
  "dependencies": {
    "cors": "^2.8.5",
    "express": "^4.19.2",
    "express-session": "^1.18.0",
    "node-fetch": "^3.3.2"
  }
}
'@

# ============================================================
# proxy/Dockerfile
# ============================================================
Write-File "proxy/Dockerfile" @'
FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev
COPY server.js .
EXPOSE 3001
CMD ["node", "server.js"]
'@

# ============================================================
# proxy/server.js
# ============================================================
Write-File "proxy/server.js" @'
import express from "express";
import session from "express-session";
import fetch from "node-fetch";
import cors from "cors";

const app = express();
const PORT = 3001;

const {
  ANTHROPIC_API_KEY,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET,
  AZURE_TENANT_ID,
  AZURE_REDIRECT_URI = "https://dashboard.es-sandbox.com/auth/callback",
  SESSION_SECRET = "change-me-in-production",
} = process.env;

app.use(express.json());
app.use(cors({ origin: "https://dashboard.es-sandbox.com", credentials: true }));
app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: { secure: true, httpOnly: true, sameSite: "lax", maxAge: 8 * 60 * 60 * 1000 }
}));

app.get("/auth/login", (req, res) => {
  const params = new URLSearchParams({
    client_id: AZURE_CLIENT_ID,
    response_type: "code",
    redirect_uri: AZURE_REDIRECT_URI,
    scope: "openid profile email Calendars.Read User.Read OnlineMeetings.Read offline_access",
    response_mode: "query",
  });
  res.redirect(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/authorize?${params}`);
});

app.get("/auth/callback", async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.status(400).send(`Auth error: ${error}`);
  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: AZURE_CLIENT_ID,
        client_secret: AZURE_CLIENT_SECRET,
        code,
        redirect_uri: AZURE_REDIRECT_URI,
        grant_type: "authorization_code",
      }),
    }
  );
  const tokens = await tokenRes.json();
  if (tokens.error) return res.status(400).send(`Token error: ${tokens.error_description}`);
  req.session.accessToken = tokens.access_token;
  req.session.refreshToken = tokens.refresh_token;
  req.session.expiresAt = Date.now() + tokens.expires_in * 1000;
  res.redirect("/");
});

app.get("/auth/status", (req, res) => {
  res.json({ authenticated: !!req.session.accessToken });
});

app.get("/auth/logout", (req, res) => {
  req.session.destroy();
  res.redirect(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/logout?post_logout_redirect_uri=https://dashboard.es-sandbox.com`);
});

async function refreshIfNeeded(req) {
  if (!req.session.expiresAt || Date.now() < req.session.expiresAt - 60000) return;
  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: AZURE_CLIENT_ID,
        client_secret: AZURE_CLIENT_SECRET,
        refresh_token: req.session.refreshToken,
        grant_type: "refresh_token",
        scope: "Calendars.Read User.Read OnlineMeetings.Read offline_access",
      }),
    }
  );
  const tokens = await tokenRes.json();
  if (!tokens.error) {
    req.session.accessToken = tokens.access_token;
    req.session.refreshToken = tokens.refresh_token ?? req.session.refreshToken;
    req.session.expiresAt = Date.now() + tokens.expires_in * 1000;
  }
}

function requireAuth(req, res, next) {
  if (!req.session.accessToken) return res.status(401).json({ error: "Not authenticated" });
  next();
}

app.post("/api/claude", requireAuth, async (req, res) => {
  await refreshIfNeeded(req);
  const body = { ...req.body };
  if (Array.isArray(body.mcp_servers)) {
    body.mcp_servers = body.mcp_servers.map(s =>
      s.url?.includes("microsoft365")
        ? { ...s, authorization_token: req.session.accessToken }
        : s
    );
  }
  const upstream = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "mcp-client-2025-04-04",
    },
    body: JSON.stringify(body),
  });
  res.status(upstream.status);
  upstream.headers.forEach((v, k) => {
    if (!["content-encoding", "transfer-encoding", "connection"].includes(k)) res.setHeader(k, v);
  });
  upstream.body.pipe(res);
});

app.listen(PORT, () => console.log(`Proxy listening on :${PORT}`));
'@

# ============================================================
# k8s/namespace.yaml
# ============================================================
Write-File "k8s/namespace.yaml" @'
apiVersion: v1
kind: Namespace
metadata:
  name: tam-dashboard
'@

# ============================================================
# k8s/secret.template.yaml  (reference only — use runbook)
# ============================================================
Write-File "k8s/secret.template.yaml" @'
# DO NOT commit real values. Use the kubectl command in the runbook.
# This file is a reference template only.
apiVersion: v1
kind: Secret
metadata:
  name: dashboard-secrets
  namespace: tam-dashboard
type: Opaque
stringData:
  ANTHROPIC_API_KEY: "REPLACE"
  AZURE_CLIENT_ID: "REPLACE"
  AZURE_CLIENT_SECRET: "REPLACE"
  AZURE_TENANT_ID: "REPLACE"
  SESSION_SECRET: "REPLACE-32-chars"
'@

# ============================================================
# k8s/deployment.yaml
# ============================================================
Write-File "k8s/deployment.yaml" @'
apiVersion: apps/v1
kind: Deployment
metadata:
  name: dashboard
  namespace: tam-dashboard
  labels:
    app: dashboard
spec:
  replicas: 1
  selector:
    matchLabels:
      app: dashboard
  template:
    metadata:
      labels:
        app: dashboard
    spec:
      containers:
        - name: frontend
          image: gcr.io/tam-aaron-hubbart/dashboard-frontend:latest
          ports:
            - containerPort: 80
          resources:
            requests:
              cpu: 25m
              memory: 32Mi
            limits:
              cpu: 100m
              memory: 64Mi
          readinessProbe:
            httpGet:
              path: /
              port: 80
            initialDelaySeconds: 5
            periodSeconds: 10

        - name: proxy
          image: gcr.io/tam-aaron-hubbart/dashboard-proxy:latest
          ports:
            - containerPort: 3001
          env:
            - name: ANTHROPIC_API_KEY
              valueFrom:
                secretKeyRef:
                  name: dashboard-secrets
                  key: ANTHROPIC_API_KEY
            - name: AZURE_CLIENT_ID
              valueFrom:
                secretKeyRef:
                  name: dashboard-secrets
                  key: AZURE_CLIENT_ID
            - name: AZURE_CLIENT_SECRET
              valueFrom:
                secretKeyRef:
                  name: dashboard-secrets
                  key: AZURE_CLIENT_SECRET
            - name: AZURE_TENANT_ID
              valueFrom:
                secretKeyRef:
                  name: dashboard-secrets
                  key: AZURE_TENANT_ID
            - name: SESSION_SECRET
              valueFrom:
                secretKeyRef:
                  name: dashboard-secrets
                  key: SESSION_SECRET
            - name: AZURE_REDIRECT_URI
              value: "https://dashboard.es-sandbox.com/auth/callback"
          resources:
            requests:
              cpu: 50m
              memory: 64Mi
            limits:
              cpu: 200m
              memory: 128Mi
          readinessProbe:
            httpGet:
              path: /auth/status
              port: 3001
            initialDelaySeconds: 5
            periodSeconds: 10
'@

# ============================================================
# k8s/service.yaml
# ============================================================
Write-File "k8s/service.yaml" @'
apiVersion: v1
kind: Service
metadata:
  name: dashboard
  namespace: tam-dashboard
spec:
  selector:
    app: dashboard
  ports:
    - name: frontend
      port: 80
      targetPort: 80
    - name: proxy
      port: 3001
      targetPort: 3001
  type: ClusterIP
'@

# ============================================================
# k8s/ingress.yaml
# ============================================================
Write-File "k8s/ingress.yaml" @'
apiVersion: networking.k8s.io/v1
kind: Ingress
metadata:
  name: dashboard
  namespace: tam-dashboard
  annotations:
    kubernetes.io/ingress.class: nginx
    cert-manager.io/cluster-issuer: letsencrypt-prod
    nginx.ingress.kubernetes.io/ssl-redirect: "true"
    nginx.ingress.kubernetes.io/configuration-snippet: |
      location ~* ^/(api|auth)(/|$) {
        proxy_pass http://dashboard.tam-dashboard.svc.cluster.local:3001;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
      }
spec:
  ingressClassName: nginx
  tls:
    - hosts:
        - dashboard.es-sandbox.com
      secretName: dashboard-tls
  rules:
    - host: dashboard.es-sandbox.com
      http:
        paths:
          - path: /
            pathType: Prefix
            backend:
              service:
                name: dashboard
                port:
                  number: 80
'@

# ============================================================
# .gitignore
# ============================================================
Write-File ".gitignore" @'
node_modules/
dist/
proxy/node_modules/
k8s/secret.yaml
*.env
.env*
'@

Write-Host "`nDone! Project scaffolded at .\$root`n"
Write-Host "Next steps:"
Write-Host "  1. Paste the dashboard React code into src\App.jsx"
Write-Host "  2. Complete the Azure App Registration (see azure-app-reg artifact)"
Write-Host "  3. Follow the runbook (deploy-runbook artifact) to deploy"
Write-Host ""
