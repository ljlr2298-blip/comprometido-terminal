@'
FROM node:20-slim

# Instalar dependencias de compilación para better-sqlite3
RUN apt-get update && apt-get install -y python3 make g++ && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copiar dependencias
COPY package.json package-lock.json ./

# Instalar dependencias
RUN npm install --omit=dev

# Copiar todo el proyecto
COPY . .

# Crear carpetas necesarias
RUN mkdir -p data bases reportes public

# Puerto
EXPOSE 3000

# Iniciar servidor
CMD ["node", "servidor.js"]
'@ | Set-Content -Path "Dockerfile" -Encoding UTF8