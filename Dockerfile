@'
FROM node:20-slim

WORKDIR /app

# Copiar dependencias
COPY package.json package-lock.json ./

# Instalar dependencias (usar npm install en lugar de npm ci)
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