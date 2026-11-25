FROM node:20-slim

WORKDIR /app

# Copiar dependencias
COPY package.json package-lock.json ./
RUN npm ci --only=production

# Copiar todo el proyecto
COPY . .

# Crear carpetas necesarias
RUN mkdir -p data bases reportes public

# Puerto
EXPOSE 3000

# Iniciar servidor
CMD ["node", "servidor.js"]