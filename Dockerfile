FROM node:20-slim

RUN apt-get update && apt-get install -y python3 make g++ unzip && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json ./

RUN npm install --omit=dev

COPY . .

# Descomprimir la base de datos
RUN unzip -o comparacion.db.zip -d /app/ && rm comparacion.db.zip

# Crear carpetas necesarias
RUN mkdir -p data bases reportes public

EXPOSE 3000

CMD ["node", "servidor.js"]