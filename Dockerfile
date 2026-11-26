FROM node:22-slim

WORKDIR /app

COPY package.json ./

RUN npm install --omit=dev

COPY . .

RUN mkdir -p data bases reportes public uploads

EXPOSE 3000

CMD ["node", "servidor.js"]