# Usa una imagen base de Node.js oficial
FROM node:20-alpine

# Establece el directorio de trabajo dentro del contenedor
WORKDIR /app

# Copia el archivo package.json y package-lock.json al directorio de trabajo
COPY package*.json ./

# Instala las dependencias del proyecto
RUN npm ci --omit=dev || npm install

# Copia todo el código de la aplicación al directorio de trabajo
COPY . .

# Expone el puerto en el que la aplicación escuchará
EXPOSE 5000

# Define el comando que se ejecutará cuando se inicie el contenedor
CMD ["npm", "run", "start"]
