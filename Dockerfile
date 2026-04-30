# Etapa de construcción (Build)
FROM node:18-alpine AS build

# Establecer directorio de trabajo
WORKDIR /app

# Copiar archivos de dependencias desde la subcarpeta pdf_generator
COPY pdf_generator/package*.json ./

# Instalar dependencias
RUN npm install

# Copiar el resto del código fuente
COPY pdf_generator/ ./

# Construir la aplicación para producción
RUN npm run build

# Etapa de producción (Servir con Nginx)
FROM nginx:alpine

# Copiar los archivos construidos de Vite al directorio público de Nginx
COPY --from=build /app/dist /usr/share/nginx/html

# Exponer el puerto 80 (puerto por defecto para Nginx)
EXPOSE 80

# Iniciar Nginx
CMD ["nginx", "-g", "daemon off;"]
